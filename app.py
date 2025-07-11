import streamlit as st
import os
import PyPDF2
from openai import OpenAI
import pandas as pd
from io import BytesIO
import re
import bcrypt
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement # Crucial for horizontal line fix
from datetime import datetime # For timestamping documents

# --- NEW/UPDATED IMPORTS FOR FIREBASE ---
import firebase_admin
from firebase_admin import credentials, auth, firestore, storage # Added 'storage'
import json
import requests # Added for direct Firebase Auth REST API calls

# --- SESSION STATE INITIALIZATION (MOVED TO TOP) ---
# Initialize session state variables to store data across reruns
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_name' not in st.session_state:
    st.session_state['user_name'] = ''
if 'ai_review' not in st.session_state:
    st.session_state['ai_review'] = '' # This will now store the SINGLE combined AI review
if 'processed_ai_review_docx_buffer' not in st.session_state:
    st.session_state['processed_ai_review_docx_buffer'] = None # This will store the SINGLE combined DOCX buffer
if 'review_triggered' not in st.session_state:
    st.session_state['review_triggered'] = False
if 'uploaded_file_names' not in st.session_state:
    st.session_state['uploaded_file_names'] = [] # Changed to list for potential multiple uploads, though current logic uses single.
# --- NEW: Add user_uid and user_role to session state ---
if 'user_uid' not in st.session_state:
    st.session_state['user_uid'] = None
if 'user_role' not in st.session_state:
    st.session_state['user_role'] = None # 'admin' or 'user'
# --- NEW: Add state for login type selection ---
if 'selected_login_type' not in st.session_state:
    st.session_state['selected_login_type'] = None # Can be 'admin', 'user', or None
# --- NEW: State for invite user form details after submission ---
if 'invite_details_to_display' not in st.session_state:
    st.session_state['invite_details_to_display'] = None # Stores {'email': ..., 'link': ..., 'role': ...}
# --- NEW: State for mandatory name setup flag ---
if 'name_needs_setup' not in st.session_state:
    st.session_state['name_needs_setup'] = False
# --- NEW: Flags for invite admin confirmation flow ---
if 'show_admin_invite_confirm' not in st.session_state:
    st.session_state['show_admin_invite_confirm'] = False
if 'temp_invited_email' not in st.session_state: # Temp store email during confirmation phase
    st.session_state['temp_invited_email'] = ""
if 'temp_invite_role' not in st.session_state: # Temp store role during confirmation phase
    st.session_state['temp_invite_role'] = ""
# --- NEW: Add state for storing the currently uploaded file objects ---
if 'current_uploaded_files' not in st.session_state:
    st.session_state['current_uploaded_files'] = [] # List to hold multiple files
# --- NEW: State for active admin tab (to control data refresh) ---
if 'active_admin_tab' not in st.session_state:
    st.session_state['active_admin_tab'] = "Invite New User" # Default active tab
# --- NEW: To track last rendered admin tab for clearing purposes ---
if 'last_rendered_admin_tab' not in st.session_state:
    st.session_state['last_rendered_admin_tab'] = "Invite New User"
# --- NEW: Flag to trigger a rerun for admin tabs ---
if 'trigger_rerun_for_admin_tab' not in st.session_state:
    st.session_state['trigger_rerun_for_admin_tab'] = False


# --- FIREBASE INITIALIZATION ---
# Initialize Firebase only once per session to prevent errors on rerun
if 'firebase_initialized' not in st.session_state:
    try:
        # Load Firebase Admin SDK config from secrets (now individual keys)
        # We assemble the dictionary required by credentials.Certificate
        service_account_info = {
            "type": st.secrets["firebase_admin_sdk"]["type"],
            "project_id": st.secrets["firebase_admin_sdk"]["project_id"],
            "private_key_id": st.secrets["firebase_admin_sdk"]["private_key_id"],
            "private_key": st.secrets["firebase_admin_sdk"]["private_key"].replace('\\n', '\n'), # Ensure newlines are correctly interpreted
            "client_email": st.secrets["firebase_admin_sdk"]["client_email"],
            "client_id": st.secrets["firebase_admin_sdk"]["client_id"],
            "auth_uri": st.secrets["firebase_admin_sdk"]["auth_uri"],
            "token_uri": st.secrets["firebase_admin_sdk"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["firebase_admin_sdk"]["client_x509_cert_url"],
            "universe_domain": st.secrets["firebase_admin_sdk"]["universe_domain"],
        }
        cred = credentials.Certificate(service_account_info)
        
        # Initialize firebase_admin app. If an app is already initialized, use get_app()
        if not firebase_admin._apps: # Check if no Firebase app has been initialized yet
            firebase_admin.initialize_app(cred, {
                'storageBucket': st.secrets["firebase"]["storageBucket"] # Provide storage bucket for client
            })
        else:
            firebase_admin.get_app() # Get the already initialized app
        
        # Store Firestore client for admin tasks and secure data access
        st.session_state['firestore_db'] = firestore.client()
        # Store Firebase Storage client
        st.session_state['firebase_storage_client'] = storage.bucket() # Get the default bucket
        
        # Store Firebase API Key for direct REST API calls for auth
        st.session_state['firebase_api_key'] = st.secrets["firebase"]["apiKey"]
        
        # Mark Firebase as initialized
        st.session_state['firebase_initialized'] = True
        st.success("Firebase Initialized Successfully!")
    except Exception as e:
        st.error(f"Error initializing Firebase: {e}. Please ensure your `.streamlit/secrets.toml` "
                 "file has the correct Firebase configurations (both `[firebase]` and `[firebase_admin_sdk]`).")
        st.stop() # Stop the app if Firebase cannot be initialized

# --- OPENAI CLIENT INITIALIZATION ---
try:
    openai_client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except KeyError:
    st.error("OpenAI API Key not found in .streamlit/secrets.toml. Please add it.")
    st.stop()

# --- HELPER FUNCTIONS ---
def add_horizontal_line(document):
    """
    Adds a horizontal line (paragraph border) to the document.
    """
    paragraph = document.add_paragraph()
    p_element = paragraph._element
    pPr = p_element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:autoSpaceDNACr', 'w:noKeyEndProp', 'w:lnNumType', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap',
        'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rdPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), "single")
    bottom.set(qn('w:sz'), "12")
    bottom.set(qn('w:space'), "1")
    bottom.set(qn('w:color'), "auto")
    pBdr.append(bottom)

def create_docx_from_text(raw_text: str) -> BytesIO:
    doc = Document()
    try:
        doc.add_picture('logo.png', width=Inches(1.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
    except Exception as e:
        print(f"Could not add SSO logo from 'logo.png': {e}")
        doc.add_paragraph("[SSO Consultants Logo Placeholder]")

    doc.add_heading("AI Ethics Review Report", level=0)

    sections = raw_text.split("---")
    for i, section_content in enumerate(sections):
        section_content = section_content.strip()
        if not section_content:
            continue

        if i > 0:
            add_horizontal_line(doc)
            doc.add_paragraph()

        lines = section_content.split("\n")
        current_section_heading_added = False
        for line_index, line in enumerate(lines):
            stripped_line = line.strip()
            if not stripped_line:
                continue

            if re.match(r"^\d+\.\s", stripped_line) or \
               stripped_line.startswith("**ICMR National Ethical Guidelines for Biomedical and Health Research Involving Human Participants, 2017 Review:**") or \
               stripped_line.startswith("**Medical Devices Rules, 2017 (MDR 2017) Review:**") or \
               stripped_line.startswith("**New Drugs and Clinical Trials Rules, 2019 (NDCTR 2019) Review:**") or \
               stripped_line.startswith("**3. English and Construction of the Questionnaire:**") or \
               stripped_line.startswith("**4. Questionnaire and Informed Consent Alignment:**") or \
               stripped_line.startswith("**5. Additional Aspects:**") or \
               stripped_line.startswith("**6. Summary and Recommendation:**") or \
               stripped_line.startswith("**7. Additional Questions:**"):
                if not current_section_heading_added:
                    if stripped_line.startswith("**"):
                        heading_text = stripped_line.replace('**', '')
                        doc.add_heading(heading_text, level=2)
                    else:
                        doc.add_heading(stripped_line, level=3)
                    current_section_heading_added = True
                    continue
            if stripped_line.startswith("* **"):
                match = re.match(r"^\*\s*\*\*(.*?):\*\*(.*)", stripped_line)
                if match:
                    label = match.group(1).strip()
                    content = match.group(2).strip()
                    p = doc.add_paragraph(style='List Bullet')
                    p.add_run(label + ":").bold = True
                    p.add_run(" " + content)
                else:
                    doc.add_paragraph(stripped_line, style='List Bullet')
            elif stripped_line.startswith("- "):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(stripped_line[2:])
            else:
                doc.add_paragraph(stripped_line)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def upload_and_save_review_document(file_buffer: BytesIO, file_name: str, uploader_uid: str, review_summary: str = ""):
    """
    Uploads the generated DOCX file to Firebase Storage and saves its metadata to Firestore.
    Returns the public download URL or None if an error occurs.
    """
    try:
        bucket = st.session_state['firebase_storage_client']
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_file_name = "".join(x for x in os.path.basename(file_name) if x.isalnum() or x in "._- ").replace(" ", "_")
        storage_path = f"ethics-output-docs/{safe_file_name.replace('.docx', '')}_{timestamp}_{uploader_uid}.docx"
        
        blob = bucket.blob(storage_path)
        file_buffer.seek(0)
        blob.upload_from_file(file_buffer, content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        blob.make_public()
        
        public_url = blob.public_url

        db = st.session_state['firestore_db']
        reviews_collection = db.collection('ethics_review_outputs')
        
        review_data = {
            'outputDocFileName': os.path.basename(file_name),
            'outputDocURL': public_url,
            'uploaderUID': uploader_uid,
            'uploadTimestamp': firestore.SERVER_TIMESTAMP,
            'status': 'Completed',
            'summary': review_summary,
            'review_date_human': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        reviews_collection.add(review_data)
        st.success(f"Generated review saved to cloud storage and database!")
        return public_url
    except Exception as e:
        st.error(f"Error saving review document to Firebase: {e}")
        return None

def extract_summary_from_ai_review(ai_review: str) -> str:
    """
    Extracts the 'Summary and Recommendation' section from the AI review text.
    """
    match = re.search(r"\*\*6\. Summary and Recommendation:\*\*(.*?)(---|\Z)", ai_review, re.DOTALL)
    if match:
        summary = match.group(1).strip()
        if len(summary) > 500:
            return summary[:497] + "..."
        return summary
    return "No summary provided."

def generate_download_buttons(user_name: str, ai_review_text: str, docx_buffer: BytesIO, uploader_uid: str):
    file_name_docx = f"{user_name.replace(' ', '_')}_Combined_ECR_Review.docx"

    buffer_key = f"docx_buffer_combined_{hash(ai_review_text)}"
    if buffer_key not in st.session_state or st.session_state.get(buffer_key) != docx_buffer:
        st.session_state[buffer_key] = docx_buffer
        st.session_state[f"docx_filename_{buffer_key}"] = file_name_docx

    if st.download_button(
        label="⬇ Download Combined AI Ethics Review (DOCX)",
        data=st.session_state[buffer_key],
        file_name=st.session_state[f"docx_filename_{buffer_key}"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"docx_download_trigger_combined_{buffer_key}",
        on_click=lambda: upload_and_save_review_document(
            st.session_state[buffer_key],
            st.session_state[f"docx_filename_{buffer_key}"],
            uploader_uid,
            review_summary=extract_summary_from_ai_review(ai_review_text)
        )
    ):
        pass

def get_ai_review_response(document_text: str) -> str:
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": """You are an AI assistant specialized in reviewing ethics review proposals based on Indian regulations like ICMR, MDR 2017, and NDCTR 2019. 
                You are reviewing a single submission that may consist of multiple combined documents (e.g., application form, research proposal, questionnaire, informed consent, RAC confirmation letter, etc.).
                Your task is to analyze ALL provided text as ONE comprehensive submission and generate a SINGLE, unified ethics review report.
                Provide a structured report with clear sections and recommendations based on the entirety of the input.
                
                Include sections for:
                1. Document Submission Status (assess overall completeness and presence of expected documents if mentioned in the combined text):
                2. Review Against Guidelines (Consolidated assessment across all documents against ICMR 2017, MDR 2017, NDCTR 2019):
                    - ICMR National Ethical Guidelines for Biomedical and Health Research Involving Human Participants, 2017 Review:
                    - Medical Devices Rules, 2017 (MDR 2017) Review:
                    - New Drugs and Clinical Trials Rules, 2019 (NDCTR 2019) Review:
                3. English and Construction of the Questionnaire (if questionnaires are present):
                4. Questionnaire and Informed Consent Alignment (if both are present):
                5. Additional Aspects (any other relevant observations or concerns):
                6. Summary and Recommendation (Overall summary and final recommendation for the entire submission):
                7. Additional Questions (Questions for the user regarding the submission):

                Use '---' to separate main sections. Use bullet points where appropriate within sections.
                Ensure that your review addresses the content from ALL provided documents comprehensively in this single report."""},
                {"role": "user", "content": f"Please perform a comprehensive ethics review of the following combined document content. Treat this as a single submission for review. Focus on compliance with Indian regulations (ICMR, MDR 2017, NDCTR 2019) and evaluate all aspects present in the text, including the questionnaire and informed consent. Provide a detailed, unified report:\n\n{document_text}"}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error calling OpenAI API: {e}")
        return "AI review failed. Please try again later."

# --- AUTHENTICATION FUNCTIONS ---
def authenticate_firebase_user(email, password, expected_role):
    try:
        api_key = st.session_state['firebase_api_key']
        rest_api_url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}"
        payload = json.dumps({"email": email, "password": password, "returnSecureToken": True})
        headers = {"Content-Type": "application/json"}
        response = requests.post(rest_api_url, data=payload, headers=headers)
        response.raise_for_status()
        response_data = response.json()
        id_token = response_data['idToken']
        user_uid = response_data['localId']

        decoded_token = auth.verify_id_token(id_token)
        verified_uid = decoded_token['uid']

        if user_uid != verified_uid:
            st.error("Authentication mismatch. Please try again.")
            return False

        st.session_state['user_uid'] = verified_uid

        user_doc_ref = st.session_state['firestore_db'].collection('User').document(verified_uid)
        user_doc = user_doc_ref.get()

        if user_doc.exists:
            user_data = user_doc.to_dict()
            actual_role = user_data.get('role', 'user')
            first_name = user_data.get('first_name', '')
            last_name = user_data.get('last_name', '')
            
            st.session_state['user_role'] = actual_role
            
            if not first_name or not last_name:
                st.session_state['name_needs_setup'] = True
                st.session_state['user_name'] = email
            else:
                st.session_state['name_needs_setup'] = False
                st.session_state['user_name'] = f"{first_name} {last_name}"
            
            if expected_role == 'admin' and actual_role != 'admin':
                st.error("You attempted to login as Admin, but your account does not have admin privileges. Please use 'Login as User' or contact support.")
                st.session_state['user_uid'] = None
                st.session_state['user_role'] = None
                st.session_state['user_name'] = ''
                st.session_state['name_needs_setup'] = False
                return False
            else:
                return True
        else:
            st.error("User profile not found in Firestore. Please contact support or ensure the user exists in Firestore's 'User' collection with the correct UID.")
            st.session_state['user_uid'] = None
            st.session_state['user_role'] = None
            st.session_state['user_name'] = ''
            st.session_state['name_needs_setup'] = False
            return False

    except requests.exceptions.RequestException as req_e:
        error_message = str(req_e)
        response_status_code = None
        if hasattr(req_e, 'response') and req_e.response is not None:
            response_status_code = req_e.response.status_code
            try:
                error_json = req_e.response.json()
                if "error" in error_json and "message" in error_json["error"]:
                    firebase_error = error_json["error"]["message"]
                    if "EMAIL_NOT_FOUND" in firebase_error or "INVALID_PASSWORD" in firebase_error:
                        st.error("Invalid email or password.")
                    elif "TOO_MANY_ATTEMPTS_TRY_LATER" in firebase_error:
                        st.error("Too many failed login attempts. Please try again later.")
                    else:
                        st.error(f"Authentication error: {firebase_error}")
                    return False
            except json.JSONDecodeError:
                pass
        if response_status_code == 400:
            st.error("Invalid email or password.")
        else:
            st.error(f"Network error: {error_message}. Please check your internet connection or API endpoint.")
        st.session_state['user_uid'] = None
        st.session_state['user_role'] = None
        st.session_state['user_name'] = ''
        st.session_state['name_needs_setup'] = False
        return False
    except Exception as e:
        st.error(f"An unexpected authentication error occurred: {e}")
        st.session_state['user_uid'] = None
        st.session_state['user_role'] = None
        st.session_state['user_name'] = ''
        st.session_state['name_needs_setup'] = False
        return False

def logout():
    """Logs out the user and clears relevant session state."""
    st.session_state['logged_in'] = False
    st.session_state['user_name'] = ''
    st.session_state['ai_review'] = ''
    st.session_state['processed_ai_review_docx_buffer'] = None
    st.session_state['review_triggered'] = False
    st.session_state['uploaded_file_names'] = []
    st.session_state['user_uid'] = None
    st.session_state['user_role'] = None
    st.session_state['selected_login_type'] = None
    st.session_state['invite_details_to_display'] = None
    st.session_state['name_needs_setup'] = False
    st.session_state['show_admin_invite_confirm'] = False
    st.session_state['temp_invited_email'] = ""
    st.session_state['temp_invite_role'] = ""
    st.session_state['current_uploaded_files'] = []
    st.session_state['active_admin_tab'] = "Invite New User"
    st.session_state['last_rendered_admin_tab'] = "Invite New User"
    st.session_state['trigger_rerun_for_admin_tab'] = False
    st.rerun()

# --- SIDEBAR WITH LOGO AND COMPANY INFO (Always visible) ---
with st.sidebar:
    st.image("logo.png", width=120)
    st.markdown("<h3 style='color:#06038D;'>ECR SYSTEM</h3>", unsafe_allow_html=True)
    st.markdown("<span style='color:#FF671F;'>Powered by SSO Consultants</span>", unsafe_allow_html=True)
    if st.session_state['logged_in']:
        st.button("Logout", on_click=logout)

# --- HEADER WITH LOGO BESIDE HEADING (Always visible) ---
col1, col2 = st.columns([8, 1])
with col1:
    st.markdown("<h1 style='color:#06038D; margin-bottom: 0;'>ECR SYSTEM</h1>", unsafe_allow_html=True)
with col2:
    st.image("logo_ecr.png", width=150) # INCREASED WIDTH TO 150

# --- CONDITIONAL UI RENDERING ---
if not st.session_state['logged_in']:
    st.markdown("---")
    st.subheader("Please Login")
    if st.session_state['selected_login_type'] is None:
        col_admin, col_user = st.columns(2)
        with col_admin:
            if st.button("Login as Admin", key="login_admin_btn"):
                st.session_state['selected_login_type'] = 'admin'
                st.rerun()
        with col_user:
            if st.button("Login as User", key="login_user_btn"):
                st.session_state['selected_login_type'] = 'user'
                st.rerun()
    else:
        login_type_display = "Admin" if st.session_state['selected_login_type'] == 'admin' else "User"
        st.markdown(f"---")
        st.subheader(f"Log in as {login_type_display}")
        with st.form(f"login_form_{st.session_state['selected_login_type']}"):
            login_email = st.text_input("Email")
            login_password = st.text_input("Password", type="password")
            login_button = st.form_submit_button("Login")
            
            if login_button:
                if authenticate_firebase_user(login_email, login_password, st.session_state['selected_login_type']):
                    st.session_state['logged_in'] = True
                    st.session_state['selected_login_type'] = None
                    st.rerun()
                    st.stop()
                else:
                    pass
        if st.button("← Back to Choose Login Type", key="back_to_role_selection_after_form"):
            st.session_state['selected_login_type'] = None
            st.rerun()

else: # User is logged in
    st.markdown("---")
    user_welcome_message = f"Welcome, {st.session_state['user_name']}! Your role: **{st.session_state['user_role'].capitalize()}**"
    if st.session_state['user_role'] == 'admin' and st.session_state['user_uid']:
        user_welcome_message += f" (User ID: `{st.session_state['user_uid']}`)"
    st.success(user_welcome_message)

    # --- MANDATORY FIRST-TIME NAME SETUP ---
    if st.session_state['name_needs_setup']:
        st.markdown("---")
        st.warning("Please complete your profile by providing your name.")
        with st.form("name_setup_form"):
            st.write(f"Your email: {st.session_state['user_name']}" if '@' in st.session_state['user_name'] else "")
            
            default_first = ""
            default_last = ""
            if st.session_state['user_name'] and '@' not in st.session_state['user_name']:
                name_parts = st.session_state['user_name'].split(maxsplit=1)
                default_first = name_parts[0]
                if len(name_parts) > 1:
                    default_last = name_parts[1]
            new_first_name = st.text_input("First Name", value=default_first)
            new_last_name = st.text_input("Last Name", value=default_last)
            save_name_button = st.form_submit_button("Save My Name")
            if save_name_button:
                if not new_first_name.strip() or not new_last_name.strip():
                    st.error("First Name and Last Name are required.")
                else:
                    try:
                        user_doc_ref = st.session_state['firestore_db'].collection('User').document(st.session_state['user_uid'])
                        user_doc_ref.update({
                            'first_name': new_first_name.strip(),
                            'last_name': new_last_name.strip()
                        })
                        st.session_state['user_name'] = f"{new_first_name.strip()} {new_last_name.strip()}"
                        st.session_state['name_needs_setup'] = False
                        st.success("Your name has been saved! Redirecting...")
                        st.rerun()
                        st.stop()
                    except Exception as e:
                        st.error(f"Error saving name: {e}")
        st.stop()

    # --- MAIN APP CONTENT FOR ALL LOGGED-IN USERS ---
    st.markdown("---")
    st.markdown("<h3 style='color:#06038D;'>📝 Upload Document(s) for Ethics Review</h3>", unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload PDF document(s) (e.g., Application, Research Proposal, Questionnaire, Informed Consent)",
        type=["pdf"],
        key="pdf_uploader",
        accept_multiple_files=True
    )

    if uploaded_files:
        current_session_file_names = [f.name for f in st.session_state['current_uploaded_files']]
        new_uploaded_file_names = [f.name for f in uploaded_files]

        if set(current_session_file_names) != set(new_uploaded_file_names):
            st.session_state['current_uploaded_files'] = uploaded_files
            st.session_state['ai_review'] = ''
            st.session_state['processed_ai_review_docx_buffer'] = None
            st.session_state['review_triggered'] = False

        if st.session_state['current_uploaded_files']:
            st.write("Selected files:")
            for file in st.session_state['current_uploaded_files']:
                st.write(f"- {file.name}")
            
            st.success(f"{len(st.session_state['current_uploaded_files'])} file(s) selected. Click 'Run Ethics Review' to process all documents.")
        else:
            st.info("No files selected yet.")
    else:
        if st.session_state['current_uploaded_files']:
            st.session_state['current_uploaded_files'] = []
            st.session_state['ai_review'] = ''
            st.session_state['processed_ai_review_docx_buffer'] = None
            st.session_state['review_triggered'] = False
            
    # Conditional rerun for admin tabs, executed at a higher level
    if st.session_state['trigger_rerun_for_admin_tab']:
        st.session_state['trigger_rerun_for_admin_tab'] = False
        st.rerun()

    if st.button("Run Ethics Review", key="run_ai_review_btn"):
        if st.session_state['current_uploaded_files']:
            combined_document_text = ""
            with st.spinner("Combining document text and running AI ethics review... This may take a moment."):
                for file_obj in st.session_state['current_uploaded_files']:
                    try:
                        pdf_reader = PyPDF2.PdfReader(file_obj)
                        for page_num in range(len(pdf_reader.pages)):
                            combined_document_text += pdf_reader.pages[page_num].extract_text() + "\n\n"
                        st.write(f"Processed: {file_obj.name}")
                    except Exception as e:
                        st.error(f"Error processing PDF '{file_obj.name}': {e}. Skipping this file.")
                        combined_document_text += f"\n\n--- Error processing {file_obj.name}: {e} ---\n\n"
            
            if combined_document_text.strip():
                ai_review_content = get_ai_review_response(combined_document_text)
                
                st.session_state['ai_review'] = ai_review_content
                st.session_state['processed_ai_review_docx_buffer'] = create_docx_from_text(st.session_state['ai_review'])
                st.session_state['review_triggered'] = True
                st.rerun()
            else:
                st.error("No text could be extracted from the uploaded PDF(s). Please check the files.")
                st.session_state['review_triggered'] = False
        else:
            st.warning("Please upload at least one PDF document to run the ethics review.")

    if st.session_state.get('review_triggered') and st.session_state.get('ai_review') and st.session_state.get('processed_ai_review_docx_buffer'):
        st.markdown("---")
        st.markdown("<h3 style='color:#06038D;'>📊 Consolidated AI Ethics Review Report</h3>", unsafe_allow_html=True)
        
        processed_ai_review_display = re.sub(r'(\*\*(\d+\.\s|ICMR|Medical Devices|New Drugs|English and Construction|Questionnaire and Informed Consent Alignment|Additional Aspects|Summary and Recommendation|Additional Questions).*?:?\*\*)\n', r'\n\n\1\n', st.session_state['ai_review'], flags=re.MULTILINE)
        processed_ai_review_display = re.sub(r'(\* \*\*Area of Review:\*\*|\* \*\*Section/Clause:\*\*|\* \*\*Compliance:\*\*|\* \*\*Explanation:\*\*|\* \*\*Concern:\*\*)\n', r'\n  \1', processed_ai_review_display, flags=re.MULTILINE)
        processed_ai_review_display = processed_ai_review_display.replace("---", "\n\n---\n\n")
        st.write(processed_ai_review_display)

        generate_download_buttons(
            st.session_state['user_name'],
            st.session_state['ai_review'],
            st.session_state['processed_ai_review_docx_buffer'],
            st.session_state['user_uid']
        )
        
        if st.button("Clear All Reviews and Upload New Document(s)", key="clear_all_reviews_btn"):
            st.session_state['ai_review'] = ''
            st.session_state['processed_ai_review_docx_buffer'] = None
            st.session_state['review_triggered'] = False
            st.session_state['current_uploaded_files'] = []
            st.rerun()

    # --- ADMIN-SPECIFIC SECTION (ONLY VISIBLE TO ADMINS) ---
    if st.session_state['user_role'] == 'admin':
        st.markdown("---")
        st.markdown("<h3 style='color:#06038D;'>⚙️ Admin Features</h3>", unsafe_allow_html=True)
        st.info("This section is visible only to users with the 'admin' role.")
        
        # --- Custom Button-based Tab Navigation ---
        col1, col2, col3 = st.columns(3)
        
        def set_admin_tab(tab_name):
            if st.session_state['active_admin_tab'] != tab_name:
                st.session_state['active_admin_tab'] = tab_name
                st.session_state['trigger_rerun_for_admin_tab'] = True
                st.cache_data.clear()

        # Display buttons with bold text for active tab, regular for inactive, and alignment styling
        with col1:
            if st.session_state['active_admin_tab'] == "Invite New User":
                # Using inline HTML for precise control over text styling and alignment
                st.markdown(f"<span style='font-size: 1.1rem; font-weight: bold; color:#06038D; display: inline-block; padding-top: 0.6rem; padding-bottom: 0.6rem; margin-top: 0; margin-bottom: 0;'>Invite New User</span>", unsafe_allow_html=True)
            else:
                st.button("Invite New User", key="tab_button_invite_new_user", on_click=set_admin_tab, args=("Invite New User",))
        with col2:
            if st.session_state['active_admin_tab'] == "User Management":
                st.markdown(f"<span style='font-size: 1.1rem; font-weight: bold; color:#06038D; display: inline-block; padding-top: 0.6rem; padding-bottom: 0.6rem; margin-top: 0; margin-bottom: 0;'>User Management</span>", unsafe_allow_html=True)
            else:
                st.button("User Management", key="tab_button_user_management", on_click=set_admin_tab, args=("User Management",))
        with col3:
            if st.session_state['active_admin_tab'] == "View All Review Reports":
                st.markdown(f"<span style='font-size: 1.1rem; font-weight: bold; color:#06038D; display: inline-block; padding-top: 0.6rem; padding-bottom: 0.6rem; margin-top: 0; margin-bottom: 0;'>View All Review Reports</span>", unsafe_allow_html=True)
            else:
                st.button("View All Review Reports", key="tab_button_view_reports", on_click=set_admin_tab, args=("View All Review Reports",))

        st.markdown("---")

        # --- Tab Content Rendering based on active_admin_tab ---
        if st.session_state['active_admin_tab'] == "Invite New User":
            st.markdown("#### Invite New User")
            
            if st.session_state['active_admin_tab'] != st.session_state.get('last_rendered_admin_tab'):
                st.session_state['invite_details_to_display'] = None
                st.session_state['show_admin_invite_confirm'] = False
                st.session_state['temp_invited_email'] = ""
                st.session_state['temp_invite_role'] = ""
            st.session_state['last_rendered_admin_tab'] = "Invite New User"

            with st.form("invite_user_form_main", clear_on_submit=False):
                invited_email_input = st.text_input("New User Email", key="new_user_email_input_form", value=st.session_state.get('temp_invited_email', ''))
                invite_role_radio = st.radio(
                    "Assign Role:",
                    ('user', 'admin'),
                    index=0,
                    key="invite_role_radio_form"
                )
                invite_button_submit = st.form_submit_button("Invite User")

                if invite_button_submit:
                    st.session_state['invite_details_to_display'] = None
                    st.session_state['show_admin_invite_confirm'] = False
                    if not invited_email_input or not invited_email_input.strip():
                        st.warning("Please enter an email address to invite a user.")
                    else:
                        st.session_state['temp_invited_email'] = invited_email_input.strip()
                        st.session_state['temp_invite_role'] = invite_role_radio
                        if invite_role_radio == 'admin':
                            st.session_state['show_admin_invite_confirm'] = True
                        st.session_state['trigger_rerun_for_admin_tab'] = True
                        
            if st.session_state['show_admin_invite_confirm']:
                st.warning(f"You are about to invite an **Admin** user with email: `{st.session_state['temp_invited_email']}`. This grants full administrative control.")
                st.info("Please confirm to proceed.")
                col_confirm, col_cancel = st.columns(2)
                with col_confirm:
                    if st.button("Confirm Invite Admin", key="confirm_admin_invite_final_btn"):
                        try:
                            user = auth.create_user(email=st.session_state['temp_invited_email'])
                            new_user_uid = user.uid
                            st.session_state['firestore_db'].collection('User').document(new_user_uid).set({
                                'email': st.session_state['temp_invited_email'],
                                'role': st.session_state['temp_invite_role'],
                                'first_name': '',
                                'last_name': '',
                                'created_at': firestore.SERVER_TIMESTAMP
                            })
                            password_reset_link = auth.generate_password_reset_link(st.session_state['temp_invited_email'])
                            
                            st.session_state['invite_details_to_display'] = {
                                'email': st.session_state['temp_invited_email'],
                                'link': password_reset_link,
                                'role': st.session_state['temp_invite_role']
                            }
                            st.success(f"User {st.session_state['temp_invited_email']} invited successfully as **{st.session_state['temp_invite_role'].capitalize()}**! See details below.")
                        except auth.EmailAlreadyExistsError:
                            st.error(f"Error: A user with email {st.session_state.get('temp_invited_email')} already exists.")
                        except Exception as e:
                            st.error(f"Error inviting user: {e}")
                        finally:
                            st.session_state['show_admin_invite_confirm'] = False
                            st.session_state['temp_invited_email'] = ""
                            st.session_state['temp_invite_role'] = ""
                            st.session_state['trigger_rerun_for_admin_tab'] = True
                with col_cancel:
                    if st.button("Cancel Invite", key="cancel_admin_invite_final_btn"):
                        st.session_state['show_admin_invite_confirm'] = False
                        st.session_state['temp_invited_email'] = ""
                        st.session_state['temp_invite_role'] = ""
                        st.info("Admin invitation cancelled.")
                        st.session_state['trigger_rerun_for_admin_tab'] = True
            elif st.session_state.get('temp_invited_email') and st.session_state.get('temp_invite_role') == 'user' and not st.session_state['show_admin_invite_confirm']:
                try:
                    user = auth.create_user(email=st.session_state['temp_invited_email'])
                    new_user_uid = user.uid
                    st.session_state['firestore_db'].collection('User').document(new_user_uid).set({
                        'email': st.session_state['temp_invited_email'],
                        'role': st.session_state['temp_invite_role'],
                        'first_name': '',
                        'last_name': '',
                        'created_at': firestore.SERVER_TIMESTAMP
                    })
                    password_reset_link = auth.generate_password_reset_link(st.session_state['temp_invited_email'])
                    
                    st.session_state['invite_details_to_display'] = {
                        'email': st.session_state['temp_invited_email'],
                        'link': password_reset_link,
                        'role': st.session_state['temp_invite_role']
                    }
                    st.success(f"User {st.session_state['temp_invited_email']} invited successfully as **{st.session_state['temp_invite_role'].capitalize()}**! See details below.")
                except auth.EmailAlreadyExistsError:
                    st.error(f"Error: A user with email {st.session_state['temp_invited_email']} already exists.")
                except Exception as e:
                    st.error(f"Error inviting user: {e}")
                finally:
                    st.session_state['temp_invited_email'] = ""
                    st.session_state['temp_invite_role'] = ""
                    st.session_state['trigger_rerun_for_admin_tab'] = True

            if st.session_state['invite_details_to_display']:
                st.markdown("---")
                st.markdown("#### New User Invitation Details")
                details = st.session_state['invite_details_to_display']
                st.info(f"**Email:** `{details['email']}`")
                st.info(f"**Role:** `{details['role'].capitalize()}`")
                
                st.markdown("---")
                st.markdown("**Action Required:**")
                st.info(f"Please **copy the link below** and **send it to the new user's email address** (`{details['email']}`). This link allows them to set their password for the first time.")
                
                st.text_area(
                    "Password Reset Link:",
                    value=details['link'],
                    height=70,
                    disabled=True,
                    key="copyable_password_reset_link_textarea"
                )
                
                st.warning("This link is valid for a limited time. Please send it promptly.")
                if st.button("Clear Invitation Details", key="clear_invite_details_btn"):
                    st.session_state['invite_details_to_display'] = None
                    st.session_state['trigger_rerun_for_admin_tab'] = True
        
        elif st.session_state['active_admin_tab'] == "User Management":
            st.markdown("#### User Management")
            st.write("Here you can view details of all registered users in your system.")

            if st.session_state['active_admin_tab'] != st.session_state.get('last_rendered_admin_tab'):
                st.session_state['temp_invited_email'] = ""
                st.session_state['temp_invite_role'] = ""
                st.session_state['invite_details_to_display'] = None
                st.session_state['show_admin_invite_confirm'] = False
            st.session_state['last_rendered_admin_tab'] = "User Management"

            if st.button("Refresh User List", key="refresh_users_btn_tab"):
                st.cache_data.clear()
                st.session_state['trigger_rerun_for_admin_tab'] = True
            
            with st.spinner("Fetching user list..."):
                @st.cache_data(ttl=300, show_spinner=False)
                def get_all_users_data_tab_cached():
                    users_list = []
                    try:
                        page = auth.list_users()
                        
                        firestore_profiles = {}
                        docs = st.session_state['firestore_db'].collection('User').stream()
                        for doc in docs:
                            firestore_profiles[doc.id] = doc.to_dict()
                        for user_record in page.users:
                            user_data_row = {
                                "UID": user_record.uid,
                                "Email": user_record.email,
                                "Email Verified": user_record.email_verified,
                                "Disabled": user_record.disabled,
                                "First Name": "",
                                "Last Name": "",
                                "Role": "N/A",
                            }
                            
                            if user_record.uid in firestore_profiles:
                                firestore_data = firestore_profiles[user_record.uid]
                                user_data_row["First Name"] = firestore_data.get('first_name', '')
                                user_data_row["Last Name"] = firestore_data.get('last_name', '')
                                user_data_row["Role"] = firestore_data.get('role', 'N/A')
                            else:
                                user_data_row["Role"] = "No Firestore Profile"

                            users_list.append(user_data_row)
                        
                        return pd.DataFrame(users_list)
                    except Exception as e:
                        st.error(f"Error fetching users: {e}")
                        return pd.DataFrame()

                all_users_df = get_all_users_data_tab_cached()
                if not all_users_df.empty:
                    st.dataframe(all_users_df, use_container_width=True)
                else:
                    st.info("No users found or an error occurred while fetching users.")
        
        elif st.session_state['active_admin_tab'] == "View All Review Reports":
            st.markdown("#### View All Review Reports")
            st.write("Here, you can view and manage all AI Ethics Review Reports generated by users.")

            if st.session_state['active_admin_tab'] != st.session_state.get('last_rendered_admin_tab'):
                st.session_state['temp_invited_email'] = ""
                st.session_state['temp_invite_role'] = ""
                st.session_state['invite_details_to_display'] = None
                st.session_state['show_admin_invite_confirm'] = False
            st.session_state['last_rendered_admin_tab'] = "View All Review Reports"

            db = st.session_state['firestore_db']
            try:
                with st.spinner("Fetching review reports..."):
                    reviews_ref = db.collection('ethics_review_outputs')
                    all_reviews_docs = reviews_ref.stream()

                    reviews_data = []
                    user_profiles_map = {}
                    users_collection_ref = db.collection('User')
                    user_docs = users_collection_ref.stream()
                    for user_doc in user_docs:
                        user_profiles_map[user_doc.id] = user_doc.to_dict()

                    @st.cache_data(ttl=3600, show_spinner=False)
                    def get_user_name_from_uid_cached(uid, profiles_map):
                        user_data = profiles_map.get(uid)
                        if user_data:
                            first = user_data.get('first_name', '')
                            last = user_data.get('last_name', '')
                            if first or last:
                                return f"{first} {last}".strip()
                            else:
                                return user_data.get('email', 'Unknown User')
                        return "Unknown User"

                    for doc in all_reviews_docs:
                        review = doc.to_dict()
                        review_id = doc.id
                        
                        uploader_name = "Unknown User"
                        if 'uploaderUID' in review and review['uploaderUID']:
                            uploader_name = get_user_name_from_uid_cached(review['uploaderUID'], user_profiles_map)
                        
                        reviews_data.append({
                            "Review ID": review_id,
                            "File Name": review.get('outputDocFileName', 'N/A'),
                            "Generated By": uploader_name,
                            "Date Generated": review.get('review_date_human', 'N/A'),
                            "Status": review.get('status', 'N/A'),
                            "Summary": review.get('summary', 'No summary'),
                            "Download URL": review.get('outputDocURL', '')
                        })
                    
                if reviews_data:
                    reviews_data.sort(key=lambda x: datetime.strptime(x["Date Generated"], "%Y-%m-%d %H:%M:%S") if x["Date Generated"] != 'N/A' else datetime.min, reverse=True)
                    
                    df = pd.DataFrame(reviews_data)
                    st.dataframe(df,
                                 column_config={
                                     "Download URL": st.column_config.LinkColumn("Download File", display_text="⬇ Download")
                                 },
                                 hide_index=True,
                                 use_container_width=True)
                else:
                    st.info("No ethics review reports found yet.")
            except Exception as e:
                st.error(f"Error fetching review reports: {e}")


# --- FOOTER (Always visible) ---
st.markdown(
    '<div style="text-align:center; color:#FF671F; margin-top:30px;">©copyright SSO Consultants</div>',
    unsafe_allow_html=True
)
