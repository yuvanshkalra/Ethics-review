
# Ethics Review System (ECR) 🏥⚖️
### AI-Powered Ethics Review for Biomedical and Health Research

[![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Framework-Streamlit-FF4B4B?logo=streamlit)](https://streamlit.io/)
[![OpenAI](https://img.shields.io/badge/AI-GPT--4o--mini-412991?logo=openai)](https://openai.com/)
[![Firebase](https://img.shields.io/badge/Backend-Firebase-FFCA28?logo=firebase)](https://firebase.google.com/)
[![Status](https://img.shields.io/badge/Project-AI%20Ethics%20Review-purple)](https://github.com/yuvanshkalra/Ethics-review)

**Ethics Review System (ECR)** is an AI-assisted application built with Streamlit and OpenAI that reviews biomedical and health research submissions against key Indian regulatory and ethical guidelines.

The system accepts multiple PDF documents belonging to a single research submission, combines their extracted text, and generates a consolidated ethics review report. It also provides Firebase-powered authentication, role-based access, administrative user management, and cloud storage for generated reports.

---

## 📌 Table of Contents

- [Overview](#-overview)
- [Problem Statement](#-problem-statement)
- [Key Features](#-key-features)
- [Regulatory Frameworks](#-regulatory-frameworks)
- [How It Works](#-how-it-works)
- [Technology Stack](#-technology-stack)
- [Project Structure](#-project-structure)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Firebase and OpenAI Configuration](#-firebase-and-openai-configuration)
- [Run the Application](#-run-the-application)
- [Generated Report](#-generated-report)
- [Security and Ethical Considerations](#-security-and-ethical-considerations)
- [Future Enhancements](#-future-enhancements)
- [Author](#-author)

---

## 🚀 Overview

Reviewing biomedical and health research submissions involves examining research proposals, questionnaires, informed consent documents, and other supporting materials for ethical and regulatory concerns.

Manual review of these documents can be time-consuming, particularly when reviewers must assess consistency across multiple documents.

The Ethics Review System helps streamline this process by consolidating the text from multiple PDF documents and using an AI language model to produce a structured preliminary ethics review.

### 🎯 Project Objectives

- Automate preliminary analysis of research ethics submissions.
- Consolidate multiple documents into one review workflow.
- Evaluate submission completeness and potential compliance concerns.
- Examine research documents against relevant Indian guidelines.
- Review questionnaire construction and informed consent alignment.
- Generate structured reports with observations and recommendations.
- Provide authenticated access and administrative user management.
- Save completed review reports in Firebase cloud storage.

---

## ❓ Problem Statement

Ethics committees and research reviewers need to evaluate research submissions carefully to identify missing information, potential ethical concerns, inconsistencies, and areas requiring clarification.

When documents are reviewed separately, important relationships between the research proposal, questionnaire, and informed consent materials can be overlooked.

This project addresses the need for an AI-assisted preliminary review workflow that analyzes multiple documents together and organizes its findings into a consolidated report for human reviewers.

---

## ✨ Key Features

### 1. 📄 Multiple PDF Document Uploads

Users can upload multiple PDF documents belonging to the same research submission, including:

- Research application forms
- Research proposals
- Questionnaires
- Informed consent documents
- Supporting letters and other relevant materials

The application extracts text from the uploaded PDFs and combines the extracted content for a unified review.

**Current limitation:** The uploader accepts PDF files only. Scanned PDFs without an extractable text layer may require OCR support, which is not implemented in the current workflow.

### 2. 🧠 AI-Powered Ethics Review

The application uses OpenAI's `gpt-4o-mini` model to generate a consolidated review of the submitted document content.

The review prompt is designed to assess the submission against relevant Indian research ethics and regulatory frameworks.

### 3. ⚖️ Regulatory Guideline Assessment

The AI review is designed to cover:

- ICMR National Ethical Guidelines for Biomedical and Health Research Involving Human Participants, 2017
- Medical Devices Rules, 2017 (MDR 2017)
- New Drugs and Clinical Trials Rules, 2019 (NDCTR 2019)

The report organizes potential compliance issues, observations, and recommendations under these frameworks where applicable.

### 4. 📋 Structured Review Report

The generated report includes seven main sections:

1. Document Submission Status
2. Review Against Guidelines
3. English and Construction of the Questionnaire
4. Questionnaire and Informed Consent Alignment
5. Additional Aspects
6. Summary and Recommendation
7. Additional Questions

Sections relating to questionnaires and informed consent are assessed when the relevant materials are present in the submission.

### 5. 📝 Downloadable DOCX Reports

The application converts the generated AI review into a formatted Microsoft Word document.

Report-generation features include:

- A report title
- Logo placement
- Structured headings
- Bullet-point observations
- Section separators
- Timestamped report filenames

Users can download the consolidated ethics review report directly from the application.

### 6. ☁️ Firebase Cloud Storage

The system integrates with Firebase to save completed reports.

- Uploads generated DOCX reports to Firebase Storage.
- Records report metadata in Cloud Firestore.
- Stores the report filename, uploader UID, timestamp, status, summary, and download URL.
- Provides a cloud-storage workflow for accessing generated reports.

### 7. 🔐 Authentication and Role-Based Access

The application integrates Firebase Authentication with Cloud Firestore for account and profile management.

Features include:

- Separate login flows for users and administrators
- Email and password authentication
- Firebase ID-token verification
- User profile management
- First-time name setup
- Session management and logout
- Role verification for administrative access

### 8. 🛠️ Administrator Dashboard

Users with administrator privileges have access to additional controls, including:

- Inviting new users
- Assigning user or administrator roles
- Confirming administrator invitations
- Generating password-reset links for invited users
- Managing user accounts
- Reviewing stored report records
- Accessing report download links
- Administrative account and report operations

Administrative functions are displayed according to the authenticated user's role.

---

## ⚖️ Regulatory Frameworks

The application focuses on three key Indian regulatory and ethical frameworks.

| Framework | Review focus |
|---|---|
| ICMR Ethical Guidelines, 2017 | Ethical considerations in biomedical and health research involving human participants |
| Medical Devices Rules, 2017 | Potential regulatory considerations for research involving medical devices |
| New Drugs and Clinical Trials Rules, 2019 | Potential regulatory considerations for new drugs and clinical trials |

The applicability of each framework depends on the nature of the research submission.

The application provides AI-generated observations rather than a legally authoritative compliance determination. Regulatory applicability and findings must be verified by qualified reviewers.

---

## ⚙️ How It Works

```text
┌──────────────────────────────┐
│     User Authentication      │
│   Firebase Authentication   │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Upload Multiple PDF Documents│
│ Proposal, Questionnaire, etc. │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Extract Text from PDF Files  │
│           PyPDF2             │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Combine Extracted Documents  │
│ Into a Single Submission     │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ OpenAI GPT-4o-mini           │
│ AI Ethics Review             │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Consolidated Review Findings │
│ Observations & Recommendations│
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Generate DOCX Report         │
│          python-docx         │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│ Download Report              │
│ Firebase Storage + Firestore│
└──────────────────────────────┘
```

### Workflow Explanation

1. A registered user signs in through the appropriate login flow.
2. The user uploads one or more PDF documents.
3. The application extracts text from each PDF using PyPDF2.
4. Extracted text is combined into a single submission.
5. The combined text is sent to OpenAI's `gpt-4o-mini` model with a structured ethics-review prompt.
6. The AI generates observations, potential compliance concerns, recommendations, and follow-up questions.
7. The results are displayed in the Streamlit interface.
8. A formatted DOCX report is created using `python-docx`.
9. When the report download action is triggered, the application uploads the report to Firebase Storage and records its metadata in Firestore.

---

## 🛠️ Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application logic |
| Streamlit | Interactive web application interface |
| OpenAI API | AI-generated ethics review |
| GPT-4o-mini | Language model used for review generation |
| PyPDF2 | PDF text extraction |
| python-docx | DOCX report generation |
| Firebase Authentication | User authentication |
| Firebase Admin SDK | Administrative Firebase operations |
| Cloud Firestore | User profiles and report metadata |
| Firebase Storage | Generated report storage |
| Requests | Direct Firebase Authentication REST API calls |
| bcrypt | Password hashing library included in the project dependencies |
| pandas | Data processing library included in dependencies |
| LangChain | Language-model tooling included in dependencies |

---

## 📂 Project Structure

The repository's top-level files are:

```text
Ethics-review/
│
├── app.py
│   └── Main Streamlit application
│       ├── Firebase initialization
│       ├── Authentication and session management
│       ├── PDF upload and text extraction
│       ├── AI ethics review generation
│       ├── DOCX report creation
│       ├── Firebase report storage
│       └── Administrator dashboard
│
├── requirements.txt
│   └── Python package dependencies
│
├── README.md
│   └── Project documentation
│
├── logo.png
│   └── Application sidebar and report logo
│
└── logo_ecr.png
    └── Header logo referenced by the application
```

**Note:** The two logo filenames are referenced in `app.py`. Ensure both image files are present in the working directory when running the application.

---

## 💻 Prerequisites

Before running the project, ensure you have:

- Python 3.x installed
- pip package manager
- An OpenAI API key with access to `gpt-4o-mini`
- A Firebase project
- Firebase Authentication configured for email/password sign-in
- Cloud Firestore enabled
- Firebase Storage configured
- Firebase Admin SDK service-account credentials
- Internet access for API calls and cloud services

---

## 📥 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yuvanshkalra/Ethics-review.git
```

### 2. Navigate to the Project Directory

```bash
cd Ethics-review
```

### 3. Create a Virtual Environment

```bash
python -m venv venv
```

Activate the virtual environment.

**Windows**

```bash
venv\Scripts\activate
```

**macOS / Linux**

```bash
source venv/bin/activate
```

### 4. Install Dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

---

## 🔑 Firebase and OpenAI Configuration

The application expects its credentials to be configured through Streamlit secrets.

Create the following file in the project root:

```text
.streamlit/
└── secrets.toml
```

Configure the following sections using the actual values from your Firebase and Google Cloud projects.

```toml
OPENAI_API_KEY = "your-openai-api-key"

[firebase]
apiKey = "your-firebase-web-api-key"
storageBucket = "your-firebase-storage-bucket"

[firebase_admin_sdk]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-private-key-id"
private_key = "your-private-key-with-escaped-newlines"
client_email = "your-service-account-email"
client_id = "your-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
client_x509_cert_url = "your-service-account-certificate-url"
universe_domain = "googleapis.com"
```

### Configuration Notes

- Replace every placeholder with the corresponding real configuration value.
- Obtain `OPENAI_API_KEY` from your OpenAI API account.
- Obtain the Firebase web API key from your Firebase project configuration.
- Obtain the service-account fields from the Firebase Admin SDK JSON file.
- Preserve the correct private-key newline formatting. The application converts escaped `\n` sequences to actual newline characters.
- Configure the Firebase Storage bucket correctly for your project.
- Enable email/password authentication in Firebase Authentication.
- Ensure the required Firestore collections and user records are configured.
- Configure appropriate Firebase Storage and Firestore permissions.

The application expects user profiles in the `User` collection and saves generated report metadata in `ethics_review_outputs`.

**Important:** Never commit `.streamlit/secrets.toml`, service-account private keys, or API keys to GitHub. For Streamlit Community Cloud, add the configuration through the deployment's secrets settings.

---

## ▶️ Run the Application

Once the dependencies and credentials are configured, run:

```bash
streamlit run app.py
```

Open the local URL shown in the terminal, usually:

```text
http://localhost:8501
```

### Using the Application

1. Choose either **Login as Admin** or **Login as User**.
2. Sign in with an account configured in Firebase.
3. Complete the profile setup if prompted.
4. Upload the relevant PDF documents for the research submission.
5. Click **Run Ethics Review**.
6. Review the consolidated report displayed on the page.
7. Download the generated DOCX report.
8. The report is uploaded to Firebase Storage and its metadata is saved in Firestore through the download workflow.

Administrators can additionally access the administrative dashboard after logging in with an account that has the required role.

---

## 📄 Generated Report

The application generates a Word document titled **AI Ethics Review Report**.

The report contains:

- Document submission status
- Review against ICMR 2017
- Medical Devices Rules, 2017 review
- New Drugs and Clinical Trials Rules, 2019 review
- Questionnaire language and construction assessment
- Questionnaire and informed consent alignment
- Additional ethical and regulatory observations
- Summary and recommendations
- Additional questions for the submitter

The report is generated from the model's text output and formatted using `python-docx`.

---

## 🔒 Security and Ethical Considerations

Because research submissions may contain sensitive personal, medical, or participant information, security and privacy must be treated as essential requirements.

- Keep API keys and service-account credentials private.
- Apply least-privilege permissions to Firebase services.
- Restrict administrator roles to authorized personnel.
- Review Firebase Storage access rules before deploying.
- Avoid exposing sensitive research documents through unrestricted public links.
- Confirm the organization's data-handling and retention requirements before uploading submissions to external AI services.
- Review AI-generated observations for factual accuracy and regulatory applicability.
- Do not treat the report as a substitute for ethics committee approval or qualified regulatory advice.

**Deployment security note:** The current code calls `make_public()` on generated report files. Review this behavior before using real research data, since public download URLs can expose documents beyond their intended audience.

---

## 🚧 Future Enhancements

Potential improvements include:

- OCR support for scanned PDF documents.
- A review history dashboard with search and filtering.
- More granular access controls for stored reports.
- Private, authenticated report downloads instead of public URLs.
- Better validation of extracted text and AI responses.
- More detailed guideline references and traceable evidence for each finding.
- Report export to PDF.
- Structured severity levels for identified concerns.
- Automated tests for authentication, document processing, and report generation.
- Audit logs for administrative actions.
- Improved error handling and API retry logic.
- Configurable review templates for different research types.

---

## 🎓 Learning Outcomes

This project demonstrates practical experience with:

- Python application development
- Streamlit web interfaces
- Large language model API integration
- PDF processing and document aggregation
- Automated report generation
- Firebase authentication and cloud services
- Role-based access management
- Cloud document storage
- AI-assisted regulatory and ethical document review

---

## 👨‍💻 Author

**Yuvansh Kalra**

- GitHub: [@yuvanshkalra](https://github.com/yuvanshkalra)
- Project Repository: [Ethics-review](https://github.com/yuvanshkalra/Ethics-review)

---

## ⚖️ Disclaimer

This project is intended to support preliminary research ethics review. AI-generated findings may be incomplete, inaccurate, or unsuitable for a particular submission. The application does not provide formal ethics committee approval, legal advice, or a definitive regulatory compliance determination.

All findings should be independently verified by qualified research ethics and regulatory professionals.

---

⭐ If you find this project useful, consider starring the repository!
