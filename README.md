# e-Hastakshar — Document Approval Workflow

A Python/Flask prototype for preparing documents, converting Word files to PDF, adding visible signature-related blocks and QR codes, and moving documents through a role-based approval workflow.

> **Status: local prototype, not production-ready.** This project is not a substitute for a cryptographically verifiable digital-signature service. Run it with synthetic documents only until security and document-integrity testing is complete.

## Current workflow

1. A Prepared user registers and signs in.
2. The user uploads a PDF or Word document.
3. The application processes the document and adds a visible signature-related block and QR code.
4. The Prepared user verifies and submits the document.
5. A Certifier account reviews and approves or rejects it.
6. The application records status and activity history.

## Technology

- Python and Flask
- HTML templates
- PyPDF2, PyMuPDF, pdfminer.six, ReportLab
- qrcode and Pillow
- pywin32 / Windows COM for Word-to-PDF conversion

## Requirements

- Windows recommended; Word-to-PDF conversion requires Microsoft Word and Windows COM.
- Python 3.10 or newer.
- For a clean local run, install dependencies from `requirements.txt`.

## Setup (PowerShell)

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
$env:DIGITAL_SIGNATURE_SECRET = [guid]::NewGuid().ToString() + [guid]::NewGuid().ToString()
python app.py
```

Open `http://127.0.0.1:5057`. The development server is bound to localhost and debug mode is disabled. The secret variable must be set again in each new shell unless you configure it in your local environment. The app creates `documents.json`, `history.json`, `users.json`, and the `upload/` directory beside `app.py` when needed.

## Accounts and approval workflow

- Use the registration form for the Prepared and Manager roles.
- The registration form does not allow users to self-assign the Certifier role.
- Create a local Certifier account from the project directory using:

```powershell
python create_certifier.py
```

The helper asks for a display name, PB number, and password, and stores a password hash in `users.json`. Keep `users.json` and uploaded documents private. This local JSON-based account store is intended for development, not multi-user production hosting.

## Security and known limitations

- New passwords are hashed; legacy plaintext passwords are upgraded to hashes after a successful login.
- Document actions now check the user's role and, where applicable, document ownership.
- Uploaded files require authentication and role/ownership authorization.
- Uploads are capped at 50 MB and the server listens on localhost with debug mode off.
- JSON files are not a concurrent database; use a proper database and transaction-safe storage for multi-user deployment.
- Add CSRF protection, stronger validation, automated tests, and a security review before any public deployment.
- A visible stamp and QR code do not, by themselves, prove cryptographic document integrity.

## Testing checklist

1. Register a Prepared account and a Manager account.
2. Run `python create_certifier.py` to create a Certifier.
3. Upload a non-sensitive test PDF and verify it.
4. Submit it to the Certifier and test both approval and rejection.
5. Confirm unauthorized roles cannot perform Certifier actions.
6. Confirm a logged-out visitor cannot access uploaded files.

## Author

Harish T · Python Backend Developer

Repository: https://github.com/Hari24022002/Digital-Signature
