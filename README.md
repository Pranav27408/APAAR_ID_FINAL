# APAAR Student Data & Verification Platform

A responsive prototype for SIH 2026 Problem Statement 26263. The frontend is React/Vite; the API is FastAPI with SQLite. All included student information is synthetic. APAAR and AICTE checks are clearly marked simulations; this project does not contact official systems.

## Included workflows

- Role-selected login for Student, College Staff, and AICTE Officer.
- Student dashboard with profile, semester records, documents, verification status, and correction request form.
- Staff dashboard with student search/filter, add records, spreadsheet import, CSV export, document review, and audit view.
- AICTE dashboard with institution submission queue, record review actions, and analytics.
- FastAPI authentication using bcrypt password hashes and signed JWTs, role checks, server-side APAAR/intake validation, duplicate detection, SQLite persistence, submission review, and audit logs.
- Responsive navigation and layouts, synthetic sample data, validation feedback, and mock integration notices.

## Requirements

- Node.js 20 or newer and npm.
- Python 3.10 or newer (3.11 recommended).
- VS Code (optional; steps below assume it).

## Run locally in VS Code

1. Open this folder in VS Code.
2. Start the API in a VS Code terminal:

   ```powershell
   cd backend
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   Copy-Item .env.example .env
   # Change JWT_SECRET in backend/.env before using beyond a local demo.
   uvicorn main:app --reload --port 8000
   ```

   On systems where `py` is unavailable, substitute `python` for `py`. The database is created as `backend/apaar.db` with synthetic records at first launch. API docs are at http://localhost:8000/docs.

3. In a second terminal at the project root:

   ```powershell
   npm install
   Copy-Item .env.example .env
   npm run dev
   ```

4. Open the Vite URL shown in the terminal (normally http://localhost:5173).

## Demo accounts

| Portal | User ID | Password |
| --- | --- | --- |
| Student | `student@demo.edu` | `Student123!` |
| College Staff | `staff@demo.edu` | `Staff123!` |
| AICTE Officer | `officer@aicte.gov.in` | `Officer123!` |

Enter the displayed CAPTCHA to sign in. If the API is not running, the UI allows local demonstration using these demo credentials; changes then remain in the browser session. Run the API for persistent CRUD and server-side checks.

## Spreadsheet import

Upload `.xlsx`, `.xlsm`, or `.csv` from Data import. Required columns are `name`, `register_no`, `department`, `year`, and `apaar_id`. Optional fields are `email`, `intake`, `enrolment`, `supernumerary`, `nri`, `oci`, and `foreign`. APAAR IDs must be exactly 12 digits. Duplicate IDs and register numbers are flagged. Successfully imported rows enter Pending review. A sample template is available at `student-import-template.csv`.

## API summary

- `POST /api/auth/login`, `GET /api/auth/me`
- `GET/POST /api/students`, `PUT/DELETE /api/students/{id}`
- `POST /api/students/{id}/review`, `POST /api/submissions`, `GET /api/submissions`
- `POST /api/import`, `GET /api/students/export.csv`
- `GET /api/analytics`, `GET /api/audit`, `GET /api/institutions`, `GET /api/duplicates/{apaar_id}`
- `POST/GET /api/corrections`, `GET /api/health`

## Prototype boundary

No real student records, official APAAR services, DigiLocker, or AICTE systems are used. Before any real deployment, configure a strong random `JWT_SECRET`, HTTPS, institutional identity integration, secure file storage and malware scanning, retention controls, rate limits, operational monitoring, and an approved API/data-sharing agreement. Do not use the demo credentials in a deployed environment.
