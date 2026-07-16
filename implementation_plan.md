# BSPS DB: Web-Based Housing Assistance Verification Database System

An implementation plan for building **BSPS DB**, a premium web-based database system designed to read, verify, reconcile, and export Calon Penerima Bantuan (CPB) housing assistance data by uploading Excel spreadsheets.

---

## User Review Required

Please review the revised architectural decisions and key design elements:

### 1. Local-First Architecture with Deployment Readiness
*   **Database**: SQLite to store CPB data locally.
*   **Backend**: A lightweight Python server (FastAPI/Flask) running locally to handle Excel parsing (`openpyxl`, `pandas`).
*   **Frontend**: React (Single Page Application) with Vite, communicating via REST APIs. This decoupling ensures the app can easily be moved to a remote cloud server (e.g., PostgreSQL + cloud storage) in the future.

### 2. Reconciliation Center
Instead of making the user search through thousands of records in a raw table, we will implement a dedicated **Reconciliation Center**. 
*   Discrepancies will be highlighted using visual side-by-side card components showing where the input data does not match the INVERS record.
*   Cards will highlight the mismatched key fields (`Nama`, `No. NIK`, `No. KK`) in red and display context like `Desa / Kelurahan` to help local verification.
*   The user will have action buttons directly on the cards to:
    *   **Accept Differences**: Keep the verified version's data and mark it as reconciled.
    *   **Manual Fix**: Manually edit the NIK, KK, or Name in-app to resolve the mismatch.

### 3. Styled Excel Export
Reports will be exported as Excel files rather than PDFs. However, they will be visually structured like `EXPORT-FORMAT-EXAMPLE.pdf` using Python (`openpyxl`) with the following strict styling parameters:
*   **Font**: Bookman Old Style, 12pt for all content.
*   **Borders**: Thin black borders around all table data.
*   **Layout**: Headers, signature blocks, and columns structured to match the visual hierarchy of the PDF reference.

---

## Technical Specifications

### Data Matching & Validation Rules
*   **Key Checks**: Matches `Nama`, `No. NIK`, and `No. KK`.
*   **Matching Mode**: Strict exact match (case-sensitive, whitespace-sensitive) as per user request.
*   **Disqualification & Replacement Relationship**: Disqualified CPB (Tidak Lolos) and their replacement (Pengganti) will be stored in a unified database table (`replacement_events`). This links the disqualified record directly with the replacement record on the same row, ensuring correct paired display in the table and exports.
*   **Validation Length**: `No. NIK` and `No. KK` must be exactly 16 characters.
*   **Duplicate Detection**: Blocks duplicate checks on `Nama`, `No. NIK`, `No. KK`, or any combinations thereof.
*   **Self-Identity Check**: `No. NIK` and `No. KK` on the same record must not be identical.

---

## Proposed Changes

We will build the application in a new directory `bsps-db-app` within the workspace.

### Backend System (Python)
A Python server handles the heavy lifting of spreadsheet analysis and report generation.

#### [NEW] [app.py](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/app.py)
The entrypoint of the FastAPI/Flask server providing endpoints for:
*   `/api/invers/upload`: Parses `INVERS.xlsx` sheets, maps them to a specific `Tahap`, checks for duplicates, and archives past revisions.
*   `/api/verified/upload`: Parses `TEMPLATE_TERVERIFIKASI.xlsx` sheets (`Lamp.IIA` and `Lamp.IIIA`), groups records into partial upload batches (e.g., Berita Acara), and validates NIK/KK string length, duplicates, and mismatch errors.
*   `/api/reconciliation`: Fetches discrepancies and updates database records based on manual corrections or user overrides.
*   `/api/export/excel`: Uses Python (`openpyxl`) to generate a highly stylized Excel sheet matching `EXPORT-FORMAT-EXAMPLE.pdf` styling (Bookman Old Style 12pt, black borders, signature areas, and multi-sheet output).

#### [NEW] [database.py](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/database.py)
Initializes SQLite database schemas:
*   `invers_records`: Stores individual CPB records with name, NIK, KK, stage (`Tahap`), revision number, status, and metadata.
*   `verified_records`: Stores verified records linked to a specific "Berita Acara" upload session, indicating whether they are `LOLOS` or `TIDAK LOLOS`, reasons for disqualification, coordinate data (`Latitude`, `Longitude`), and replacement details.
*   `replacement_events`: A table linking a disqualified CPB ID with their replacement CPB ID to maintain a strict 1-to-1 replacement pairing.
*   `upload_batches`: Keeps track of upload sessions (Berita Acara batches: Berita Acara Pertama, Kedua, etc.) pointing to an INVERS stage. Supports rollbacks by deleting records associated with a specific batch ID.

---

### Frontend System (React + Vite)
Built with CSS variables to ensure a sleek green-accented layout, rounded cards, left-side navigation panel, and a modern clean feel.

#### [NEW] [index.css](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/src/index.css)
Declares the global styling system:
*   Green theme accents (Forest green matching the logo and screenshot UI).
*   Clean typography using `Inter` or `Outfit` fonts.
*   Glassmorphism shadows, soft transitions, and hover states.

#### [NEW] [Dashboard.jsx](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/src/components/Dashboard.jsx)
Main dashboard screen showing metrics:
*   Active INVERS stages select list.
*   Status summary widgets: Total CPB, Verified Qualified (Lolos), Disqualified (Tidak Lolos), and Pending Mismatches.
*   "Upload Session Tracking" list displaying all uploaded Berita Acara batches with record counts, duplicate counts, and "Delete Batch" options.

#### [NEW] [DataList.jsx](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/src/components/DataList.jsx)
Dynamic data viewer showing INVERS and Verified Data side-by-side or in tabs:
*   Highlighting invalid/mismatched records with a red background.
*   Interactive tooltip detailing the specific failed validation check (e.g., "NIK must be 16 digits", "Duplicate found", "NIK and KK are identical").
*   Mismatched verified records include a quick-link button to navigate to the exact row on the INVERS list.

#### [NEW] [ReconciliationCenter.jsx](file:///c:/Users/Hidayat/Downloads/DATABASE%20EXCEL%20EXPORT/src/components/ReconciliationCenter.jsx)
Card-based interface showing discrepancies side-by-side:
*   **Expected (INVERS)** vs **Actual (Verified Data)** card comparisons.
*   Buttons to resolve disputes: "Accept Verification Data", "Edit Manually", or "Ignore and Keep Flagged".

---

## Verification Plan

### Automated Tests
*   **Upload Validation Suite**: Run Python tests using `pytest` to verify excel sheets parsing:
    *   Test validation of NIK / KK length (verify 16-character constraint).
    *   Test detection of duplicate records within the same sheet.
    *   Test rejection of uploads where NIK and KK are identical.
    *   Test batch deduplication logic (verifying that duplicate items in subsequent Berita Acara uploads are correctly flagged and skipped).

### Manual Verification
*   **Excel Upload Testing**: Upload the provided `INVERS.xlsx` and verified sheets. Verify that stages are recorded correctly, revisions are tracked, and matching comparison results behave exactly as expected.
*   **Rollback & Batch Check**: Perform partial uploads of verified records in multiple Berita Acara batches, check the duplicates statistics panel, delete one of the Berita Acara batches, and verify that only the target records are safely removed from the system.
*   **Excel Styled Export**: Export individual and combined Berita Acara files, open them in Microsoft Excel/Google Sheets, and verify that:
    *   The font is set to Bookman Old Style 12pt.
    *   The table borders are black.
    *   The formatting structure matches the visual layout shown in `EXPORT-FORMAT-EXAMPLE.pdf`.
