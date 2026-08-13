# PII Redactor - Comprehensive Code Audit & QA Report

This report provides a detailed files-by-files structural audit, QA review, and dependency validation of the modular Personally Identifiable Information (PII) Redaction Tool. It serves as a professional handoff document when sharing the project with a company or stakeholder.

---

## 📂 1. File-by-File Code Audit

### 📄 `app.py` (Streamlit Web Interface)
- **Role:** Orchestrates the user-facing web dashboard, handles in-memory file uploads/downloads, displays live metrics, and provides an interactive audit log.
- **Audit Findings:**
  - **In-Memory Buffering:** High-security design. Files are read and written using `io.BytesIO`, ensuring no sensitive data is written to the server's local storage.
  - **Download Fix:** Custom client-side HTML5 `<a download>` block with base64 encoding prevents browser download managers from dropping filenames or downloading files with UUID slugs.
  - **Redundancies:** No duplicate code blocks or unused imports found.
  - **Aesthetics:** Clean, enterprise-themed interface styling using customized CSS pills, metrics containers, and structured tabs.

### 📄 `pii_redactor/detection_engine.py` (PII Detector)
- **Role:** Evaluates input text using a hybrid framework: regular expressions for structured patterns and spaCy NER for unstructured natural language.
- **Audit Findings:**
  - **spaCy Load Resolution:** Updated `__init__` with a fallback mechanism that runs `sys.executable -m spacy download` dynamically via `subprocess` if `spacy.load` fails.
  - **Conflict Resolution (`_resolve_overlaps`):** Correctly resolves overlapping spans (e.g. preventing a phone number from being partially redacted as a date or name).
  - **Regex Limitations:** The date regex matches standard ISO, slash, and hyphenated numeric formats but doesn't catch fully-spelled textual dates (e.g., "fifteenth of July"). This is documented in the F/P and F/N section.

### 📄 `pii_redactor/replacement_manager.py` (PII Stateful Mapping)
- **Role:** Stateful manager assigning consistent replacements to matching PII entities.
- **Audit Findings:**
  - **Redaction Styles:** Successfully supports:
    1. `token` (Default): `[NAME_1]`, `[COMPANY_1]` (non-random, preserves relationships).
    2. `mask`: Standardized compliance tags `[REDACTED: FULL_NAME]`.
    3. `blackout`: Generic blackout `[REDACTED_FULL_NAME]`.
    4. `synthetic`: seeded, locale-aware fake data via Faker.
  - **Redundancy:** Unused import `from datetime import datetime` was detected. It is non-breaking but can be cleaned up.
  - **Consistent Normalization:** All inputs are stripped, lowercased, and space-collapsed before mapping, preventing identical names with minor spacing differences from getting different tags.

### 📄 `pii_redactor/document_parser.py` (DOCX Parser)
- **Role:** Handles low-level DOCX file loading, iteration of elements (paragraphs, tables, headers, footers), and run-splitting replacements.
- **Audit Findings:**
  - **Nested Table Parsing:** Recursive method `_extract_table_paragraphs` is robust and catches all paragraphs inside tables of arbitrary depth.
  - **Header/Footer Coverage:** Accurately scans `header`, `first_page_header`, `even_page_header` for all document sections.
  - **Run-Splitting Algorithm:** High-complexity format-safe replacement logic. It reconstructs paragraph-level text, applies replacements right-to-left to maintain offsets, and preserves the style parameters of the first affected run.

### 📄 `pii_redactor/redactor.py` (CLI orchestrator)
- **Role:** Tying everything together for command-line execution.
- **Audit Findings:**
  - **Exposed Mode Parameter:** Added a `-m` / `--mode` CLI argument to let command-line users select their desired redaction style, aligning the CLI capability with the web interface.
  - **Logging:** Properly manages runtime verbose logs for debugging.

---

## 🛠️ 2. Dependency & Deployment Analysis

### 📦 `requirements.txt`
- **Configuration:**
  - `streamlit>=1.30.0`
  - `python-docx>=1.1.0`
  - `spacy>=3.7.0`
  - `faker>=20.0.0`
  - `pandas>=1.4.0`
- **Why we use PEP 508 Direct Reference for spaCy model:**
  - Declaring a raw wheel URL (like `https://github.com/.../en_core_web_sm.whl`) without a package name description (e.g. `package @ url`) causes modern `pip` versions to fail.
  - Attempting to dynamically download and write the spaCy model to the system package directory (`site-packages`) at runtime results in `[Errno 13] Permission denied` errors on Streamlit Community Cloud due to read-only container file system permissions.
  - Adding the model using PEP 508 format (`en_core_web_sm @ https://github.com/explosion/...`) instructs Streamlit's build container to install the package at build time (with root permission) so that it is pre-loaded and accessible read-only at runtime.

---

## 📈 3. Strategic Recommendations for Shareability

If you are sharing this project with a company, the following enhancements will elevate the project to enterprise-grade:

1. **Whitelisting Specific Non-Sensitive Entities:**
   - **Feature:** Let users upload a whitelist of words/entities that should *never* be redacted (e.g. the company's own name, specific industry terms, or public figures).
2. **Context-Aware NER Thresholding:**
   - **Feature:** Adjust spaCy NER threshold confidence levels to prevent common nouns (like "Director" or "Chairman") from triggering PERSON false positives.
3. **Multi-File Bulk Processing:**
   - **Feature:** Expand the Streamlit interface to support uploading a ZIP archive of multiple DOCX files and generating a download archive of redacted files in bulk.
