# PII Redaction Tool - Technical Documentation

## Overview

A production-grade, context-preserving Personally Identifiable Information (PII) redaction tool for MS Word (`.docx`) documents. Instead of blacking out sensitive data, this tool replaces real PII with **realistic, contextually consistent fake alternatives** — maintaining the same fake identity for each individual throughout the entire document.

Built to handle the complexity of corporate prospectuses: nested tables, multi-run paragraph formatting, headers, footers, and thousands of PII instances.

---

## Design & Approach

### Technology Stack

| Component | Technology | Purpose |
|---|---|---|
| Document Parsing | `python-docx` | Read/write Word XML, traverse paragraphs, tables, headers, footers |
| Structured PII Detection | Python `re` (regex) | Emails, phone numbers, SSNs, credit cards, IP addresses, dates |
| Context-Based PII Detection | `spaCy` NER (`en_core_web_sm`) | Full names, company names, physical addresses |
| Fake Data Generation | `Faker` | Locale-aware, seeded fake data with format-matching |
| CLI Interface | `argparse` | Lightweight command-line orchestration |

### Architecture

```
pii_redactor/
├── document_parser.py      # DOCX ingestion, run-splitting replacement engine
├── detection_engine.py     # Hybrid regex + spaCy NER detection pipeline
├── replacement_manager.py  # Stateful PII→fake mapping with Faker
└── redactor.py             # CLI orchestrator tying all modules together

test_redactor.py            # Validation suite with precision/recall metrics
```

**Data flow:** `Load DOCX → Extract paragraphs → Detect PII (full text) → Map to fakes → Replace in runs (format-preserving) → Save DOCX`

---

## Technical Trade-offs

### Regex vs. NER

| Aspect | Regex | NER (spaCy) |
|---|---|---|
| **Used for** | Structured patterns: emails, phones, SSNs, credit cards, IPs, dates | Context-dependent entities: names, organizations, addresses |
| **Strengths** | Deterministic, fast, zero false negatives for well-formed patterns | Handles variations, typos, and context clues |
| **Weaknesses** | Cannot detect unstructured PII like names | Lower precision for short or ambiguous entities |
| **Why both?** | Regex catches 100% of pattern-based PII; NER catches entities that have no fixed pattern. Together they cover all 9 categories. |

### Run-Splitting in `python-docx` — The Core Challenge

**Problem:** Word internally splits text across multiple XML `<w:r>` (Run) elements. A name like "Kushal Hegde" might be stored as:
```xml
<w:r><w:rPr><w:b/></w:rPr><w:t>Kushal </w:t></w:r>
<w:r><w:rPr><w:i/></w:rPr><w:t>Hegde</w:t></w:r>
```
A naive per-run regex will **never find "Kushal Hegde"** because no single run contains the full name.

**Solution — Paragraph-Level Offset Mapping:**
1. **Concatenate** all run texts to reconstruct the full paragraph string.
2. **Detect PII** on this full string, obtaining character offsets `(start, end)`.
3. **Map offsets back to runs** using a `(run_index, char_start, char_end)` mapping.
4. **Replace right-to-left** (descending start offset) to prevent offset invalidation.
5. For cross-run matches: replace text in the first affected run, clear text in subsequent affected runs, preserving all font formatting.

### Contextual Consistency

The `ReplacementManager` maintains an in-memory dictionary: `{entity_type: {normalized_real_value: fake_value}}`. Values are normalized (lowercased, whitespace-collapsed) before lookup, so "Kushal  Hegde", "kushal hegde", and "Kushal Hegde" all map to the same fake name.

---

## Installation & Usage

### Prerequisites
- Python 3.8+
- pip

### Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Download spaCy English model
python -m spacy download en_core_web_sm
```

### Running the Redactor
```bash
# Basic usage
python -m pii_redactor.redactor --input "Red Herring Prospectus.docx" --output "Redacted_Prospectus.docx"

# With options
python -m pii_redactor.redactor \
  --input "Red Herring Prospectus.docx" \
  --output "Redacted_Prospectus.docx" \
  --seed 42 \
  --exclude SSN,CREDIT_CARD \
  --verbose
```

### Running the Test Suite
```bash
python test_redactor.py
```

### CLI Arguments

| Argument | Short | Required | Default | Description |
|---|---|---|---|---|
| `--input` | `-i` | Yes | — | Path to input `.docx` file |
| `--output` | `-o` | Yes | — | Path for output redacted `.docx` |
| `--exclude` | `-e` | No | None | Comma-separated PII types to skip |
| `--seed` | `-s` | No | 42 | Faker random seed for reproducibility |
| `--verbose` | `-v` | No | False | Enable debug-level logging |

---

## Observed False Positives / Negatives

### Known False Positives
- **Common capitalized words as names:** spaCy may tag proper nouns like "Board", "Chairman", or "Director" as PERSON entities when they appear in title-case contexts.
- **Section/page numbers as phones:** Long numeric sequences in financial tables (e.g., large monetary amounts, CIN numbers) may partially match phone regex patterns.
- **PIN codes as dates:** Indian PIN codes (6-digit numbers) may be caught by date patterns in certain contexts.
- **Standalone city/country names as addresses:** spaCy tags "India", "Mumbai", "Pune" as GPE, which are classified as ADDRESS even when used in generic context.

### Known False Negatives
- **Names not in spaCy's vocabulary:** Uncommon Indian names without standard English capitalization may be missed by the `en_core_web_sm` model.
- **Addresses without named entities:** Addresses that consist purely of numbers and street descriptors without a city/state name may not be caught by NER.
- **Emails with unusual TLDs:** Very new or uncommon TLDs (e.g., `.crypto`, `.internal`) might be missed by the email regex.
- **Dates in prose format:** "the fifteenth of March" (spelled-out dates) are not detected.

### Mitigation Strategies
1. Use `en_core_web_trf` (transformer-based) spaCy model for higher NER accuracy.
2. Add a custom dictionary of known names from the document for pre-seeding.
3. Expand regex patterns iteratively based on false negative analysis.

---

## Extending to New PII Types

To add a new PII category (e.g., `PASSPORT_NUMBER`):

1. **`detection_engine.py`**: Add a regex pattern to `self._patterns` dict.
2. **`replacement_manager.py`**: Add a `_generate_passport` method and register it in `self._generators`.
3. No changes needed in `document_parser.py` or `redactor.py` — the pipeline is fully generic.

---

## Project Structure

```
scaler/
├── Red Herring Prospectus.docx         # Input document
├── Redacted_Prospectus.docx            # Output (generated)
├── requirements.txt                     # Python dependencies
├── README.md                            # This file
├── EVALUATION_REPORT.md                 # Performance metrics
├── test_redactor.py                     # Validation suite
└── pii_redactor/
    ├── __init__.py
    ├── document_parser.py              # DOCX parser & run-splitting engine
    ├── detection_engine.py             # Regex + spaCy NER detector
    ├── replacement_manager.py          # Stateful fake data mapper
    └── redactor.py                     # CLI orchestrator
```
