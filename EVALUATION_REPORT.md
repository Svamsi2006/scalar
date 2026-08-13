# PII Redaction Tool - Performance Evaluation Report

## Testing Methodology

The tool was evaluated using a **controlled test corpus** containing exactly **30 known PII items** across all 9 required categories, embedded in realistic corporate text. The evaluation process:

1. **Ground Truth Construction:** A multi-paragraph text was created with precisely counted PII items, each annotated with its type and exact text span.
2. **Detection Execution:** The `DetectionEngine` (hybrid regex + spaCy NER) was run on the test text.
3. **Matching Algorithm:** For each PII category, detected entities were compared against ground truth using **substring matching** — a detection is a True Positive if its value contains the ground truth value or vice versa, with matching entity types.
4. **Metric Computation:**
   - `Precision = TP / (TP + FP)` — What fraction of detections are correct?
   - `Recall = TP / (TP + FN)` — What fraction of actual PII was caught?
   - `Accuracy = (TP + TN) / (TP + TN + FP + FN)` — Overall correctness (TN estimated from total word count minus TP, FP, FN)

---

## Performance Metrics Table

| PII Category     | TP | FP | FN | Precision (%) | Recall (%) | Accuracy (%) |
|---|---|---|---|---|---|---|
| Full Names       | 5  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| Emails           | 4  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| Phone Numbers    | 4  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| Companies        | 4  | 2  | 0  | 66.67%         | 100.00%    | 98.88%       |
| Addresses        | 2  | 2  | 1  | 50.00%         | 66.67%     | 98.32%       |
| SSNs             | 2  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| Credit Cards     | 2  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| Dates            | 4  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| IP Addresses     | 2  | 0  | 0  | 100.00%        | 100.00%    | 100.00%      |
| **OVERALL**      | **29** | **4** | **1** | **87.88%** | **96.67%** | **99.69%** |

---

## Metric Interpretations & Key Insights

### Recall: 96.67%

The system catches nearly all PII instances. The single False Negative was a **complex multi-line address** (`1600 Pennsylvania Avenue NW, Washington DC 20500`) — the spaCy `en_core_web_sm` model recognized parts of it (e.g., "Washington DC") as `GPE` but did not capture the full street-level address as a single entity.

**Why this happens:** NER models treat addresses as composites of multiple entity types (street names, cities, states), and our matching requires the full address text to be covered. In practice, partial redaction still removes the most identifying portions.

### Precision: 87.88%

The 4 False Positives break down as:
- **2 Company FPs:** spaCy tagged contextual phrases like "HR" abbreviations or descriptive terms adjacent to real company names as ORG entities.
- **2 Address FPs:** Location names used in generic context (e.g., standalone city names like "Pune", "London") were tagged as addresses when they appeared outside of actual address strings.

**Why this is acceptable:** In a security-critical application, higher recall (catching more PII) is preferred over higher precision (fewer false redactions). False positives mean we redact slightly more than necessary, which is the safer error mode.

### Accuracy: 99.69%

The extremely high accuracy reflects the fact that PII entities are sparse relative to the total text — most words are correctly identified as non-PII (True Negatives).

---

## Category-Specific Analysis

### Perfect Detection (100% Precision & Recall)
- **Full Names:** spaCy's PERSON entity recognizer correctly identified all 5 names, including multi-part Indian names (e.g., "Rajesh Kumar Sharma") and shorter names ("Anita Rao"). The 2-word minimum filter prevented common nouns from being flagged.
- **Emails:** The regex pattern `[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}` caught all formats including `.co.uk` and `.in` TLDs.
- **Phone Numbers:** All 4 formats were detected — Indian mobile (+91 98765 43210), Indian landline (+91 20 4505 3237), US format (+1 555-867-5309), and STD code (022-25678901). The digit-count validation (7-15 digits) prevented random numbers from being flagged.
- **SSNs, Credit Cards, IP Addresses:** Regex-only detection with strong pattern boundaries achieved perfect results.
- **Dates:** Multi-format regex (DD/MM/YYYY, Month DD YYYY, DD-Mon-YYYY, ISO) caught all 4 test dates.

### Areas for Improvement
- **Addresses:** Consider adding regex-based address detection patterns (e.g., matching "No./Plot/Tower + number + area name + city + PIN") alongside NER to improve recall for Indian address formats.
- **Companies:** A custom dictionary of known company suffixes (Ltd, Pvt, Inc, LLP, Corp) could be used as a post-processing filter to reduce ORG false positives.

---

## Optimization Recommendations

| Optimization | Expected Impact | Effort |
|---|---|---|
| Use `en_core_web_trf` (transformer model) | +5-10% precision for names/orgs | Low (pip install) |
| Custom company suffix whitelist | -50% company FPs | Low |
| Indian address regex patterns | +20% address recall | Medium |
| Document-specific name seed list | +5% name recall for rare names | Medium |
| Fine-tune NER on Indian legal documents | +10-15% across all NER categories | High |

---

## Consistency Validation

The `ReplacementManager` was separately tested to confirm:
- **Same input → same output:** "Rajesh Kumar Sharma" mapped to "Allison Hill" consistently across multiple calls. ✓
- **Different inputs → different outputs:** "Priya Anand Mehta" mapped to "Noah Rhodes" (distinct from above). ✓
- **Normalization:** Whitespace and case variations in the same name resolve to the same fake. ✓
