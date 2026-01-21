# NLP Based Automated Audit Evaluation System (ISO/IEC 27001)

This project is a **Python-based NLP system** that automatically evaluates **ISO/IEC 27001 audit evidence** by scanning organized evidence folders and auto-filling an audit checklist (Excel).

It supports the standard ISO 27001 structure:

* **Clauses 4–10** (ISMS requirements)
* **Annex A controls** (ISO 27001:2022)

  * A.5 Organizational Controls
  * A.6 People Controls
  * A.7 Physical Controls
  * A.8 Technological Controls

---

## Recommended Evidence Folder Structure

Your evidence folder should be organized by ISO 27001 clause/control IDs.

Example:

```bash
evidence_root/
│
├── Clause 4/
│   ├── 4.1 Context of the organization/
│   │   ├── context_document.pdf
│   │   └── meeting_notes.docx
│
├── Clause 5/
│   ├── 5.2 Information Security Policy/
│   │   ├── is_policy.pdf
│
└── Annex A/
    ├── A.5 Organizational controls/
    │   ├── A.5.1 Policies for information security/
    │   │   ├── policy.pdf
    │   │   └── approval_email.txt
    │
    └── A.8 Technological controls/
        ├── A.8.9 Configuration management/
        │   ├── hardening_guide.docx
        │   ├── firewall_rules.pdf

'''
```

✅ The code automatically detects IDs inside folder names like:

* '4.1'
* '6.1.2'
* 'A.5.1'
* 'A.8.34'


## How the NLP Evaluation Works

The engine checks evidence sufficiency using:

### 1) Keyword Coverage

Measures how many keywords appear in evidence.

### 2) Fuzzy Matching (Text Similarity)

Uses fuzzy match between requirement text and extracted evidence text.

### 3) Semantic Similarity (Optional)

Uses embeddings from **sentence-transformers** for semantic matching.

Final score is calculated as:

'''
Score = 0.45(keyword) + 0.45(fuzzy) + 0.10(semantic)
'''


## ⚙️ Installation

### 1) Create a virtual environment (recommended)

'''bash
python -m venv venv
source venv/bin/activate     # Linux/Mac
venv\Scripts\activate        # Windows
'''

### 2) Install dependencies

'''bash
pip install openpyxl rapidfuzz pypdf python-docx
'''

### Optional: enable semantic similarity

'''bash
pip install sentence-transformers torch
'''

---

## How to Run

Edit the script’s main block:

'''python
run_audit_iso27001(
    evidence_root=r"./evidence_root",
    checklist_xlsx=r"./iso27001_checklist.xlsx",
    output_xlsx=r"./iso27001_checklist_filled.xlsx",
    sheet_name=None,
    use_semantic=True
)
'''

Run:

'''bash
python main.py
'''

---

## 📌 Output Files

After execution, the tool generates:

### ✅ 1) Filled Checklist

'''
iso27001_checklist_filled.xlsx
'''

### ✅ 2) JSON Summary Report

'''
iso27001_checklist_filled_summary.json
'''

The JSON report includes:

* control/clause id
* status
* confidence score
* detected evidence files
* notes

---

## ✅ Auto-Generating a Template Checklist (ISO Structure)

If no checklist file is found, the tool can generate a **template Excel**:

* Clauses 4–10
* Annex A controls (A.5 → A.8)

Then you can manually fill 'requirement_text' for each control.


## 🛑 Important Notes 

### 1) Scanned PDFs & Images

If evidence is scanned or image-based, text extraction may fail.
You should add OCR support (Tesseract / EasyOCR) for full automation.

### 2) Requirement Text Must Exist

For best scoring accuracy, each control should have a meaningful 'requirement_text'.

### 3) Folder Naming Must Contain IDs

Evidence folder names must include ISO IDs (e.g., 'A.5.1', '6.1.2', etc.)


Example Project Structure

project/
│
├── main.py
├── iso27001_checklist.xlsx
├── evidence_root/
│
└── outputs/
    ├── iso27001_checklist_filled.xlsx
    └── iso27001_checklist_filled_summary.json


