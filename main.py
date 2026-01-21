import os
import re
import json
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

from openpyxl import load_workbook
from rapidfuzz import fuzz

# PDF + DOCX parsing
from pypdf import PdfReader
from docx import Document


# -----------------------------
# Config / Models
# -----------------------------

@dataclass
class ClauseResult:
    clause_id: str
    status: str                 # "PASS" / "FAIL" / "REVIEW"
    confidence: float           # 0..1
    found_files: List[str]
    notes: str


@dataclass
class ClauseDefinition:
    clause_id: str
    title: str
    requirement_text: str
    keywords: List[str]         # baseline approach


class OptionalSemanticMatcher:
    """
    Optional semantic matcher using sentence-transformers.
    If not installed/available, it gracefully disables itself.
    """
    def __init__(self, model_name: str = "sentence-transformers/all-MiniLM-L6-v2"):
        self.enabled = False
        self.model = None
        self.util = None
        try:
            from sentence_transformers import SentenceTransformer, util
            self.model = SentenceTransformer(model_name)
            self.util = util
            self.enabled = True
        except Exception:
            self.enabled = False

    def similarity(self, a: str, b: str) -> float:
        if not self.enabled:
            return 0.0
        emb_a = self.model.encode(a, convert_to_tensor=True, normalize_embeddings=True)
        emb_b = self.model.encode(b, convert_to_tensor=True, normalize_embeddings=True)
        sim = float(self.util.cos_sim(emb_a, emb_b).item())
        # sim is typically in [-1,1], clamp to [0,1] for convenience
        return max(0.0, min(1.0, (sim + 1) / 2))


# -----------------------------
# Evidence ingestion
# -----------------------------

TEXT_EXTS = {".txt", ".md", ".log"}
DOCX_EXTS = {".docx"}
PDF_EXTS = {".pdf"}


def safe_read_text_file(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""


def extract_text_from_pdf(path: str) -> str:
    try:
        reader = PdfReader(path)
        pages = []
        for p in reader.pages:
            pages.append(p.extract_text() or "")
        return "\n".join(pages)
    except Exception:
        return ""


def extract_text_from_docx(path: str) -> str:
    try:
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    except Exception:
        return ""


def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in TEXT_EXTS:
        return safe_read_text_file(path)
    if ext in PDF_EXTS:
        return extract_text_from_pdf(path)
    if ext in DOCX_EXTS:
        return extract_text_from_docx(path)
    return ""


def list_files_recursive(root: str) -> List[str]:
    files = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            files.append(os.path.join(dirpath, fn))
    return files


def normalize_clause_id(s: str) -> str:
    """
    Normalizes IDs like 'A.5.1', 'a 5 1', 'A-5-1' => 'A.5.1'
    and '6.1.2' stays '6.1.2'
    """
    s = s.strip()
    s = s.replace("_", ".").replace("-", ".").replace(" ", ".")
    s = re.sub(r"\.+", ".", s)
    return s.upper()


def build_clause_folder_index(evidence_root: str) -> Dict[str, str]:
    """
    Map clause_id -> folder_path
    Assumes folder names correspond to clause/control IDs (e.g., 'A.5.1', '6.1.2')
    """
    index = {}
    for entry in os.listdir(evidence_root):
        p = os.path.join(evidence_root, entry)
        if os.path.isdir(p):
            cid = normalize_clause_id(entry)
            index[cid] = p
    return index


# -----------------------------
# NLP / Scoring
# -----------------------------

def keyword_coverage(text: str, keywords: List[str]) -> float:
    """
    Return ratio of keywords present (simple baseline).
    """
    if not keywords:
        return 0.0
    t = text.lower()
    hit = 0
    for kw in keywords:
        if kw.lower() in t:
            hit += 1
    return hit / len(keywords)


def fuzzy_requirement_match(text: str, requirement: str) -> float:
    """
    Fuzzy score for requirement presence (0..1).
    """
    if not requirement.strip():
        return 0.0
    # Use partial_ratio because evidence might contain a fragment
    score = fuzz.partial_ratio(requirement.lower(), text.lower()) / 100.0
    return score


def evidence_sufficiency_score(
    requirement_text: str,
    combined_evidence_text: str,
    keywords: List[str],
    semantic: OptionalSemanticMatcher
) -> Tuple[float, Dict[str, float]]:
    """
    Combine:
    - keyword coverage
    - fuzzy requirement match
    - optional semantic similarity
    into one final score 0..1
    """
    kw = keyword_coverage(combined_evidence_text, keywords)
    fz = fuzzy_requirement_match(combined_evidence_text, requirement_text)
    sem = semantic.similarity(requirement_text, combined_evidence_text[:4000]) if semantic.enabled else 0.0

    # Weighted blend (tune later)
    # Start conservative: keywords + fuzzy dominate; semantic adds small boost if enabled
    score = (0.45 * kw) + (0.45 * fz) + (0.10 * sem)

    details = {"keyword": kw, "fuzzy": fz, "semantic": sem, "final": score}
    return score, details


def classify(score: float, num_files: int) -> str:
    """
    Decision thresholds. Tune to your org.
    """
    if num_files == 0:
        return "FAIL"
    if score >= 0.70:
        return "PASS"
    if score >= 0.45:
        return "REVIEW"
    return "FAIL"


# -----------------------------
# Checklist IO (Excel)
# -----------------------------

def load_clauses_from_excel(
    checklist_xlsx: str,
    sheet_name: Optional[str] = None,
    col_map: Optional[Dict[str, str]] = None
) -> Tuple[List[ClauseDefinition], str, Dict[str, int], object]:
    """
    Reads checklist rows into ClauseDefinition list.
    Returns: (clauses, sheet_used, header_to_colidx, worksheet)
    """
    wb = load_workbook(checklist_xlsx)
    ws = wb[sheet_name] if sheet_name else wb.active

    # Default expected columns (edit if your file differs)
    # Map internal name -> header in sheet
    default_col_map = {
        "id": "id",
        "title": "title",
        "requirement_text": "requirement_text",
        "keywords": "keywords",  # optional column with comma-separated keywords
        "status": "status",
        "confidence": "confidence",
        "found_files": "found_files",
        "notes": "notes",
    }
    col_map = col_map or default_col_map

    # Read header row (assumes row 1)
    header = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if isinstance(v, str) and v.strip():
            header[v.strip().lower()] = col

    def col_idx(key: str) -> Optional[int]:
        h = col_map[key].lower()
        return header.get(h)

    required = ["id", "title", "requirement_text"]
    for r in required:
        if col_idx(r) is None:
            raise ValueError(f"Checklist missing required column header: '{col_map[r]}' (case-insensitive)")

    clauses: List[ClauseDefinition] = []
    for row in range(2, ws.max_row + 1):
        cid = ws.cell(row=row, column=col_idx("id")).value
        if cid is None:
            continue
        cid = normalize_clause_id(str(cid))

        title = ws.cell(row=row, column=col_idx("title")).value or ""
        req = ws.cell(row=row, column=col_idx("requirement_text")).value or ""

        kw_cell = col_idx("keywords")
        if kw_cell is not None:
            raw = ws.cell(row=row, column=kw_cell).value or ""
            keywords = [k.strip() for k in str(raw).split(",") if k.strip()]
        else:
            # Fallback: auto-keywords from requirement (very basic)
            keywords = auto_keywords_from_requirement(str(req))

        clauses.append(ClauseDefinition(clause_id=cid, title=str(title), requirement_text=str(req), keywords=keywords))

    return clauses, ws.title, header, wb


def auto_keywords_from_requirement(requirement: str) -> List[str]:
    """
    Very simple keyword generator. Replace with a real NLP keyword extractor later.
    """
    stop = {"the", "a", "an", "and", "or", "to", "of", "in", "for", "with", "on", "by", "is", "are", "be"}
    tokens = re.findall(r"[A-Za-z][A-Za-z0-9_\-]{2,}", requirement.lower())
    uniq = []
    for t in tokens:
        if t not in stop and t not in uniq:
            uniq.append(t)
    # keep top N
    return uniq[:12]


def write_results_to_excel(
    wb,
    sheet_name: str,
    header: Dict[str, int],
    results: Dict[str, ClauseResult],
    col_headers: Dict[str, str]
) -> None:
    """
    Writes results into rows by matching clause_id.
    """
    ws = wb[sheet_name]

    # Map internal -> actual header name in sheet (case-insensitive matching already in header dict)
    def find_col(header_name: str) -> Optional[int]:
        return header.get(header_name.lower())

    status_col = find_col(col_headers["status"])
    conf_col = find_col(col_headers["confidence"])
    files_col = find_col(col_headers["found_files"])
    notes_col = find_col(col_headers["notes"])

    # If output cols don't exist, create them at the end
    def ensure_col(hname: str) -> int:
        idx = find_col(hname)
        if idx is not None:
            return idx
        new_col = ws.max_column + 1
        ws.cell(row=1, column=new_col).value = hname
        header[hname.lower()] = new_col
        return new_col

    status_col = status_col or ensure_col(col_headers["status"])
    conf_col = conf_col or ensure_col(col_headers["confidence"])
    files_col = files_col or ensure_col(col_headers["found_files"])
    notes_col = notes_col or ensure_col(col_headers["notes"])

    # Find id column
    id_col = header.get(col_headers["id"].lower())
    if id_col is None:
        raise ValueError("Cannot find ID column to write results.")

    for row in range(2, ws.max_row + 1):
        cid = ws.cell(row=row, column=id_col).value
        if cid is None:
            continue
        cid = normalize_clause_id(str(cid))
        if cid not in results:
            continue
        r = results[cid]
        ws.cell(row=row, column=status_col).value = r.status
        ws.cell(row=row, column=conf_col).value = round(r.confidence, 3)
        ws.cell(row=row, column=files_col).value = ", ".join(r.found_files[:20])  # avoid huge cells
        ws.cell(row=row, column=notes_col).value = r.notes


# -----------------------------
# Main Audit Evaluator
# -----------------------------

def evaluate_clause(
    clause: ClauseDefinition,
    clause_folder: Optional[str],
    semantic: OptionalSemanticMatcher,
    min_text_chars: int = 400
) -> ClauseResult:
    if not clause_folder or not os.path.isdir(clause_folder):
        return ClauseResult(
            clause_id=clause.clause_id,
            status="FAIL",
            confidence=0.0,
            found_files=[],
            notes="No folder found for this clause/control ID."
        )

    all_files = list_files_recursive(clause_folder)
    # Filter to common evidence formats
    evidence_files = [f for f in all_files if os.path.splitext(f)[1].lower() in (TEXT_EXTS | DOCX_EXTS | PDF_EXTS)]

    texts = []
    used_files = []
    for f in evidence_files:
        t = extract_text(f)
        if t and t.strip():
            texts.append(t)
            used_files.append(os.path.relpath(f, clause_folder))

    combined = "\n\n".join(texts).strip()
    if len(combined) < min_text_chars:
        # Might be screenshots-only / images-only; mark review
        base_status = "REVIEW" if used_files else "FAIL"
        return ClauseResult(
            clause_id=clause.clause_id,
            status=base_status,
            confidence=0.30 if used_files else 0.0,
            found_files=used_files,
            notes="Low extractable text. Evidence may be image-based/scanned; manual review recommended."
        )

    score, details = evidence_sufficiency_score(
        requirement_text=clause.requirement_text,
        combined_evidence_text=combined,
        keywords=clause.keywords,
        semantic=semantic
    )
    status = classify(score, len(used_files))

    notes = (
        f"Scores => keyword={details['keyword']:.2f}, fuzzy={details['fuzzy']:.2f}, "
        f"semantic={details['semantic']:.2f}, final={details['final']:.2f}. "
        f"Matched {len(used_files)} file(s)."
    )

    return ClauseResult(
        clause_id=clause.clause_id,
        status=status,
        confidence=float(max(0.0, min(1.0, score))),
        found_files=used_files,
        notes=notes
    )


def run_audit(
    evidence_root: str,
    checklist_xlsx: str,
    output_xlsx: str,
    sheet_name: Optional[str] = None,
    use_semantic: bool = True
) -> None:
    # Load checklist
    col_headers = {
        "id": "id",
        "title": "title",
        "requirement_text": "requirement_text",
        "keywords": "keywords",
        "status": "status",
        "confidence": "confidence",
        "found_files": "found_files",
        "notes": "notes",
    }
    clauses, used_sheet, header, wb = load_clauses_from_excel(
        checklist_xlsx=checklist_xlsx,
        sheet_name=sheet_name,
        col_map=col_headers
    )

    # Build folder index
    index = build_clause_folder_index(evidence_root)

    # NLP matcher
    semantic = OptionalSemanticMatcher() if use_semantic else OptionalSemanticMatcher(model_name="__disabled__")
    if use_semantic and not semantic.enabled:
        print("[WARN] sentence-transformers not available. Running without semantic similarity.")

    # Evaluate all clauses
    results: Dict[str, ClauseResult] = {}
    for c in clauses:
        folder = index.get(c.clause_id)
        results[c.clause_id] = evaluate_clause(c, folder, semantic)

    # Write results back
    write_results_to_excel(
        wb=wb,
        sheet_name=used_sheet,
        header=header,
        results=results,
        col_headers=col_headers
    )
    wb.save(output_xlsx)

    # Optional: also dump JSON summary
    summary_path = os.path.splitext(output_xlsx)[0] + "_summary.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({k: results[k].__dict__ for k in results}, f, indent=2)

    print(f"Saved: {output_xlsx}")
    print(f"Saved: {summary_path}")


# -----------------------------
# Example usage
# -----------------------------
if __name__ == "__main__":
    run_audit(
        evidence_root=r"./evidence_root",
        checklist_xlsx=r"./iso27001_checklist.xlsx",
        output_xlsx=r"./iso27001_checklist_filled.xlsx",
        sheet_name=None,          # or "Sheet1"
        use_semantic=True
    )
