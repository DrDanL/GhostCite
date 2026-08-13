import io
import os
import re
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timezone

import requests
from docx import Document as DocxDocument
from flask import Flask, render_template, request, send_file
from pypdf import PdfReader
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

# Disk-based store for summary reports, shared across gunicorn workers.
_STORE_DIR = os.path.join(tempfile.gettempdir(), "ghostcite_reports")
os.makedirs(_STORE_DIR, exist_ok=True)

# Maximum number of stored files and max age in seconds (1 hour)
_STORE_MAX_FILES = 50
_STORE_MAX_AGE = 3600


def _cleanup_store() -> None:
    """Remove old files from the store to prevent disk bloat."""
    try:
        files = []
        for name in os.listdir(_STORE_DIR):
            path = os.path.join(_STORE_DIR, name)
            if os.path.isfile(path):
                files.append((path, os.path.getmtime(path)))

        # Remove files older than max age
        now = time.time()
        for path, mtime in files:
            if now - mtime > _STORE_MAX_AGE:
                os.unlink(path)

        # If still too many, remove oldest
        files = [(p, m) for p, m in files if os.path.exists(p)]
        if len(files) > _STORE_MAX_FILES:
            files.sort(key=lambda x: x[1])
            for path, _ in files[: len(files) - _STORE_MAX_FILES]:
                os.unlink(path)
    except OSError:
        pass


def _store_report(download_id: str, data: bytes) -> None:
    """Save a summary report to disk."""
    _cleanup_store()
    path = os.path.join(_STORE_DIR, download_id + ".pdf")
    with open(path, "wb") as f:
        f.write(data)


def _get_report_path(download_id: str) -> str | None:
    """Return a report path when the ID is valid and the file still exists."""
    try:
        safe_id = str(uuid.UUID(download_id))
    except ValueError:
        return None
    path = os.path.join(_STORE_DIR, safe_id + ".pdf")
    # Guard against path traversal
    if not os.path.realpath(path).startswith(os.path.realpath(_STORE_DIR)):
        return None
    if os.path.isfile(path):
        return path
    return None


@dataclass
class ReferenceResult:
    reference: str
    matched: bool
    title: str | None = None
    doi: str | None = None
    partial_match: bool = False
    partial_match_details: str | None = None


DOI_PATTERN = re.compile(r"10\.\d{4,9}/[^\s,;\"'>\])}]+", re.IGNORECASE)

# Patterns that indicate the start of a new reference entry
REF_START_PATTERNS = [
    re.compile(r"^\[\d+\]"),           # [1], [23]
    re.compile(r"^\d{1,3}\.\s"),       # 1. , 23.
    re.compile(r"^\d{1,3}\)\s"),       # 1) , 23)
]

# Harvard-style reference start: "Surname, I. ... (Year)"
# Captures the position of author-year patterns for splitting.
_AUTHOR_YEAR_RE = re.compile(
    r"("
    # Standard author: Surname, I. [co-authors] (Year) or Surname, Firstname. [co-authors] (Year)
    # Allow multi-word surnames (e.g. "Haliloğlu Kahraman") and spaces before
    # periods in initials (PDF extraction artifact: "Y ." instead of "Y.").
    r"[A-Z][a-zA-Zà-öø-ÿÀ-ÖØ-Ý'\u011e\u011f\-]+"   # first part of surname
    r"(?:\s+[A-Z][a-zA-Zà-öø-ÿÀ-ÖØ-Ý'\u011e\u011f\-]+)*"  # optional extra surname words
    r",\s*[A-Z][a-zA-Zà-öø-ÿ]*\s*\."                 # initial (A .) or first name (Marcin.)
    r"[^(]{0,300}?"
    r"\(\d{4}[a-z]?(?:,\s*[A-Za-z]+\.?)?\)"           # (Year) or (2025, June)
    r"|"
    # Organisation / institutional author (no initials): Org Name. (Year)
    r"(?:HM\s+Government|NHS\s+England|Mental\s+Health\s+Foundation|World\s+Health\s+Organization)"
    r"[^(]{0,100}?"
    r"\(\d{4}[a-z]?\)"
    r")"
)

REFERENCE_HEADINGS = {
    "references",
    "bibliography",
    "works cited",
    "literature cited",
    "reference list",
}


def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def extract_text_from_docx(file_bytes: bytes) -> str:
    """Extract text from a Word (.docx) file."""
    doc = DocxDocument(io.BytesIO(file_bytes))
    return "\n".join(paragraph.text for paragraph in doc.paragraphs)


def _strip_page_number(line: str) -> str:
    """Remove a leading page number like '51  ' from a line."""
    return re.sub(r"^\d{1,4}\s{2,}", "", line)


def find_reference_section(text: str) -> str:
    """Extract only the reference section text from the full document text.

    Handles cases where the heading and references are on the same line
    (common when PDF pages are extracted as single long strings).
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    # Search from the end of the document backwards to find the *last*
    # occurrence of a reference heading (avoids matching table of contents).
    ref_start_idx = None
    ref_remainder = ""
    for i in range(len(lines) - 1, -1, -1):
        cleaned = _strip_page_number(lines[i]).strip()
        lower = cleaned.lower()
        # Check if the line starts with a reference heading
        for heading in REFERENCE_HEADINGS:
            if lower == heading or lower.startswith(heading + " "):
                ref_start_idx = i
                # Text after the heading on the same line
                if lower.startswith(heading + " "):
                    ref_remainder = cleaned[len(heading):].strip()
                break
        if ref_start_idx is not None:
            break

    if ref_start_idx is None:
        return ""

    # Build the reference section text
    parts: list[str] = []
    if ref_remainder:
        parts.append(ref_remainder)

    # Collect subsequent lines until we hit appendices or end
    appendix_pattern = re.compile(r"^\s*(?:\d{1,4}\s{2,})?(?:appendix|appendices)\b", re.IGNORECASE)
    for line in lines[ref_start_idx + 1:]:
        stripped = line.strip()
        if appendix_pattern.match(stripped):
            break
        # Strip leading page numbers
        stripped = _strip_page_number(stripped)
        if stripped:
            parts.append(stripped)

    return " ".join(parts)


def is_ref_start(line: str) -> bool:
    """Check if a line looks like the beginning of a new reference."""
    for pattern in REF_START_PATTERNS:
        if pattern.match(line):
            return True
    return False


def _split_harvard_references(text: str) -> list[str]:
    """Split a block of text into individual Harvard-style references.

    Uses the pattern 'Surname, I. ... (Year)' to detect where each new
    reference begins, then splits accordingly.  Only accepts a match as a
    genuine reference start when it sits at position 0 or is preceded by
    sentence-ending punctuation (period, closing paren, digit) — this avoids
    splitting on author names that appear mid-sentence in a title.
    """
    # Collect candidate split positions
    starts: list[int] = []

    for m in _AUTHOR_YEAR_RE.finditer(text):
        pos = m.start()
        if pos == 0:
            starts.append(pos)
            continue
        # Check character immediately before this match (skip whitespace)
        before = text[:pos].rstrip()
        if not before:
            continue
        last_char = before[-1]
        # Accept sentence-ending punctuation, digits, or URL/DOI endings
        if (last_char in ".!?)0123456789"
                or re.search(r"https?://\S+$", before) is not None
                or re.search(r"10\.\d{4,}/\S+$", before) is not None):
            starts.append(pos)

    if not starts:
        return [text] if len(text) > 30 else []

    references: list[str] = []
    for i, s in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        ref = text[s:end].strip()
        if ref:
            references.append(ref)

    return references


def extract_references(text: str) -> list[str]:
    """Extract references from text, handling Harvard and numbered formats."""

    # First, try to isolate the reference section
    ref_text = find_reference_section(text)

    if not ref_text:
        # Fall back: use the whole text and look for numbered references
        ref_text = text

    # Try Harvard-style splitting first
    references = _split_harvard_references(ref_text)

    # If Harvard splitting found very few, also try numbered patterns
    if len(references) <= 3:
        lines = [line.strip() for line in ref_text.splitlines() if line.strip()]
        numbered_refs: list[str] = []
        current_ref = ""
        for line in lines:
            if is_ref_start(line):
                if current_ref:
                    numbered_refs.append(current_ref.strip())
                current_ref = line
            elif current_ref:
                current_ref += " " + line
            else:
                current_ref = line
        if current_ref:
            numbered_refs.append(current_ref.strip())
        if len(numbered_refs) > len(references):
            references = numbered_refs

    # Filter out very short entries and duplicates
    seen: set[str] = set()
    filtered: list[str] = []
    for r in references:
        r = r.strip()
        if len(r) > 30 and r not in seen:
            seen.add(r)
            filtered.append(r)

    return filtered if filtered else references[:30]


# --- Rate limiting for Crossref API ---
# Crossref polite pool allows ~50 req/s with a mailto in User-Agent,
# but we add a conservative per-request delay and honour 429 responses.
CROSSREF_HEADERS = {"User-Agent": "ghostcite/0.2 (mailto:ghostcite@example.com)"}
CROSSREF_TIMEOUT = 15
CROSSREF_MIN_INTERVAL = 0.1  # seconds between requests (10 req/s max)
CROSSREF_MAX_RETRIES = 3

_last_request_time: float = 0.0
_rate_limit_lock = threading.Lock()


def _rate_limited_get(url: str, params: dict | None = None) -> requests.Response:
    """Make a GET request to Crossref respecting rate limits and 429 back-off."""
    global _last_request_time

    for attempt in range(CROSSREF_MAX_RETRIES):
        # Enforce minimum interval between requests (thread-safe)
        with _rate_limit_lock:
            elapsed = time.time() - _last_request_time
            if elapsed < CROSSREF_MIN_INTERVAL:
                time.sleep(CROSSREF_MIN_INTERVAL - elapsed)
            _last_request_time = time.time()

        response = requests.get(
            url, params=params, headers=CROSSREF_HEADERS, timeout=CROSSREF_TIMEOUT
        )

        if response.status_code == 429:
            # Respect Retry-After header if present, otherwise exponential back-off
            retry_after = response.headers.get("Retry-After")
            if retry_after:
                wait = min(float(retry_after), 60)
            else:
                wait = 2 ** attempt
            time.sleep(wait)
            continue

        return response

    # Return last response even if still 429 after retries
    return response


def _normalise(text: str) -> str:
    """Lower-case and collapse whitespace for fuzzy comparison."""
    return re.sub(r"\s+", " ", text.lower().strip())


def _check_partial_match(reference: str, crossref_title: str | None,
                         crossref_authors: list[str] | None = None) -> tuple[bool, str | None]:
    """Compare a reference string against Crossref metadata and detect partial matches.

    Returns (is_partial, detail_string).
    A partial match means *some* but not *all* key metadata aligns.
    """
    if not crossref_title:
        return False, None

    ref_norm = _normalise(reference)
    title_norm = _normalise(crossref_title)

    # Split title into significant words (≥4 chars to skip articles/prepositions)
    title_words = [w for w in title_norm.split() if len(w) >= 4]
    if not title_words:
        return False, None

    matching_words = sum(1 for w in title_words if w in ref_norm)
    title_match_ratio = matching_words / len(title_words) if title_words else 0

    # Full title match — not partial
    if title_match_ratio >= 0.8:
        return False, None

    details: list[str] = []

    # Partial title match: some words match but not enough for a full match
    if 0.3 <= title_match_ratio < 0.8:
        pct = round(title_match_ratio * 100)
        details.append(f"title partially matches ({pct}% of keywords)")

    # Check author overlap if authors provided
    if crossref_authors:
        authors_norm = [_normalise(a) for a in crossref_authors]
        # Extract surname (first word of each author name)
        surnames = []
        for a in authors_norm:
            parts = a.split()
            if parts:
                surnames.append(parts[-1] if len(parts) > 1 else parts[0])

        matched_authors = sum(1 for s in surnames if s in ref_norm)
        if surnames:
            author_ratio = matched_authors / len(surnames)
            if 0 < matched_authors and author_ratio < 0.8:
                details.append(f"only {matched_authors}/{len(surnames)} authors match")
            elif matched_authors == 0 and title_match_ratio > 0:
                details.append("no author names match")

    if details:
        return True, "; ".join(details)

    return False, None


def lookup_crossref(reference: str) -> ReferenceResult:
    """Look up a reference in Crossref. Uses DOI if present, else bibliographic search."""
    doi_match = DOI_PATTERN.search(reference)

    try:
        if doi_match:
            doi = doi_match.group(0).rstrip(".")
            url = f"https://api.crossref.org/works/{doi}"
            response = _rate_limited_get(url)
            if response.ok:
                message = response.json().get("message", {})
                titles = message.get("title", [])
                cr_title = titles[0] if titles else None
                # Extract author surnames for partial match check
                cr_authors = []
                for author in message.get("author", []):
                    name = author.get("family") or author.get("name", "")
                    if name:
                        cr_authors.append(name)
                partial, partial_detail = _check_partial_match(reference, cr_title, cr_authors)
                return ReferenceResult(
                    reference=reference,
                    matched=True,
                    title=cr_title,
                    doi=doi,
                    partial_match=partial,
                    partial_match_details=partial_detail,
                )
            return ReferenceResult(reference=reference, matched=False, doi=doi)

        # Bibliographic query search
        response = _rate_limited_get(
            "https://api.crossref.org/works",
            params={"query.bibliographic": reference[:300], "rows": 1},
        )
        if not response.ok:
            return ReferenceResult(reference=reference, matched=False)

        items = response.json().get("message", {}).get("items", [])
        if not items:
            return ReferenceResult(reference=reference, matched=False)

        top = items[0]
        score = float(top.get("score", 0))
        titles = top.get("title", [])
        cr_title = titles[0] if titles else None
        cr_authors = []
        for author in top.get("author", []):
            name = author.get("family") or author.get("name", "")
            if name:
                cr_authors.append(name)

        is_matched = score >= 40
        partial, partial_detail = _check_partial_match(reference, cr_title, cr_authors)

        # If score is between 20–40, treat as a partial match even if title
        # overlap is low — Crossref found *something* related.
        if not is_matched and score >= 20 and cr_title:
            partial = True
            if partial_detail:
                partial_detail = (
                    f"low confidence Crossref match (score {score:.0f}); "
                    f"{partial_detail}"
                )
            else:
                partial_detail = f"low confidence Crossref match (score {score:.0f})"

        return ReferenceResult(
            reference=reference,
            matched=is_matched,
            title=cr_title,
            doi=top.get("DOI"),
            partial_match=partial,
            partial_match_details=partial_detail,
        )
    except requests.RequestException:
        return ReferenceResult(reference=reference, matched=False)


def generate_summary_pdf(results: list[ReferenceResult]) -> bytes:
    """Generate a standalone PDF summary report."""
    summary_buffer = io.BytesIO()
    c = canvas.Canvas(summary_buffer, pagesize=letter)
    _, height = letter
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "GhostCite - Reference Verification Summary")
    c.setFont("Helvetica", 9)

    y = height - 80
    for result in results:
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 9)
            y = height - 50

        status = "VERIFIED" if result.matched else "NOT FOUND"
        if result.partial_match:
            status = "PARTIAL MATCH" if result.matched else "PARTIAL MATCH (NOT VERIFIED)"
        if result.matched and not result.partial_match:
            color = Color(0, 0.5, 0)
        elif result.partial_match:
            color = Color(0.8, 0.5, 0)
        else:
            color = Color(0.7, 0, 0)
        c.setFillColor(color)
        c.drawString(50, y, f"[{status}]")
        c.setFillColor(Color(0, 0, 0))

        ref_display = result.reference[:100] + ("..." if len(result.reference) > 100 else "")
        c.drawString(120, y, ref_display)
        y -= 14

        if result.doi:
            c.setFillColor(Color(0.3, 0.3, 0.3))
            c.drawString(120, y, f"DOI: {result.doi}")
            c.setFillColor(Color(0, 0, 0))
            y -= 14

        if result.partial_match and result.partial_match_details:
            c.setFillColor(Color(0.8, 0.5, 0))
            c.drawString(120, y, f"⚠ Partial: {result.partial_match_details}")
            c.setFillColor(Color(0, 0, 0))
            y -= 14

    c.save()
    return summary_buffer.getvalue()


ALLOWED_EXTENSIONS = {".pdf", ".docx"}


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    results: list[ReferenceResult] = []
    download_id = None

    if request.method == "POST":
        upload = request.files.get("document")
        if not upload or upload.filename == "":
            error = "Please upload a PDF or Word (.docx) file."
        else:
            ext = os.path.splitext(upload.filename)[1].lower()
            if ext not in ALLOWED_EXTENSIONS:
                error = "Only PDF and Word (.docx) files are supported."
            else:
                file_bytes = upload.read()
                try:
                    if ext == ".pdf":
                        text = extract_text_from_pdf(file_bytes)
                    else:
                        text = extract_text_from_docx(file_bytes)

                    references = extract_references(text)
                    refs_to_lookup = references[:100]
                    pending_results: list[ReferenceResult | None]
                    pending_results = [None] * len(refs_to_lookup)
                    with ThreadPoolExecutor(max_workers=10) as executor:
                        future_to_idx = {
                            executor.submit(lookup_crossref, ref): i
                            for i, ref in enumerate(refs_to_lookup)
                        }
                        for future in as_completed(future_to_idx):
                            idx = future_to_idx[future]
                            try:
                                pending_results[idx] = future.result()
                            except Exception:
                                pending_results[idx] = ReferenceResult(
                                    reference=refs_to_lookup[idx], matched=False
                                )
                    results = [r for r in pending_results if r is not None]

                    # Generate summary PDF
                    report_pdf = generate_summary_pdf(results)
                    download_id = str(uuid.uuid4())
                    _store_report(download_id, report_pdf)

                except Exception as e:
                    error = f"Could not process this file: {e}"

    return render_template("index.html", error=error, results=results, download_id=download_id)


@app.route("/download/<download_id>")
def download(download_id: str):
    pdf_path = _get_report_path(download_id)
    if not pdf_path:
        return "File not found or expired.", 404
    return send_file(
        pdf_path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"ghostcite_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.pdf",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
