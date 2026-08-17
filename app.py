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
from docx.oxml.ns import qn
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

REFERENCE_DATE_PATTERN = r"(?:\d{4}[a-z]?|n\.?d\.?|in\s+press|forthcoming)"
BARE_REFERENCE_YEAR_PATTERN = r"(?:18|19|20|21)\d{2}[a-z]?"

# Patterns that indicate the start of a new reference entry
REF_START_PATTERNS = [
    re.compile(r"^\[\d+\]"),           # [1], [23]
    re.compile(r"^\d{1,4}\.\s"),       # 1. , 1234.
    re.compile(r"^\d{1,4}\)\s"),       # 1) , 1234)
    re.compile(r"^[•▪◦‣]\s*"),          # common Word/PDF bullet markers
]

# Harvard/APA-style reference starts.  The name token deliberately uses
# Unicode word characters and common apostrophe/dash variants so names such as
# O'Neill, O’Neill, Haliloğlu and hyphenated surnames are not merged into the
# preceding entry.
# The ASCII-lowercase guard keeps prose fragments such as "Springer, pp."
# from being mistaken for a new author while retaining Unicode name support.
_NAME_TOKEN = r"(?![a-z])[^\W\d_][\w'’\-‐‑‒–—]*"
_PERSON_AUTHOR_PREFIX = (
    _NAME_TOKEN
    + r"(?:\s+" + _NAME_TOKEN + r")*"
    + r",\s*" + _NAME_TOKEN + r"\s*\."
)
_AUTHOR_YEAR_RE = re.compile(
    r"("
    + _PERSON_AUTHOR_PREFIX
    # PDF extraction can wrap a long author list before the year.
    + r"[^()]{0,600}?"
    + r"\(" + REFERENCE_DATE_PATTERN + r"(?:,\s*[A-Za-z]+\.?)?\)"
    r"|"
    # MLA/Chicago author-first entries often place a bare year later in the
    # entry rather than immediately after the author in parentheses.
    + _PERSON_AUTHOR_PREFIX
    + r"[^()]{0,600}?\b" + BARE_REFERENCE_YEAR_PATTERN + r"\b"
    r"|"
    # Institutional authors are open-ended rather than a four-name allowlist.
    # Optional acronyms cover forms such as "Agency Name (ANA) (2024)".
    + _NAME_TOKEN
    + r"(?:[ \t]+[^()\n]{1,120}?)?"
    + r"(?:[ \t]+\([A-Z][A-Z0-9&.\-]{1,20}\))?"
    + r"\.?[ \t]+\(" + REFERENCE_DATE_PATTERN + r"\)"
    r")"
)

REFERENCE_HEADING_RE = re.compile(
    r"^(?:(?:section|chapter)\s+)?"
    r"(?:\d+(?:\.\d+)*[.)]?\s+)?"
    r"(?P<heading>references(?:\s+and\s+notes)?|notes\s+and\s+references"
    r"|bibliography|works\s+cited|literature\s+cited|cited\s+references"
    r"|reference\s+list|sources\s+cited)"
    r"\s*(?:[:\-\u2013\u2014]\s*)?(?P<remainder>.*)$",
    re.IGNORECASE,
)

# These headings commonly follow a bibliography.  Stopping at them prevents
# appendices, declarations and supplementary material from being submitted to
# Crossref as if they were part of the final reference entry.
REFERENCE_END_HEADING_RE = re.compile(
    r"^(?:(?:section|chapter)\s+)?"
    r"(?:\d+(?:\.\d+)*[.)]?\s+)?"
    r"(?:appendix|appendices|supplement(?:ary|al)(?:\s+(?:material|materials|information))?"
    r"|acknowledg(?:e)?ments?|author\s+contributions?|funding|declarations?"
    r"|conflicts?\s+of\s+interest|competing\s+interests?)\b",
    re.IGNORECASE,
)

# A contents-page "References" entry is often followed by a body heading.  It
# should not win merely because author-year citations appear later in the body.
BODY_SECTION_HEADING_RE = re.compile(
    r"^(?:\d+(?:\.\d+)*[.)]?\s+)?"
    r"(?:abstract|introduction|background|literature\s+review|methods?|methodology"
    r"|materials?|results?|findings?|analysis|discussion|conclusions?|contents)\b",
    re.IGNORECASE,
)

_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_T = qn("w:t")
_W_TAB = qn("w:tab")
_W_BREAKS = {qn("w:br"), qn("w:cr")}


def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def _ooxml_text(element) -> str:
    """Return visible Word text, including hyperlinks and text boxes."""
    parts: list[str] = []
    for node in element.iter():
        if node is not element and node.tag == _W_P and parts:
            if parts[-1] != "\n":
                parts.append("\n")
        elif node.tag == _W_T and node.text:
            parts.append(node.text)
        elif node.tag == _W_TAB:
            parts.append("\t")
        elif node.tag in _W_BREAKS:
            parts.append("\n")
    return "".join(parts)


def _normalise_extracted_block(text: str) -> list[str]:
    """Normalise a Word block while preserving explicit/manual line breaks."""
    lines: list[str] = []
    for raw_line in text.splitlines() or [text]:
        cells = [
            re.sub(r"[ \t]+", " ", cell).strip()
            for cell in raw_line.split("\t")
        ]
        cleaned = "\t".join(cell for cell in cells if cell)
        if cleaned:
            lines.append(cleaned)
    return lines


def _iter_docx_blocks(element):
    """Yield paragraph and table-row text in document order.

    Word can wrap body content in structured document tags, tracked insertions,
    custom XML or other containers, so unrecognised containers are traversed
    recursively instead of being discarded.
    """
    for child in element.iterchildren():
        if child.tag == _W_P:
            yield from _normalise_extracted_block(_ooxml_text(child))
        elif child.tag == _W_TBL:
            for row in child.iterchildren(_W_TR):
                cells: list[str] = []
                for cell in row.iterchildren(_W_TC):
                    cell_text = re.sub(r"\s+", " ", _ooxml_text(cell)).strip()
                    if cell_text:
                        cells.append(cell_text)
                if cells:
                    yield "\t".join(cells)
        else:
            yield from _iter_docx_blocks(child)


def extract_text_from_docx(file_bytes: bytes) -> str:
    """Extract Word text in reading order, including tables and text boxes."""
    doc = DocxDocument(io.BytesIO(file_bytes))
    return "\n".join(_iter_docx_blocks(doc.element.body))


def _strip_page_number(line: str) -> str:
    """Remove a leading page number like '51  ' from a line."""
    return re.sub(r"^\d{1,4}\s{2,}", "", line)


def _clean_extracted_line(line: str) -> str:
    """Remove extraction-only page furniture while preserving reference text."""
    cleaned = _strip_page_number(line.strip()).strip()
    if re.fullmatch(r"\d{1,4}", cleaned):
        return ""
    return cleaned


def _looks_like_reference_start(text: str) -> bool:
    """Return whether text plausibly starts an actual reference entry."""
    return (
        any(pattern.match(text) for pattern in REF_START_PATTERNS)
        or _AUTHOR_YEAR_RE.match(text) is not None
        or DOI_PATTERN.search(text) is not None
    )


def _reference_heading_match(line: str):
    """Match a real heading, rejecting TOC entries such as 'References 53'."""
    match = REFERENCE_HEADING_RE.fullmatch(_clean_extracted_line(line))
    if match is None:
        return None
    remainder = match.group("remainder").strip()
    if remainder and not _looks_like_reference_start(remainder):
        return None
    return match


def _collect_reference_candidate(
    lines: list[str], start_idx: int, remainder: str
) -> list[str]:
    """Collect one candidate section until the next structural boundary."""
    parts = [remainder] if remainder else []
    for line in lines[start_idx + 1:]:
        cleaned = _clean_extracted_line(line)
        if not cleaned:
            continue
        if REFERENCE_END_HEADING_RE.match(cleaned):
            break
        if _reference_heading_match(cleaned) is not None:
            break
        parts.append(cleaned)
    return parts


def _candidate_reference_score(parts: list[str], start_idx: int) -> tuple | None:
    """Score a candidate using nearby citation evidence, not document position."""
    if not parts:
        return None

    # Permit a short bibliography note, but reject a contents heading followed
    # by Introduction/Methods/etc. before the first citation-shaped entry.
    first_hit = None
    for idx, line in enumerate(parts[:3]):
        combined = " ".join(parts[idx:min(idx + 3, len(parts))])
        if _looks_like_reference_start(line) or _looks_like_reference_start(combined):
            first_hit = idx
            break
    if first_hit is None:
        return None
    if any(
        BODY_SECTION_HEADING_RE.match(line)
        and not _looks_like_reference_start(line)
        for line in parts[:first_hit + 1]
    ):
        return None

    hits = sum(1 for line in parts if _looks_like_reference_start(line))
    doi_count = sum(len(DOI_PATTERN.findall(line)) for line in parts)
    if hits == 0 and doi_count == 0:
        return None

    # Density defeats a clean "References" entry in a contents page followed
    # by an entire chapter, while hit count favours the actual bibliography.
    density = hits / max(1, len(parts))
    quality = (
        density * 100
        + min(hits, 50) * 5
        + min(doi_count, 50)
        - first_hit * 25
    )
    return (quality, hits, density, doi_count, start_idx)


def find_reference_section(text: str) -> str:
    """Select the most reference-like section headed by a known label.

    Handles cases where the heading and references are on the same line
    (common in PDF extraction), tables, large post-reference appendices and
    multiple heading-like entries such as a table of contents.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return ""

    candidates: list[tuple[tuple, list[str]]] = []
    for idx, line in enumerate(lines):
        match = _reference_heading_match(line)
        if match is None:
            continue
        parts = _collect_reference_candidate(
            lines, idx, match.group("remainder").strip()
        )
        score = _candidate_reference_score(parts, idx)
        if score is not None:
            candidates.append((score, parts))

    if not candidates:
        return ""

    _, best_parts = max(candidates, key=lambda candidate: candidate[0])
    return "\n".join(best_parts)


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

    # Never scan the document body: in-text author/year citations can otherwise
    # be mistaken for bibliography entries.
    ref_text = find_reference_section(text)

    if not ref_text:
        return []

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
        if numbered_refs and len(numbered_refs) >= len(references):
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
                    if not references:
                        error = (
                            "No References or Bibliography section was found "
                            "near the end of this document."
                        )
                    else:
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
