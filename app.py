import io
import os
import re
import tempfile
import threading
import time
import uuid
import zipfile
from collections import Counter
from collections.abc import Iterable
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime, timezone
from itertools import islice
from math import hypot, isfinite
from xml.etree import ElementTree

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


@dataclass
class DocumentMetadata:
    """Metadata extracted locally from the uploaded document."""

    filename: str
    file_type: str
    file_size_bytes: int
    properties: dict[str, str] = field(default_factory=dict)
    reference_manager_indicators: list[str] = field(default_factory=list)
    provenance_signals: list["ProvenanceSignal"] = field(default_factory=list)

    def display_items(self) -> list[tuple[str, str]]:
        """Return all metadata in the order used by the web and PDF reports."""
        items = [
            ("File name", self.filename),
            ("File type", self.file_type),
            ("File size", _format_file_size(self.file_size_bytes)),
        ]
        items.extend(self.properties.items())
        items.append(
            (
                "Reference manager indicators",
                "; ".join(self.reference_manager_indicators)
                if self.reference_manager_indicators
                else "No embedded reference-manager markers detected",
            )
        )
        return items


@dataclass(frozen=True)
class ProvenanceSignal:
    """One observable file signal, presented without an automated verdict."""

    category: str
    description: str
    evidence: str
    evidence_items: tuple[str, ...] = ()


_MAX_METADATA_FIELDS = 100
_MAX_METADATA_VALUE_LENGTH = 500


def _format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB ({size_bytes:,} bytes)"
    return f"{size_bytes / (1024 * 1024):.1f} MB ({size_bytes:,} bytes)"


def _clean_metadata_value(value) -> str | None:
    """Convert common PDF/OOXML metadata values to bounded display text."""
    if value is None:
        return None
    if isinstance(value, datetime):
        text = value.isoformat(sep=" ")
    elif isinstance(value, dict):
        # XMP language alternatives commonly use {"x-default": "..."}.
        if value.get("x-default"):
            text = str(value["x-default"])
        else:
            text = "; ".join(
                f"{key}: {item}" for key, item in value.items() if item is not None
            )
    elif isinstance(value, (list, tuple, set)):
        cleaned_items = [_clean_metadata_value(item) for item in value]
        text = "; ".join(item for item in cleaned_items if item)
    else:
        text = str(value)

    # Keep metadata on one line and remove embedded control characters.
    text = "".join(char if char.isprintable() else " " for char in text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return None
    if len(text) > _MAX_METADATA_VALUE_LENGTH:
        return text[: _MAX_METADATA_VALUE_LENGTH - 1].rstrip() + "…"
    return text


def _clean_full_evidence_value(value) -> str | None:
    """Normalise evidence text without applying the metadata display limit."""
    if value is None:
        return None
    text = str(value)
    text = "".join(char if char.isprintable() else " " for char in text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or None


def _add_metadata_property(properties: dict[str, str], label: str, value) -> str | None:
    """Add a non-empty property without allowing metadata to flood a report."""
    if len(properties) >= _MAX_METADATA_FIELDS:
        return None
    clean_label = _clean_metadata_value(label)
    clean_value = _clean_metadata_value(value)
    if not clean_label or not clean_value:
        return None
    if clean_label not in properties:
        properties[clean_label] = clean_value
    return clean_value


def _metadata_key_label(key: str) -> str:
    """Turn keys such as '/SourceModified' into readable labels."""
    key = key.strip().lstrip("/")
    key = re.sub(r"[_\-]+", " ", key)
    key = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", key)
    return " ".join(word.capitalize() for word in key.split()) or "Property"


_AI_TOOL_PATTERNS = (
    ("OpenAI/ChatGPT", re.compile(r"\bopenai\b|\bchatgpt\b", re.IGNORECASE)),
    ("Microsoft Copilot", re.compile(r"\b(?:microsoft\s+)?copilot\b", re.IGNORECASE)),
    (
        "Anthropic Claude",
        re.compile(r"\banthropic\b|\bclaude(?:\.ai|\s+ai)\b", re.IGNORECASE),
    ),
    ("Google Gemini", re.compile(r"\b(?:google\s+)?gemini\b", re.IGNORECASE)),
    ("Google NotebookLM", re.compile(r"\bnotebooklm\b", re.IGNORECASE)),
    ("Grammarly AI", re.compile(r"\bgrammarly(?:go)?\b", re.IGNORECASE)),
    ("QuillBot", re.compile(r"\bquillbot\b", re.IGNORECASE)),
    ("Wordtune", re.compile(r"\bwordtune\b", re.IGNORECASE)),
    ("Jasper AI", re.compile(r"\bjasper\s+ai\b", re.IGNORECASE)),
    ("Writesonic", re.compile(r"\bwritesonic\b", re.IGNORECASE)),
    ("Notion AI", re.compile(r"\bnotion\s+ai\b", re.IGNORECASE)),
    ("DeepL Write", re.compile(r"\bdeepl\s+write\b", re.IGNORECASE)),
    ("DeepSeek", re.compile(r"\bdeepseek\b", re.IGNORECASE)),
    ("DALL-E", re.compile(r"\bdall[\s._-]*e\b", re.IGNORECASE)),
    ("Midjourney", re.compile(r"\bmidjourney\b", re.IGNORECASE)),
    ("Stable Diffusion", re.compile(r"\bstable\s+diffusion\b", re.IGNORECASE)),
    ("Adobe Firefly", re.compile(r"\b(?:adobe\s+)?firefly\b", re.IGNORECASE)),
)

_AUTOMATION_TOOL_PATTERNS = (
    ("python-docx", re.compile(r"\bpython-docx\b", re.IGNORECASE)),
    ("ReportLab", re.compile(r"\breportlab\b", re.IGNORECASE)),
    ("pypdf", re.compile(r"\bpypdf\b", re.IGNORECASE)),
    ("Aspose", re.compile(r"\baspose(?:\.words|\.pdf)?\b", re.IGNORECASE)),
    ("docx4j", re.compile(r"\bdocx4j\b", re.IGNORECASE)),
    ("Apache POI", re.compile(r"\bapache\s+poi\b", re.IGNORECASE)),
    ("Pandoc", re.compile(r"\bpandoc\b", re.IGNORECASE)),
    ("wkhtmltopdf", re.compile(r"\bwkhtmltopdf\b", re.IGNORECASE)),
    ("WeasyPrint", re.compile(r"\bweasyprint\b", re.IGNORECASE)),
    ("LibreOffice", re.compile(r"\blibreoffice\b", re.IGNORECASE)),
    ("Google Docs", re.compile(r"\bgoogle\s+docs\b", re.IGNORECASE)),
    ("LaTeX/TeX", re.compile(r"\b(?:pdf|xe|lua)?tex\b|\blatex\b", re.IGNORECASE)),
)

_PROVENANCE_PROPERTY_LABELS = {
    "Author",
    "Last modified by",
    "Creator application",
    "PDF producer",
    "XMP creator application",
    "XMP PDF producer",
    "Company",
    "Manager",
    "Template",
}


def _append_provenance_signal(
    signals: list[ProvenanceSignal],
    category: str,
    description: str,
    evidence: str,
    evidence_items: Iterable[str] = (),
) -> None:
    cleaned_items = tuple(
        cleaned
        for item in evidence_items
        if (cleaned := _clean_full_evidence_value(item))
    )
    signal = ProvenanceSignal(
        category=category,
        description=description,
        evidence=_clean_metadata_value(evidence) or "Embedded file metadata",
        evidence_items=cleaned_items,
    )
    if signal not in signals:
        signals.append(signal)


def _tool_signals_from_text(
    text: str,
    source: str,
    include_automation: bool = True,
) -> list[ProvenanceSignal]:
    """Return explicit product markers with their precise metadata source."""
    signals: list[ProvenanceSignal] = []
    for tool_name, pattern in _AI_TOOL_PATTERNS:
        match = pattern.search(text)
        if match:
            excerpt = (
                _clean_metadata_value(
                    text[max(0, match.start() - 60) : match.end() + 60]
                )
                or tool_name
            )
            _append_provenance_signal(
                signals,
                "Explicit GenAI tool marker",
                f"Embedded file data contains a {tool_name} marker.",
                f"{source}: {excerpt}",
            )
    if include_automation:
        for tool_name, pattern in _AUTOMATION_TOOL_PATTERNS:
            match = pattern.search(text)
            if match:
                excerpt = (
                    _clean_metadata_value(
                        text[max(0, match.start() - 60) : match.end() + 60]
                    )
                    or tool_name
                )
                _append_provenance_signal(
                    signals,
                    "Document workflow",
                    f"The file identifies a {tool_name} generation or conversion tool.",
                    f"{source}: {excerpt}",
                )
    return signals


def _metadata_tool_signals(properties: dict[str, str]) -> list[ProvenanceSignal]:
    signals: list[ProvenanceSignal] = []
    for label, value in properties.items():
        is_provenance_field = (
            label in _PROVENANCE_PROPERTY_LABELS
            or label.startswith("Custom:")
            or label.startswith("PDF ")
            or label.startswith("XMP ")
        )
        if is_provenance_field:
            signals.extend(
                _tool_signals_from_text(
                    value,
                    label,
                    include_automation=label not in {"Author", "Last modified by"},
                )
            )
    return list(dict.fromkeys(signals))


def _content_credential_signals(content: bytes, source: str) -> list[ProvenanceSignal]:
    """Read declarative AI source terms from an unvalidated C2PA payload."""
    text = content[: 5 * 1024 * 1024].decode("latin-1", errors="ignore")
    lower_text = text.lower()
    signals: list[ProvenanceSignal] = []
    if "compositewithtrainedalgorithmicmedia" in lower_text:
        description = "Embedded provenance data declares content edited using GenAI."
        source_type = "compositeWithTrainedAlgorithmicMedia"
    elif "trainedalgorithmicmedia" in lower_text:
        description = "Embedded provenance data declares AI-generated media."
        source_type = "trainedAlgorithmicMedia"
    elif "trainedalgorithmicdata" in lower_text:
        description = "Embedded provenance data declares AI-generated data."
        source_type = "trainedAlgorithmicData"
    else:
        description = "An embedded C2PA/Content Credentials manifest is present."
        source_type = "C2PA manifest"
    _append_provenance_signal(
        signals,
        "Content provenance",
        description + " GhostCite has not validated its signature.",
        f"{source}: {source_type}",
    )
    signals.extend(_tool_signals_from_text(text, source))
    return list(dict.fromkeys(signals))


def _pdf_colour_value(operator: bytes, operands, colour_space: str):
    try:
        values = tuple(float(value) for value in operands)
    except (TypeError, ValueError):
        return None
    if operator in {b"g", b"G"} and len(values) >= 1:
        return ("gray", values[:1])
    if operator in {b"rg", b"RG"} and len(values) >= 3:
        return ("rgb", values[:3])
    if operator in {b"k", b"K"} and len(values) >= 4:
        return ("cmyk", values[:4])
    if operator in {b"sc", b"SC", b"scn", b"SCN"}:
        space = colour_space.lower().lstrip("/")
        if space in {"devicegray", "g"} and len(values) >= 1:
            return ("gray", values[:1])
        if space in {"devicergb", "rgb"} and len(values) >= 3:
            return ("rgb", values[:3])
        if space in {"devicecmyk", "cmyk"} and len(values) >= 4:
            return ("cmyk", values[:4])
    return None


def _pdf_colour_is_white(colour) -> bool:
    if not colour:
        return False
    kind, values = colour
    if kind == "gray":
        return values[0] >= 0.99
    if kind == "rgb":
        return all(value >= 0.99 for value in values[:3])
    if kind == "cmyk":
        return all(value <= 0.01 for value in values[:4])
    return False


def _pdf_colour_description(colour) -> str:
    if not colour:
        return "white"
    kind, values = colour
    formatted = ", ".join(f"{value:g}" for value in values)
    return f"{kind.upper()} {formatted}"


def _pdf_effective_font_size(font_size, cm, tm, user_unit: float = 1.0):
    """Return rendered point size after PDF text and page transformations."""
    try:
        stored_size = abs(float(font_size))
        page_scale = float(user_unit)
        if not isfinite(stored_size) or not isfinite(page_scale) or page_scale <= 0:
            return None

        # A PDF may keep /Tf at 1 and encode its apparent font size in Tm,
        # then scale the whole text object again through the current matrix.
        vertical_x = float(cm[0]) * float(tm[2]) + float(cm[2]) * float(tm[3])
        vertical_y = float(cm[1]) * float(tm[2]) + float(cm[3]) * float(tm[3])
        rendered_size = stored_size * hypot(vertical_x, vertical_y) * page_scale
    except (IndexError, TypeError, ValueError):
        return None
    return rendered_size if isfinite(rendered_size) else None


def _pdf_text_visibility_signals(reader) -> list[ProvenanceSignal]:
    """Detect and retain every unusually small or explicitly white PDF segment."""
    tiny_pages: set[int] = set()
    white_pages: set[int] = set()
    tiny_occurrences: list[str] = []
    white_occurrences: list[str] = []

    for page_number, page in enumerate(reader.pages, start=1):
        try:
            user_unit = float(page.get("/UserUnit", 1) or 1)
        except (TypeError, ValueError):
            user_unit = 1.0
        state = {
            "fill": ("gray", (0.0,)),
            "stroke": ("gray", (0.0,)),
            "fill_space": "/DeviceGray",
            "stroke_space": "/DeviceGray",
            "render_mode": 0,
        }
        stack: list[dict] = []

        def visit_operand(operator, operands, _cm, _tm):
            if operator == b"q":
                stack.append(state.copy())
            elif operator == b"Q" and stack:
                previous = stack.pop()
                state.clear()
                state.update(previous)
            elif operator == b"BT":
                state["render_mode"] = 0
            elif operator == b"cs" and operands:
                state["fill_space"] = str(operands[0])
            elif operator == b"CS" and operands:
                state["stroke_space"] = str(operands[0])
            elif operator == b"Tr" and operands:
                try:
                    state["render_mode"] = int(operands[0])
                except (TypeError, ValueError):
                    pass
            elif operator in {b"g", b"rg", b"k", b"sc", b"scn"}:
                colour = _pdf_colour_value(
                    operator,
                    operands,
                    str(state["fill_space"]),
                )
                if colour:
                    state["fill"] = colour
            elif operator in {b"G", b"RG", b"K", b"SC", b"SCN"}:
                colour = _pdf_colour_value(
                    operator,
                    operands,
                    str(state["stroke_space"]),
                )
                if colour:
                    state["stroke"] = colour

        def visit_text(text, cm, tm, _font, font_size):
            cleaned = _clean_full_evidence_value(text)
            if not cleaned:
                return
            try:
                stored_size = abs(float(font_size))
            except (TypeError, ValueError):
                stored_size = None
            size = _pdf_effective_font_size(font_size, cm, tm, user_unit)
            if size is not None and 0 <= size < 2:
                tiny_pages.add(page_number)
                size_description = f"effective font size {size:g} pt"
                if stored_size is not None and abs(stored_size - size) >= 0.005:
                    size_description += (
                        f" (stored {stored_size:g} pt before PDF transforms)"
                    )
                tiny_occurrences.append(
                    f"Page {page_number}; {size_description}; text: '{cleaned}'"
                )

            render_mode = int(state["render_mode"])
            uses_fill = render_mode in {0, 2, 4, 6}
            uses_stroke = render_mode in {1, 2, 5, 6}
            white_colour = None
            if uses_fill and _pdf_colour_is_white(state["fill"]):
                white_colour = state["fill"]
            elif uses_stroke and _pdf_colour_is_white(state["stroke"]):
                white_colour = state["stroke"]
            if white_colour:
                white_pages.add(page_number)
                white_occurrences.append(
                    f"Page {page_number}; colour "
                    f"{_pdf_colour_description(white_colour)}; text: '{cleaned}'"
                )

        try:
            page.extract_text(
                visitor_operand_before=visit_operand,
                visitor_text=visit_text,
            )
        except Exception:
            continue

    signals: list[ProvenanceSignal] = []
    if tiny_occurrences:
        tiny_count = len(tiny_occurrences)
        _append_provenance_signal(
            signals,
            "Text visibility",
            f"The PDF contains {tiny_count} text segment"
            f"{'s' if tiny_count != 1 else ''} with an effective rendered "
            "font size below 2 pt.",
            "Pages "
            + ", ".join(str(page) for page in sorted(tiny_pages))
            + f"; all {tiny_count} detected occurrences are listed in full below.",
            evidence_items=tiny_occurrences,
        )
    if white_occurrences:
        white_count = len(white_occurrences)
        _append_provenance_signal(
            signals,
            "Text visibility",
            f"The PDF contains {white_count} explicitly white text segment"
            f"{'s' if white_count != 1 else ''}.",
            "Pages "
            + ", ".join(str(page) for page in sorted(white_pages))
            + f"; all {white_count} detected occurrences are listed in full below.",
            evidence_items=white_occurrences,
        )
    return signals


_REFERENCE_MANAGER_PATTERNS = (
    ("Zotero", re.compile(r"\bzotero(?:_item)?\b", re.IGNORECASE)),
    ("Mendeley", re.compile(r"\bmendeley\b", re.IGNORECASE)),
    ("EndNote", re.compile(r"\bendnote\b|\ben\.cite\b|\ben\.reflist\b", re.IGNORECASE)),
    ("RefWorks", re.compile(r"\brefworks\b", re.IGNORECASE)),
    ("Citavi", re.compile(r"\bcitavi\b", re.IGNORECASE)),
    ("Paperpile", re.compile(r"\bpaperpile\b", re.IGNORECASE)),
    (
        "Papers/ReadCube",
        re.compile(r"\breadcube\b|\bpapers2_citation\b", re.IGNORECASE),
    ),
)


def _reference_manager_indicators(
    signals: Iterable[str], field_codes: Iterable[str] = ()
) -> list[str]:
    """Identify citation tools from properties and embedded field codes."""
    signal_text = "\n".join(str(signal) for signal in signals if signal)
    codes = [str(code) for code in field_codes if code]
    code_text = "\n".join(codes)
    combined = signal_text + "\n" + code_text
    indicators: list[str] = []

    csl_count = len(re.findall(r"\bCSL_CITATION\b", code_text, re.IGNORECASE))
    csl_bibliography_count = len(
        re.findall(r"\bCSL_BIBLIOGRAPHY\b", code_text, re.IGNORECASE)
    )
    word_count = len(re.findall(r"(?:^|\s)CITATION\s+", code_text, re.IGNORECASE))
    word_bibliography_count = len(
        re.findall(r"(?:^|\s)BIBLIOGRAPHY(?:\s|$)", code_text, re.IGNORECASE)
    )

    detected_names: list[str] = []
    for name, pattern in _REFERENCE_MANAGER_PATTERNS:
        if pattern.search(combined):
            detected_names.append(name)
            if name in {"Zotero", "Mendeley"} and csl_count:
                indicators.append(
                    f"{name} ({csl_count} embedded CSL citation field"
                    f"{'s' if csl_count != 1 else ''})"
                )
            else:
                indicators.append(name)

    if csl_count and not any(name in {"Zotero", "Mendeley"} for name in detected_names):
        indicators.append(
            f"CSL-compatible citation fields ({csl_count}; manager not identifiable)"
        )
    if csl_bibliography_count and not any(
        name in {"Zotero", "Mendeley"} for name in detected_names
    ):
        indicators.append(
            "CSL-compatible bibliography field"
            f"{'s' if csl_bibliography_count != 1 else ''} "
            f"({csl_bibliography_count}; manager not identifiable)"
        )
    if word_count:
        indicators.append(f"Microsoft Word citation fields ({word_count})")
    if word_bibliography_count:
        indicators.append(
            "Microsoft Word bibliography field"
            f"{'s' if word_bibliography_count != 1 else ''} "
            f"({word_bibliography_count})"
        )
    if re.search(r"\b(?:pdf|xe|lua)?tex\b|\blatex\b", signal_text, re.IGNORECASE):
        indicators.append("TeX/LaTeX authoring workflow")

    # Retain order while removing duplicate signals.
    return list(dict.fromkeys(indicators))


def _pdf_resolve(value):
    """Resolve a pypdf indirect object without trusting it to be well formed."""
    try:
        return value.get_object()
    except Exception:
        return value


def _pdf_name(value) -> str | None:
    cleaned = _clean_metadata_value(_pdf_resolve(value))
    if not cleaned:
        return None
    return cleaned.lstrip("/")


def _pdf_identifier(value) -> str | None:
    """Return a PDF byte-string identifier as bounded hexadecimal text."""
    value = _pdf_resolve(value)
    try:
        if isinstance(value, bytes):
            return value.hex()
        original_bytes = getattr(value, "original_bytes", None)
        if isinstance(original_bytes, bytes):
            return original_bytes.hex()
    except Exception:
        pass
    return _clean_metadata_value(value)


def _pdf_reference_key(value):
    """Create a stable key for comparing direct and indirect PDF objects."""
    for candidate in (value, _pdf_resolve(value)):
        try:
            if hasattr(candidate, "idnum"):
                return (
                    int(candidate.idnum),
                    int(getattr(candidate, "generation", 0)),
                )
            reference = getattr(candidate, "indirect_reference", None)
            if reference is not None and hasattr(reference, "idnum"):
                return (
                    int(reference.idnum),
                    int(getattr(reference, "generation", 0)),
                )
        except Exception:
            continue
    return ("direct", id(value))


def _pdf_page_geometry_metadata(reader, metadata: DocumentMetadata) -> None:
    media_sizes: Counter[str] = Counter()
    crop_sizes: Counter[str] = Counter()
    rotations: Counter[str] = Counter()
    crop_differs = False

    def rectangle_size(rectangle) -> str | None:
        try:
            width = abs(float(rectangle.right) - float(rectangle.left))
            height = abs(float(rectangle.top) - float(rectangle.bottom))
        except Exception:
            return None

        def number(value: float) -> str:
            return f"{value:.2f}".rstrip("0").rstrip(".")

        return f"{number(width)} × {number(height)} pt"

    try:
        pages = reader.pages
    except Exception:
        pages = ()
    for page in pages:
        try:
            media_size = rectangle_size(page.mediabox)
            if media_size:
                media_sizes[media_size] += 1
        except Exception:
            media_size = None
        try:
            crop_size = rectangle_size(page.cropbox)
            if crop_size:
                crop_sizes[crop_size] += 1
                crop_differs = crop_differs or crop_size != media_size
        except Exception:
            pass
        try:
            rotation = int(page.get("/Rotate", 0) or 0) % 360
            rotations[f"{rotation}°"] += 1
        except Exception:
            rotations["unreadable"] += 1

    def counter_summary(values: Counter[str]) -> str:
        return "; ".join(
            f"{value} ({count} page{'s' if count != 1 else ''})"
            for value, count in values.most_common(8)
        )

    if media_sizes:
        _add_metadata_property(
            metadata.properties,
            "PDF page size (MediaBox)",
            counter_summary(media_sizes),
        )
    if crop_differs and crop_sizes:
        _add_metadata_property(
            metadata.properties,
            "PDF visible page size (CropBox)",
            counter_summary(crop_sizes),
        )
    if rotations and (len(rotations) > 1 or "0°" not in rotations):
        _add_metadata_property(
            metadata.properties,
            "PDF page rotations",
            counter_summary(rotations),
        )


def _pdf_annotation_metadata(reader, metadata: DocumentMetadata) -> None:
    subtype_counts: Counter[str] = Counter()
    authors: list[str] = []
    dates: list[str] = []
    examples: list[str] = []
    total = 0

    try:
        pages = reader.pages
    except Exception:
        pages = ()
    for page_number, page in enumerate(pages, start=1):
        try:
            annotations = _pdf_resolve(page.get("/Annots", ())) or ()
        except Exception:
            annotations = ()
        for annotation_reference in annotations:
            if total >= 5000:
                break
            annotation = _pdf_resolve(annotation_reference)
            if not hasattr(annotation, "get"):
                continue
            total += 1
            subtype = _pdf_name(annotation.get("/Subtype")) or "Unknown"
            subtype_counts[subtype] += 1
            author = _clean_metadata_value(annotation.get("/T"))
            if author and author not in authors and len(authors) < 20:
                authors.append(author)
            for date_key in ("/CreationDate", "/M"):
                date_value = _clean_metadata_value(annotation.get(date_key))
                if date_value and date_value not in dates and len(dates) < 20:
                    dates.append(date_value)
            contents = _clean_metadata_value(annotation.get("/Contents"))
            subject = _clean_metadata_value(annotation.get("/Subj"))
            if (contents or subject) and len(examples) < 12:
                details = []
                if author:
                    details.append(f"author {author}")
                if subject:
                    details.append(f"subject {subject[:100]}")
                if contents:
                    details.append(f"text '{contents[:160]}'")
                examples.append(f"page {page_number} {subtype}: " + ", ".join(details))

    if not total:
        return
    subtype_summary = ", ".join(
        f"{subtype}: {count}" for subtype, count in subtype_counts.most_common(12)
    )
    _add_metadata_property(
        metadata.properties,
        "PDF annotations",
        f"{total} ({subtype_summary})",
    )
    if authors:
        _add_metadata_property(
            metadata.properties,
            "PDF annotation authors",
            "; ".join(authors),
        )
    if dates:
        _add_metadata_property(
            metadata.properties,
            "PDF annotation dates",
            "; ".join(dates),
        )
    evidence_parts = [f"Types: {subtype_summary}"]
    if authors:
        evidence_parts.append("authors: " + ", ".join(authors))
    if dates:
        evidence_parts.append("dates: " + ", ".join(dates[:8]))
    if examples:
        evidence_parts.append("examples: " + "; ".join(examples))
    _append_provenance_signal(
        metadata.provenance_signals,
        "Annotations and review",
        f"The PDF contains {total} page annotation object"
        f"{'s' if total != 1 else ''}; these can include comments, links, and form controls.",
        "; ".join(evidence_parts),
    )


def _pdf_layer_metadata(root, metadata: DocumentMetadata) -> None:
    try:
        optional_content = _pdf_resolve(root.get("/OCProperties"))
    except Exception:
        optional_content = None
    if not hasattr(optional_content, "get"):
        return
    try:
        group_references = _pdf_resolve(optional_content.get("/OCGs", ())) or ()
    except Exception:
        group_references = ()
    if not group_references:
        return

    default_config = _pdf_resolve(optional_content.get("/D", {}))
    if not hasattr(default_config, "get"):
        default_config = {}
    off_keys = {
        _pdf_reference_key(item)
        for item in (_pdf_resolve(default_config.get("/OFF", ())) or ())
    }
    on_keys = {
        _pdf_reference_key(item)
        for item in (_pdf_resolve(default_config.get("/ON", ())) or ())
    }
    base_state = _pdf_name(default_config.get("/BaseState")) or "ON"
    names: list[str] = []
    off_names: list[str] = []
    for index, reference in enumerate(group_references, start=1):
        group = _pdf_resolve(reference)
        name = None
        if hasattr(group, "get"):
            name = _clean_metadata_value(group.get("/Name"))
        name = name or f"Layer {index}"
        names.append(name)
        key = _pdf_reference_key(reference)
        is_off = key in off_keys or (base_state == "OFF" and key not in on_keys)
        if is_off:
            off_names.append(name)

    shown_names = "; ".join(names[:20])
    if len(names) > 20:
        shown_names += f"; and {len(names) - 20} more"
    _add_metadata_property(
        metadata.properties,
        "PDF optional-content layers",
        f"{len(names)} ({shown_names})",
    )
    description = (
        f"The PDF contains {len(names)} optional-content layer"
        f"{'s' if len(names) != 1 else ''}."
    )
    evidence = f"Default base state: {base_state}; layers: {shown_names}"
    if off_names:
        description += (
            f" {len(off_names)} layer{'s are' if len(off_names) != 1 else ' is'} "
            "configured off by default."
        )
        evidence += "; off by default: " + ", ".join(off_names[:20])
    _append_provenance_signal(
        metadata.provenance_signals,
        "Optional content",
        description,
        evidence,
    )


def _pdf_piece_info_metadata(reader, root, metadata: DocumentMetadata) -> None:
    entries: list[str] = []

    def inspect_piece_info(container, location: str) -> None:
        if len(entries) >= 30 or not hasattr(container, "get"):
            return
        piece_info = _pdf_resolve(container.get("/PieceInfo"))
        if not hasattr(piece_info, "items"):
            return
        try:
            items = list(piece_info.items())[:20]
        except Exception:
            return
        for application, value in items:
            details = _pdf_resolve(value)
            modified = None
            if hasattr(details, "get"):
                modified = _clean_metadata_value(details.get("/LastModified"))
            entry = f"{location}: {_pdf_name(application) or 'application data'}"
            if modified:
                entry += f" (last modified {modified})"
            if entry not in entries:
                entries.append(entry)

    inspect_piece_info(root, "document")
    try:
        pages = reader.pages
    except Exception:
        pages = ()
    for page_number, page in enumerate(pages, start=1):
        inspect_piece_info(page, f"page {page_number}")
        if len(entries) >= 30:
            break
    if not entries:
        return
    summary = "; ".join(entries)
    _add_metadata_property(
        metadata.properties,
        "PDF page-piece metadata",
        summary,
    )
    _append_provenance_signal(
        metadata.provenance_signals,
        "Editing history",
        "The PDF contains application-specific page-piece metadata.",
        summary,
    )


def _pdf_action_metadata(reader, root, metadata: DocumentMetadata) -> None:
    actions: list[str] = []

    def record_action(location: str, value, depth: int = 0) -> None:
        if len(actions) >= 100 or depth > 4 or value is None:
            return
        action = _pdf_resolve(value)
        if isinstance(action, (list, tuple)):
            if location == "document open":
                actions.append("document open destination")
            return
        if not hasattr(action, "get"):
            actions.append(location)
            return
        action_type = _pdf_name(action.get("/S")) or "unspecified action"
        label = f"{location}: {action_type}"
        if label not in actions:
            actions.append(label)
        next_actions = action.get("/Next")
        if next_actions is not None:
            resolved_next = _pdf_resolve(next_actions)
            if isinstance(resolved_next, (list, tuple)):
                for next_action in resolved_next[:20]:
                    record_action(f"{location} next", next_action, depth + 1)
            else:
                record_action(f"{location} next", resolved_next, depth + 1)

    def record_additional_actions(location: str, value) -> None:
        action_dictionary = _pdf_resolve(value)
        if not hasattr(action_dictionary, "items"):
            return
        try:
            items = list(action_dictionary.items())[:30]
        except Exception:
            return
        for trigger, action in items:
            record_action(f"{location} {_pdf_name(trigger) or 'trigger'}", action)

    try:
        if root.get("/OpenAction") is not None:
            record_action("document open", root.get("/OpenAction"))
        record_additional_actions("document", root.get("/AA"))
    except Exception:
        pass

    try:
        names = _pdf_resolve(root.get("/Names"))
        javascript_tree = (
            _pdf_resolve(names.get("/JavaScript")) if hasattr(names, "get") else None
        )
    except Exception:
        javascript_tree = None

    def walk_javascript_tree(tree, depth: int = 0) -> None:
        tree = _pdf_resolve(tree)
        if depth > 8 or len(actions) >= 100 or not hasattr(tree, "get"):
            return
        try:
            named_values = _pdf_resolve(tree.get("/Names", ())) or ()
            for index in range(0, min(len(named_values), 100), 2):
                name = _clean_metadata_value(named_values[index]) or "unnamed"
                if index + 1 < len(named_values):
                    record_action(f"named JavaScript '{name}'", named_values[index + 1])
            for kid in (_pdf_resolve(tree.get("/Kids", ())) or ())[:50]:
                walk_javascript_tree(kid, depth + 1)
        except Exception:
            return

    if javascript_tree is not None:
        walk_javascript_tree(javascript_tree)

    try:
        pages = reader.pages
    except Exception:
        pages = ()
    for page_number, page in enumerate(pages, start=1):
        try:
            record_additional_actions(f"page {page_number}", page.get("/AA"))
            annotations = _pdf_resolve(page.get("/Annots", ())) or ()
        except Exception:
            annotations = ()
        for annotation_reference in annotations[:500]:
            annotation = _pdf_resolve(annotation_reference)
            if not hasattr(annotation, "get"):
                continue
            record_action(f"page {page_number} annotation", annotation.get("/A"))
            record_additional_actions(
                f"page {page_number} annotation", annotation.get("/AA")
            )
        if len(actions) >= 100:
            break

    if not actions:
        return
    summary = "; ".join(actions[:20])
    if len(actions) > 20:
        summary += f"; and {len(actions) - 20} more"
    _add_metadata_property(
        metadata.properties,
        "PDF actions and scripts",
        f"{len(actions)} ({summary})",
    )
    _append_provenance_signal(
        metadata.provenance_signals,
        "Interactive document behavior",
        "The PDF declares document, page, annotation, or JavaScript actions. "
        "GhostCite did not execute them.",
        summary,
    )


def _pdf_structural_metadata(reader, metadata: DocumentMetadata) -> None:
    try:
        trailer = reader.trailer
    except Exception:
        trailer = {}
    try:
        identifiers = _pdf_resolve(trailer.get("/ID", ())) or ()
    except Exception:
        identifiers = ()
    if identifiers:
        try:
            _add_metadata_property(
                metadata.properties,
                "PDF permanent file ID",
                _pdf_identifier(identifiers[0]),
            )
            if len(identifiers) > 1:
                _add_metadata_property(
                    metadata.properties,
                    "PDF revision file ID",
                    _pdf_identifier(identifiers[1]),
                )
        except Exception:
            pass
    try:
        _add_metadata_property(
            metadata.properties,
            "PDF trailer size",
            trailer.get("/Size"),
        )
    except Exception:
        pass
    try:
        root = reader.root_object
    except Exception:
        root = _pdf_resolve(trailer.get("/Root", {}))
    if not hasattr(root, "get"):
        return

    catalog_properties = (
        ("PDF catalog version", "/Version", True),
        ("Document language", "/Lang", False),
        ("PDF page mode", "/PageMode", True),
        ("PDF page layout", "/PageLayout", True),
    )
    for label, key, is_name in catalog_properties:
        try:
            value = root.get(key)
        except Exception:
            value = None
        if value is not None:
            _add_metadata_property(
                metadata.properties,
                label,
                _pdf_name(value) if is_name else value,
            )

    try:
        mark_info = _pdf_resolve(root.get("/MarkInfo"))
        marked = bool(mark_info.get("/Marked")) if hasattr(mark_info, "get") else False
        tagged = marked or root.get("/StructTreeRoot") is not None
        _add_metadata_property(
            metadata.properties,
            "Tagged PDF",
            "Yes" if tagged else "No",
        )
    except Exception:
        pass

    try:
        viewer_preferences = _pdf_resolve(root.get("/ViewerPreferences"))
        preference_items = []
        if hasattr(viewer_preferences, "items"):
            for key, value in list(viewer_preferences.items())[:30]:
                cleaned = _clean_metadata_value(_pdf_resolve(value))
                if cleaned:
                    preference_items.append(
                        f"{_metadata_key_label(str(key))}: {cleaned.lstrip('/')}"
                    )
        if preference_items:
            _add_metadata_property(
                metadata.properties,
                "PDF viewer preferences",
                "; ".join(preference_items),
            )
    except Exception:
        pass

    try:
        permissions = _pdf_resolve(root.get("/Perms"))
        if hasattr(permissions, "keys"):
            permission_types = [
                _pdf_name(key) or "permission" for key in list(permissions.keys())[:20]
            ]
            if permission_types:
                _add_metadata_property(
                    metadata.properties,
                    "PDF permission dictionaries",
                    "; ".join(permission_types),
                )
    except Exception:
        pass

    structural_extractors = (
        (_pdf_page_geometry_metadata, (reader, metadata)),
        (_pdf_annotation_metadata, (reader, metadata)),
        (_pdf_layer_metadata, (root, metadata)),
        (_pdf_piece_info_metadata, (reader, root, metadata)),
        (_pdf_action_metadata, (reader, root, metadata)),
    )
    for extractor, arguments in structural_extractors:
        try:
            extractor(*arguments)
        except Exception:
            continue


_XMP_NAMESPACE_NAMES = {
    "http://purl.org/dc/elements/1.1/": "Dublin Core",
    "http://ns.adobe.com/xap/1.0/": "XMP",
    "http://ns.adobe.com/xap/1.0/mm/": "XMP media management",
    "http://ns.adobe.com/pdf/1.3/": "Adobe PDF",
    "http://ns.adobe.com/pdfx/1.3/": "PDF extension",
    "http://ns.adobe.com/photoshop/1.0/": "Photoshop",
    "http://ns.adobe.com/xap/1.0/rights/": "XMP rights",
    "http://www.aiim.org/pdfa/ns/id/": "PDF/A",
    "http://iptc.org/std/Iptc4xmpCore/1.0/xmlns/": "IPTC Core",
    "http://iptc.org/std/Iptc4xmpExt/2008-02-29/": "IPTC Extension",
}


def _xmp_expanded_name(tag) -> tuple[str, str]:
    if not isinstance(tag, str):
        return "", ""
    if tag.startswith("{") and "}" in tag:
        namespace, local_name = tag[1:].split("}", 1)
        return namespace, local_name
    return "", tag


def _xmp_property_values(element) -> list[str]:
    values: list[str] = []
    ignored_attributes = {"about", "parseType", "datatype"}
    for node in element.iter():
        _, node_name = _xmp_expanded_name(node.tag)
        text = _clean_metadata_value(node.text)
        if text:
            value = (
                text
                if node_name in {"li", "Alt", "Bag", "Seq"}
                else f"{node_name}: {text}"
            )
            if value not in values:
                values.append(value)
        for attribute, raw_value in node.attrib.items():
            _, attribute_name = _xmp_expanded_name(attribute)
            if attribute_name in ignored_attributes:
                continue
            cleaned = _clean_metadata_value(raw_value)
            if cleaned:
                value = f"{attribute_name}: {cleaned}"
                if value not in values:
                    values.append(value)
        if len(values) >= 16:
            break
    return values


def _xmp_history_events(root) -> list[str]:
    events: list[str] = []
    for history in root.iter():
        namespace, local_name = _xmp_expanded_name(history.tag)
        if local_name != "History" or namespace != "http://ns.adobe.com/xap/1.0/mm/":
            continue
        candidates = [
            item
            for item in history.iter()
            if _xmp_expanded_name(item.tag)[1] in {"li", "ResourceEvent"}
        ]
        for candidate in candidates:
            fields: dict[str, str] = {}
            for node in candidate.iter():
                _, node_name = _xmp_expanded_name(node.tag)
                text = _clean_metadata_value(node.text)
                if text and node_name not in {"li", "Description", "ResourceEvent"}:
                    fields.setdefault(node_name, text)
                for attribute, raw_value in node.attrib.items():
                    _, attribute_name = _xmp_expanded_name(attribute)
                    cleaned = _clean_metadata_value(raw_value)
                    if cleaned and attribute_name not in {"about", "parseType"}:
                        fields.setdefault(attribute_name, cleaned)
            if not fields:
                continue
            preferred_order = (
                "action",
                "softwareAgent",
                "when",
                "changed",
                "parameters",
                "instanceID",
            )
            ordered_names = [name for name in preferred_order if name in fields]
            ordered_names.extend(name for name in fields if name not in ordered_names)
            summary = ", ".join(
                f"{_metadata_key_label(name)}: {fields[name]}"
                for name in ordered_names[:8]
            )
            if summary not in events:
                events.append(summary)
            if len(events) >= 30:
                return events
    return events


def _xmp_digital_source_signals(values: list[str]) -> list[ProvenanceSignal]:
    signals: list[ProvenanceSignal] = []
    combined = " ".join(values).lower()
    description = None
    source_type = None
    if "compositewithtrainedalgorithmicmedia" in combined:
        description = "Embedded XMP declares content edited using GenAI."
        source_type = "compositeWithTrainedAlgorithmicMedia"
    elif "trainedalgorithmicmedia" in combined:
        description = "Embedded XMP declares AI-generated media."
        source_type = "trainedAlgorithmicMedia"
    elif "trainedalgorithmicdata" in combined:
        description = "Embedded XMP declares AI-generated data."
        source_type = "trainedAlgorithmicData"
    if description and source_type:
        _append_provenance_signal(
            signals,
            "Content provenance",
            description
            + " This is declarative metadata and was not independently verified.",
            f"XMP DigitalSourceType: {source_type}",
        )
    return signals


def _pdf_extended_xmp_metadata(xmp, metadata: DocumentMetadata) -> None:
    """Read bounded edit history and custom properties from the raw XMP packet."""
    try:
        raw_xmp = xmp.stream.get_data()
    except Exception:
        return
    if not raw_xmp or len(raw_xmp) > 2 * 1024 * 1024:
        return
    try:
        root = ElementTree.fromstring(raw_xmp)
    except (ElementTree.ParseError, ValueError):
        return

    history_events = _xmp_history_events(root)
    if history_events:
        history_summary = "; ".join(
            f"event {index}: {event}"
            for index, event in enumerate(history_events[:10], start=1)
        )
        _add_metadata_property(
            metadata.properties,
            "XMP editing history",
            f"{len(history_events)} event{'s' if len(history_events) != 1 else ''}; "
            + history_summary,
        )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Editing history",
            f"The embedded XMP packet contains {len(history_events)} recorded edit-history event"
            f"{'s' if len(history_events) != 1 else ''}.",
            history_summary,
        )

    standard_properties = {
        ("http://purl.org/dc/elements/1.1/", name)
        for name in (
            "title",
            "creator",
            "description",
            "subject",
            "contributor",
            "coverage",
            "date",
            "format",
            "language",
            "publisher",
            "relation",
            "rights",
            "source",
            "type",
            "identifier",
        )
    }
    standard_properties.update(
        {
            ("http://ns.adobe.com/xap/1.0/", name)
            for name in ("CreatorTool", "CreateDate", "ModifyDate", "MetadataDate")
        }
    )
    standard_properties.update(
        {
            ("http://ns.adobe.com/xap/1.0/mm/", "DocumentID"),
            ("http://ns.adobe.com/xap/1.0/mm/", "InstanceID"),
            ("http://ns.adobe.com/xap/1.0/mm/", "History"),
            ("http://ns.adobe.com/pdf/1.3/", "Keywords"),
            ("http://ns.adobe.com/pdf/1.3/", "PDFVersion"),
            ("http://ns.adobe.com/pdf/1.3/", "Producer"),
            ("http://www.aiim.org/pdfa/ns/id/", "part"),
            ("http://www.aiim.org/pdfa/ns/id/", "conformance"),
        }
    )

    digital_source_types: list[str] = []
    for node in root.iter():
        _, node_name = _xmp_expanded_name(node.tag)
        if node_name.lower() == "digitalsourcetype":
            for value in _xmp_property_values(node):
                if value not in digital_source_types:
                    digital_source_types.append(value)
        for attribute, value in node.attrib.items():
            _, attribute_name = _xmp_expanded_name(attribute)
            if attribute_name.lower() == "digitalsourcetype":
                cleaned = _clean_metadata_value(value)
                if cleaned and cleaned not in digital_source_types:
                    digital_source_types.append(cleaned)

    generic_count = 0
    for description in root.iter():
        _, description_name = _xmp_expanded_name(description.tag)
        if description_name != "Description":
            continue
        candidates = list(description.attrib.items()) + [
            (child.tag, child) for child in list(description)
        ]
        for qualified_name, raw_value in candidates:
            namespace, local_name = _xmp_expanded_name(qualified_name)
            if not local_name or (namespace, local_name) in standard_properties:
                continue
            if local_name in {"about", "parseType"}:
                continue
            if hasattr(raw_value, "tag"):
                values = _xmp_property_values(raw_value)
            else:
                cleaned = _clean_metadata_value(raw_value)
                values = [cleaned] if cleaned else []
            if not values:
                continue
            namespace_name = _XMP_NAMESPACE_NAMES.get(namespace)
            label_prefix = namespace_name or "custom"
            label = f"XMP {label_prefix} {_metadata_key_label(local_name)}"
            if _add_metadata_property(
                metadata.properties,
                label,
                "; ".join(values),
            ):
                generic_count += 1
            if generic_count >= 30 or len(metadata.properties) >= _MAX_METADATA_FIELDS:
                break
        if generic_count >= 30 or len(metadata.properties) >= _MAX_METADATA_FIELDS:
            break

    if digital_source_types:
        metadata.provenance_signals.extend(
            _xmp_digital_source_signals(digital_source_types)
        )


def _pdf_metadata(file_bytes: bytes, filename: str) -> DocumentMetadata:
    metadata = DocumentMetadata(
        filename=filename,
        file_type="PDF",
        file_size_bytes=len(file_bytes),
    )
    signals: list[str] = []

    if b"%PDF-" not in file_bytes[:1024]:
        return metadata
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
    except Exception:
        return metadata

    pdf_header = _clean_metadata_value(getattr(reader, "pdf_header", None))
    if pdf_header:
        metadata.file_type = pdf_header.lstrip("%")
    try:
        _add_metadata_property(metadata.properties, "Pages", len(reader.pages))
    except Exception:
        pass
    try:
        _add_metadata_property(
            metadata.properties,
            "Encrypted",
            "Yes" if reader.is_encrypted else "No",
        )
    except Exception:
        pass

    document_info = None
    try:
        document_info = reader.metadata
    except Exception:
        pass

    if document_info:
        standard_attributes = (
            ("Title", "title"),
            ("Author", "author"),
            ("Subject", "subject"),
            ("Keywords", "keywords"),
            ("Creator application", "creator"),
            ("PDF producer", "producer"),
            ("Created", "creation_date"),
            ("Modified", "modification_date"),
        )
        for label, attribute in standard_attributes:
            try:
                value = getattr(document_info, attribute, None)
            except Exception:
                value = None
            cleaned = _add_metadata_property(metadata.properties, label, value)
            if cleaned:
                signals.append(cleaned)

        standard_keys = {
            "/Title",
            "/Author",
            "/Subject",
            "/Keywords",
            "/Creator",
            "/Producer",
            "/CreationDate",
            "/ModDate",
        }
        try:
            raw_items = document_info.items()
        except Exception:
            raw_items = ()
        for key, value in raw_items:
            if str(key) in standard_keys:
                continue
            cleaned = _add_metadata_property(
                metadata.properties,
                f"PDF {_metadata_key_label(str(key))}",
                value,
            )
            if cleaned:
                signals.append(cleaned)

    _pdf_structural_metadata(reader, metadata)

    try:
        xmp = reader.xmp_metadata
    except Exception:
        xmp = None
    if xmp:
        xmp_attributes = (
            ("Title", "dc_title"),
            ("Author", "dc_creator"),
            ("Subject", "dc_description"),
            ("Keywords", "pdf_keywords"),
            ("XMP subjects", "dc_subject"),
            ("Contributors", "dc_contributor"),
            ("Coverage", "dc_coverage"),
            ("XMP dates", "dc_date"),
            ("Format", "dc_format"),
            ("Language", "dc_language"),
            ("Publisher", "dc_publisher"),
            ("Relation", "dc_relation"),
            ("Rights", "dc_rights"),
            ("Source", "dc_source"),
            ("Document type", "dc_type"),
            ("Document identifier", "dc_identifier"),
            ("XMP creator application", "xmp_creator_tool"),
            ("XMP created", "xmp_create_date"),
            ("XMP modified", "xmp_modify_date"),
            ("XMP metadata date", "xmp_metadata_date"),
            ("XMP document ID", "xmpmm_document_id"),
            ("XMP instance ID", "xmpmm_instance_id"),
            ("XMP PDF version", "pdf_pdfversion"),
            ("XMP PDF producer", "pdf_producer"),
            ("PDF/A part", "pdfaid_part"),
            ("PDF/A conformance", "pdfaid_conformance"),
        )
        for label, attribute in xmp_attributes:
            try:
                value = getattr(xmp, attribute, None)
            except Exception:
                value = None
            cleaned = _add_metadata_property(metadata.properties, label, value)
            if cleaned:
                signals.append(cleaned)
        try:
            custom_properties = xmp.custom_properties or {}
        except Exception:
            custom_properties = {}
        for key, value in islice(custom_properties.items(), 50):
            cleaned = _add_metadata_property(
                metadata.properties,
                f"XMP {_metadata_key_label(str(key))}",
                value,
            )
            if cleaned:
                signals.append(cleaned)
        _pdf_extended_xmp_metadata(xmp, metadata)

    metadata.provenance_signals.extend(_pdf_text_visibility_signals(reader))

    attachment_names: list[str] = []
    attachment_details: list[str] = []
    c2pa_attachment_found = False
    try:
        attachments = reader.attachment_list
    except Exception:
        attachments = ()
    for attachment in attachments:
        try:
            attachment_name = _clean_metadata_value(attachment.name) or "attachment"
            attachment_names.append(attachment_name)
            try:
                subtype = _clean_metadata_value(attachment.subtype) or ""
            except Exception:
                subtype = ""
            try:
                relationship = (
                    _clean_metadata_value(attachment.associated_file_relationship) or ""
                )
            except Exception:
                relationship = ""
            detail_parts = [attachment_name]
            attachment_attributes = (
                ("alternative name", "alternative_name"),
                ("description", "description"),
                ("MIME type", "subtype"),
                ("relationship", "associated_file_relationship"),
                ("size", "size"),
                ("created", "creation_date"),
                ("modified", "modification_date"),
                ("checksum", "checksum"),
            )
            for label, attribute in attachment_attributes:
                try:
                    value = getattr(attachment, attribute)
                    if attribute == "checksum" and isinstance(value, bytes):
                        value = value.hex()
                    if attribute == "size" and value is not None:
                        value = _format_file_size(int(value))
                    cleaned = _clean_metadata_value(value)
                except Exception:
                    cleaned = None
                if cleaned and not (
                    attribute == "alternative_name" and cleaned == attachment_name
                ):
                    detail_parts.append(f"{label}: {cleaned.lstrip('/')}")
            attachment_details.append("; ".join(detail_parts))
            is_c2pa = (
                attachment_name.lower().endswith(".c2pa")
                or "c2pa" in subtype.lower()
                or "c2pa_manifest" in relationship.lower()
            )
            if is_c2pa:
                c2pa_attachment_found = True
                metadata.provenance_signals.extend(
                    _content_credential_signals(
                        attachment.content,
                        f"PDF attachment {attachment_name}",
                    )
                )
        except Exception:
            continue

    if attachment_names:
        shown_names = ", ".join(attachment_names[:10])
        if len(attachment_names) > 10:
            shown_names += f", and {len(attachment_names) - 10} more"
        _add_metadata_property(
            metadata.properties,
            "PDF embedded files",
            "; ".join(attachment_details[:10]),
        )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Embedded content",
            f"The PDF contains {len(attachment_names)} embedded file attachment"
            f"{'s' if len(attachment_names) != 1 else ''}.",
            "Attachments: "
            + (
                "; ".join(attachment_details[:10])
                if attachment_details
                else shown_names
            ),
        )

    if not c2pa_attachment_found and (
        b"application/c2pa" in file_bytes
        or b"C2PA_Manifest" in file_bytes
        or b"urn:c2pa" in file_bytes
    ):
        _append_provenance_signal(
            metadata.provenance_signals,
            "Content provenance",
            "A C2PA/Content Credentials marker is present in the PDF. "
            "GhostCite has not validated its signature.",
            "PDF object data: C2PA marker",
        )

    try:
        form_fields = reader.get_fields() or {}
    except Exception:
        form_fields = {}
    field_summaries: list[str] = []
    signature_fields: list[str] = []
    signature_details: list[str] = []
    for name, value in islice(form_fields.items(), 2000):
        if not hasattr(value, "get"):
            continue
        field_type = _pdf_name(value.get("/FT")) or "unknown type"
        alternate_name = _clean_metadata_value(value.get("/TU"))
        summary = f"{name} [{field_type}]"
        if alternate_name and alternate_name != str(name):
            summary += f" ({alternate_name})"
        if len(field_summaries) < 50:
            field_summaries.append(summary)
        if field_type != "Sig":
            continue
        signature_fields.append(str(name))
        signature = _pdf_resolve(value.get("/V"))
        detail_parts = [str(name)]
        if hasattr(signature, "get"):
            signature_attributes = (
                ("signer", "/Name"),
                ("signed", "/M"),
                ("reason", "/Reason"),
                ("location", "/Location"),
                ("contact", "/ContactInfo"),
                ("filter", "/Filter"),
                ("subfilter", "/SubFilter"),
            )
            for label, key in signature_attributes:
                cleaned = _clean_metadata_value(signature.get(key))
                if cleaned:
                    detail_parts.append(f"{label}: {cleaned.lstrip('/')}")
        signature_details.append("; ".join(detail_parts))
    if field_summaries:
        field_summary = "; ".join(field_summaries)
        if len(form_fields) > len(field_summaries):
            field_summary += f"; and {len(form_fields) - len(field_summaries)} more"
        _add_metadata_property(
            metadata.properties,
            "PDF form fields",
            f"{len(form_fields)} ({field_summary})",
        )
    if signature_fields:
        signature_summary = "; ".join(signature_details[:20])
        _add_metadata_property(
            metadata.properties,
            "PDF digital-signature metadata",
            signature_summary,
        )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Document provenance",
            f"The PDF contains {len(signature_fields)} digital signature field"
            f"{'s' if len(signature_fields) != 1 else ''}.",
            "Signature fields: " + signature_summary,
        )

    revision_count = file_bytes.count(b"%%EOF")
    if revision_count > 1:
        _add_metadata_property(
            metadata.properties,
            "PDF incremental save revisions",
            revision_count,
        )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Editing history",
            "The PDF contains multiple incremental save revisions.",
            f"PDF end-of-file revision markers: {revision_count}",
        )

    metadata.reference_manager_indicators = _reference_manager_indicators(
        list(metadata.properties.values()) + signals
    )
    metadata.provenance_signals.extend(_metadata_tool_signals(metadata.properties))

    metadata.provenance_signals = list(dict.fromkeys(metadata.provenance_signals))
    return metadata


def _xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_attribute(element, local_name: str) -> str | None:
    for key, value in element.attrib.items():
        if _xml_local_name(key) == local_name:
            return value
    return None


def _read_zip_part(
    archive: zipfile.ZipFile, name: str, max_bytes: int = 5 * 1024 * 1024
) -> bytes:
    """Read a bounded OOXML part to avoid expanding an oversized ZIP member."""
    try:
        info = archive.getinfo(name)
    except KeyError:
        return b""
    if info.file_size > max_bytes:
        return b""
    try:
        return archive.read(name)
    except (OSError, RuntimeError, zipfile.BadZipFile):
        return b""


def _xml_child(element, local_name: str):
    if element is None:
        return None
    for child in element:
        if _xml_local_name(child.tag) == local_name:
            return child
    return None


def _docx_run_properties(run_properties) -> tuple[float | None, str | None]:
    if run_properties is None:
        return (None, None)
    sizes: list[float] = []
    colour = None
    for child in run_properties:
        local_name = _xml_local_name(child.tag)
        value = _xml_attribute(child, "val")
        if local_name in {"sz", "szCs"} and value is not None:
            try:
                sizes.append(float(value) / 2)
            except ValueError:
                pass
        elif local_name == "color" and value is not None:
            colour = value
    return (min(sizes) if sizes else None, colour)


def _merge_docx_run_properties(
    base: tuple[float | None, str | None],
    override: tuple[float | None, str | None],
) -> tuple[float | None, str | None]:
    return (
        override[0] if override[0] is not None else base[0],
        override[1] if override[1] is not None else base[1],
    )


def _docx_style_resolver(archive: zipfile.ZipFile):
    styles_xml = _read_zip_part(archive, "word/styles.xml")
    if not styles_xml:
        return (lambda _style_id: (None, None)), (None, None)
    try:
        root = ElementTree.fromstring(styles_xml)
    except ElementTree.ParseError:
        return (lambda _style_id: (None, None)), (None, None)

    defaults = (None, None)
    style_data: dict[str, tuple[str | None, tuple[float | None, str | None]]] = {}
    for child in root:
        local_name = _xml_local_name(child.tag)
        if local_name == "docDefaults":
            run_default = _xml_child(child, "rPrDefault")
            defaults = _docx_run_properties(_xml_child(run_default, "rPr"))
        elif local_name == "style":
            style_id = _xml_attribute(child, "styleId")
            if not style_id:
                continue
            based_on_node = _xml_child(child, "basedOn")
            based_on = (
                _xml_attribute(based_on_node, "val")
                if based_on_node is not None
                else None
            )
            style_data[style_id] = (
                based_on,
                _docx_run_properties(_xml_child(child, "rPr")),
            )

    cache: dict[str, tuple[float | None, str | None]] = {}

    def resolve_style(
        style_id: str | None,
        resolving: set[str] | None = None,
    ) -> tuple[float | None, str | None]:
        if not style_id or style_id not in style_data:
            return (None, None)
        if style_id in cache:
            return cache[style_id]
        resolving = set() if resolving is None else resolving
        if style_id in resolving:
            return (None, None)
        resolving.add(style_id)
        based_on, own_properties = style_data[style_id]
        resolved = _merge_docx_run_properties(
            resolve_style(based_on, resolving),
            own_properties,
        )
        resolving.remove(style_id)
        cache[style_id] = resolved
        return resolved

    return resolve_style, defaults


def _iter_docx_runs(paragraph):
    """Yield runs in one paragraph without entering nested text-box paragraphs."""
    for child in paragraph:
        local_name = _xml_local_name(child.tag)
        if local_name == "r":
            yield child
        elif local_name != "p":
            yield from _iter_docx_runs(child)


def _docx_text_visibility_signals(
    archive: zipfile.ZipFile,
    package_names: list[str],
) -> list[ProvenanceSignal]:
    """Detect and retain every DOCX run below 2 pt or formatted white."""
    resolve_style, defaults = _docx_style_resolver(archive)
    tiny_occurrences: list[str] = []
    white_occurrences: list[str] = []
    scanned_bytes = 0

    content_part_pattern = re.compile(
        r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$"
    )
    for name in package_names:
        if not content_part_pattern.fullmatch(name):
            continue
        raw_part = _read_zip_part(archive, name, max_bytes=20 * 1024 * 1024)
        if not raw_part:
            continue
        scanned_bytes += len(raw_part)
        if scanned_bytes > 30 * 1024 * 1024:
            break
        try:
            root = ElementTree.fromstring(raw_part)
        except ElementTree.ParseError:
            continue

        paragraph_number = 0
        for paragraph in root.iter():
            if _xml_local_name(paragraph.tag) != "p":
                continue
            paragraph_number += 1
            paragraph_properties = _xml_child(paragraph, "pPr")
            paragraph_style_node = _xml_child(paragraph_properties, "pStyle")
            paragraph_style = (
                _xml_attribute(paragraph_style_node, "val")
                if paragraph_style_node is not None
                else None
            )
            for run_number, run in enumerate(_iter_docx_runs(paragraph), start=1):
                text = "".join(
                    node.text or ""
                    for node in run.iter()
                    if _xml_local_name(node.tag) in {"t", "delText"}
                )
                cleaned = _clean_full_evidence_value(text)
                if not cleaned:
                    continue

                run_properties = _xml_child(run, "rPr")
                run_style_node = _xml_child(run_properties, "rStyle")
                run_style = (
                    _xml_attribute(run_style_node, "val")
                    if run_style_node is not None
                    else None
                )
                effective = defaults
                effective = _merge_docx_run_properties(
                    effective,
                    resolve_style(paragraph_style),
                )
                effective = _merge_docx_run_properties(
                    effective,
                    resolve_style(run_style),
                )
                effective = _merge_docx_run_properties(
                    effective,
                    _docx_run_properties(run_properties),
                )
                size, colour = effective
                style_source = run_style or paragraph_style
                source_suffix = f", style {style_source}" if style_source else ""
                location = f"{name}; paragraph {paragraph_number}; run {run_number}"
                if size is not None and 0 <= size < 2:
                    tiny_occurrences.append(
                        f"{location}; font size {size:g} pt{source_suffix}; "
                        f"text: '{cleaned}'"
                    )
                if colour and colour.lstrip("#").upper() == "FFFFFF":
                    white_occurrences.append(
                        f"{location}; colour #{colour.lstrip('#').upper()}"
                        f"{source_suffix}; text: '{cleaned}'"
                    )

    signals: list[ProvenanceSignal] = []
    if tiny_occurrences:
        tiny_count = len(tiny_occurrences)
        _append_provenance_signal(
            signals,
            "Text visibility",
            f"The Word document contains {tiny_count} text run"
            f"{'s' if tiny_count != 1 else ''} with a font size below 2 pt.",
            f"All {tiny_count} detected occurrences are listed in full below.",
            evidence_items=tiny_occurrences,
        )
    if white_occurrences:
        white_count = len(white_occurrences)
        _append_provenance_signal(
            signals,
            "Text visibility",
            f"The Word document contains {white_count} explicitly white text run"
            f"{'s' if white_count != 1 else ''}.",
            f"All {white_count} detected occurrences are listed in full below.",
            evidence_items=white_occurrences,
        )
    return signals


def _docx_metadata(file_bytes: bytes, filename: str) -> DocumentMetadata:
    metadata = DocumentMetadata(
        filename=filename,
        file_type="Word document (.docx)",
        file_size_bytes=len(file_bytes),
    )
    signals: list[str] = []
    field_codes: list[str] = []
    tracked_changes = {"insertions": 0, "deletions": 0, "other": 0}
    revision_authors: set[str] = set()
    revision_dates: set[str] = set()
    comment_count = 0
    comment_authors: set[str] = set()
    comment_dates: set[str] = set()

    try:
        document = DocxDocument(io.BytesIO(file_bytes))
        core = document.core_properties
        core_attributes = (
            ("Title", "title"),
            ("Author", "author"),
            ("Subject", "subject"),
            ("Keywords", "keywords"),
            ("Category", "category"),
            ("Comments", "comments"),
            ("Language", "language"),
            ("Identifier", "identifier"),
            ("Content status", "content_status"),
            ("Version", "version"),
            ("Revision", "revision"),
            ("Last modified by", "last_modified_by"),
            ("Created", "created"),
            ("Modified", "modified"),
            ("Last printed", "last_printed"),
        )
        for label, attribute in core_attributes:
            try:
                value = getattr(core, attribute, None)
            except Exception:
                value = None
            cleaned = _add_metadata_property(metadata.properties, label, value)
            if cleaned:
                signals.append(cleaned)
    except Exception:
        pass

    try:
        archive = zipfile.ZipFile(io.BytesIO(file_bytes))
    except (OSError, zipfile.BadZipFile):
        return metadata

    with archive:
        package_names = archive.namelist()
        metadata.provenance_signals.extend(
            _docx_text_visibility_signals(archive, package_names)
        )
        c2pa_part_name = next(
            (
                name
                for name in package_names
                if name.lower() == "meta-inf/content_credential.c2pa"
            ),
            None,
        )
        if c2pa_part_name:
            c2pa_content = _read_zip_part(archive, c2pa_part_name)
            metadata.provenance_signals.extend(
                _content_credential_signals(
                    c2pa_content,
                    f"OOXML part {c2pa_part_name}",
                )
            )

        signature_parts = [
            name for name in package_names if name.startswith("_xmlsignatures/")
        ]
        if signature_parts:
            _append_provenance_signal(
                metadata.provenance_signals,
                "Document provenance",
                "The Word package contains OOXML digital signature parts.",
                f"OOXML signature parts: {len(signature_parts)}",
            )

        addin_parts = [name for name in package_names if "webextension" in name.lower()]
        if addin_parts:
            _append_provenance_signal(
                metadata.provenance_signals,
                "Document workflow",
                "The Word package contains Office web add-in metadata.",
                "OOXML add-in parts: " + ", ".join(addin_parts[:10]),
            )

        app_xml = _read_zip_part(archive, "docProps/app.xml")
        if app_xml:
            try:
                root = ElementTree.fromstring(app_xml)
            except ElementTree.ParseError:
                root = None
            if root is not None:
                app_labels = {
                    "Application": "Creator application",
                    "AppVersion": "Application version",
                    "Company": "Company",
                    "Manager": "Manager",
                    "Template": "Template",
                    "TotalTime": "Editing time (minutes, saved metadata)",
                    "Pages": "Pages (saved metadata)",
                    "Words": "Words (saved metadata)",
                    "Characters": "Characters (saved metadata)",
                    "CharactersWithSpaces": "Characters with spaces (saved metadata)",
                    "Lines": "Lines (saved metadata)",
                    "Paragraphs": "Paragraphs (saved metadata)",
                    "LinksUpToDate": "Links up to date",
                    "SharedDoc": "Shared document",
                    "HyperlinksChanged": "Hyperlinks changed",
                }
                for child in root:
                    local_name = _xml_local_name(child.tag)
                    if local_name not in app_labels:
                        continue
                    cleaned = _add_metadata_property(
                        metadata.properties,
                        app_labels[local_name],
                        child.text,
                    )
                    if cleaned:
                        signals.append(cleaned)

        custom_xml = _read_zip_part(archive, "docProps/custom.xml")
        if custom_xml:
            try:
                root = ElementTree.fromstring(custom_xml)
            except ElementTree.ParseError:
                root = None
            if root is not None:
                for prop in root:
                    name = prop.attrib.get("name") or "Property"
                    values = [node.text for node in prop.iter() if node.text]
                    cleaned = _add_metadata_property(
                        metadata.properties,
                        f"Custom: {name}",
                        "; ".join(values),
                    )
                    if cleaned:
                        signals.append(cleaned)

        # Citation fields can occur in the body, headers, footnotes or endnotes.
        # Only inspect field instructions/custom XML, not visible prose, to avoid
        # reporting a manager merely because its name appears in the paper.
        package_bytes_scanned = 0
        for name in package_names:
            is_word_xml = name.startswith("word/") and name.endswith(".xml")
            is_custom_part = name.startswith("customXml/") and name.endswith(
                (".xml", ".rels")
            )
            is_addin_part = "webextension" in name.lower() and name.endswith(
                (".xml", ".rels")
            )
            if not (is_word_xml or is_custom_part or is_addin_part):
                continue
            raw_part = _read_zip_part(archive, name)
            if not raw_part:
                continue
            package_bytes_scanned += len(raw_part)
            if package_bytes_scanned > 20 * 1024 * 1024:
                break
            if is_custom_part or is_addin_part:
                structural_text = raw_part.decode("utf-8", errors="ignore")
                # Retain only matched product names, not entire custom XML parts.
                for manager_name, pattern in _REFERENCE_MANAGER_PATTERNS:
                    if pattern.search(structural_text):
                        signals.append(manager_name)
                metadata.provenance_signals.extend(
                    _tool_signals_from_text(
                        structural_text,
                        f"OOXML part {name}",
                    )
                )
            if is_word_xml or is_addin_part:
                try:
                    root = ElementTree.fromstring(raw_part)
                except ElementTree.ParseError:
                    continue
                for node in root.iter():
                    local_name = _xml_local_name(node.tag)
                    if local_name == "instrText" and node.text:
                        field_codes.append(node.text)
                    if local_name in {"ins", "del", "moveFrom", "moveTo"}:
                        if local_name == "ins":
                            tracked_changes["insertions"] += 1
                        elif local_name == "del":
                            tracked_changes["deletions"] += 1
                        else:
                            tracked_changes["other"] += 1
                        author = _xml_attribute(node, "author")
                        date = _xml_attribute(node, "date")
                        if author:
                            revision_authors.add(author)
                        if date:
                            revision_dates.add(date)
                    elif local_name.endswith("PrChange"):
                        tracked_changes["other"] += 1
                        author = _xml_attribute(node, "author")
                        date = _xml_attribute(node, "date")
                        if author:
                            revision_authors.add(author)
                        if date:
                            revision_dates.add(date)
                    if name == "word/comments.xml" and local_name == "comment":
                        comment_count += 1
                        author = _xml_attribute(node, "author")
                        date = _xml_attribute(node, "date")
                        if author:
                            comment_authors.add(author)
                        if date:
                            comment_dates.add(date)

    metadata.reference_manager_indicators = _reference_manager_indicators(
        signals, field_codes
    )
    metadata.provenance_signals.extend(_metadata_tool_signals(metadata.properties))
    if field_codes:
        metadata.provenance_signals.extend(
            _tool_signals_from_text(
                "\n".join(field_codes),
                "OOXML field instructions",
            )
        )

    total_changes = sum(tracked_changes.values())
    if total_changes:
        if revision_authors:
            metadata.provenance_signals.extend(
                _tool_signals_from_text(
                    "\n".join(revision_authors),
                    "OOXML tracked-change authors",
                    include_automation=False,
                )
            )
        evidence_parts = [
            f"insertions: {tracked_changes['insertions']}",
            f"deletions: {tracked_changes['deletions']}",
            f"other changes: {tracked_changes['other']}",
        ]
        if revision_authors:
            evidence_parts.append(
                "authors: " + ", ".join(sorted(revision_authors)[:10])
            )
        if revision_dates:
            sorted_dates = sorted(revision_dates)
            evidence_parts.append(
                f"date range: {sorted_dates[0]} to {sorted_dates[-1]}"
            )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Editing history",
            f"The Word document contains {total_changes} tracked change record"
            f"{'s' if total_changes != 1 else ''}.",
            "; ".join(evidence_parts),
        )

    if comment_count:
        if comment_authors:
            metadata.provenance_signals.extend(
                _tool_signals_from_text(
                    "\n".join(comment_authors),
                    "OOXML comment authors",
                    include_automation=False,
                )
            )
        evidence_parts = [f"comments: {comment_count}"]
        if comment_authors:
            evidence_parts.append("authors: " + ", ".join(sorted(comment_authors)[:10]))
        if comment_dates:
            sorted_dates = sorted(comment_dates)
            evidence_parts.append(
                f"date range: {sorted_dates[0]} to {sorted_dates[-1]}"
            )
        _append_provenance_signal(
            metadata.provenance_signals,
            "Editing history",
            f"The Word document contains {comment_count} embedded comment"
            f"{'s' if comment_count != 1 else ''}.",
            "; ".join(evidence_parts),
        )

    metadata.provenance_signals = list(dict.fromkeys(metadata.provenance_signals))
    return metadata


def extract_document_metadata(
    file_bytes: bytes, filename: str, extension: str
) -> DocumentMetadata:
    """Extract standard, extended and citation-workflow metadata locally."""
    safe_filename = os.path.basename(filename.replace("\\", "/")) or "document"
    safe_filename = _clean_metadata_value(safe_filename) or "document"
    file_type = "PDF" if extension.lower() == ".pdf" else "Word document (.docx)"
    try:
        if extension.lower() == ".pdf":
            return _pdf_metadata(file_bytes, safe_filename)
        return _docx_metadata(file_bytes, safe_filename)
    except Exception:
        # Metadata is supplementary and must never prevent reference analysis.
        return DocumentMetadata(safe_filename, file_type, len(file_bytes))


DOI_PATTERN = re.compile(r"10\.\d{4,9}/[^\s,;\"'>\])}]+", re.IGNORECASE)

REFERENCE_DATE_PATTERN = r"(?:\d{4}[a-z]?|n\.?d\.?|in\s+press|forthcoming)"
BARE_REFERENCE_YEAR_PATTERN = r"(?:18|19|20|21)\d{2}[a-z]?"

# Patterns that indicate the start of a new reference entry
REF_START_PATTERNS = [
    re.compile(r"^\[\d+\]"),  # [1], [23]
    re.compile(r"^\d{1,4}\.\s"),  # 1. , 1234.
    re.compile(r"^\d{1,4}\)\s"),  # 1) , 1234)
    re.compile(r"^[•▪◦‣]\s*"),  # common Word/PDF bullet markers
]

# Harvard/APA-style reference starts.  The name token deliberately uses
# Unicode word characters and common apostrophe/dash variants so names such as
# O'Neill, O’Neill, Haliloğlu and hyphenated surnames are not merged into the
# preceding entry.
# The ASCII-lowercase guard keeps prose fragments such as "Springer, pp."
# from being mistaken for a new author while retaining Unicode name support.
_NAME_TOKEN = r"(?![a-z])[^\W\d_][\w'’\-‐‑‒–—]*"
_PERSON_AUTHOR_PREFIX = (
    _NAME_TOKEN + r"(?:\s+" + _NAME_TOKEN + r")*" + r",\s*" + _NAME_TOKEN + r"\s*\."
)
_AUTHOR_YEAR_RE = re.compile(
    r"(" + _PERSON_AUTHOR_PREFIX
    # PDF extraction can wrap a long author list before the year.
    + r"[^()]{0,600}?" + r"\(" + REFERENCE_DATE_PATTERN + r"(?:,\s*[A-Za-z]+\.?)?\)"
    r"|"
    # MLA/Chicago author-first entries often place a bare year later in the
    # entry rather than immediately after the author in parentheses.
    + _PERSON_AUTHOR_PREFIX + r"[^()]{0,600}?\b" + BARE_REFERENCE_YEAR_PATTERN + r"\b"
    r"|"
    # Institutional authors are open-ended rather than a four-name allowlist.
    # Optional acronyms cover forms such as "Agency Name (ANA) (2024)".
    + _NAME_TOKEN
    + r"(?:[ \t]+[^()\n]{1,120}?)?"
    + r"(?:[ \t]+\([A-Z][A-Z0-9&.\-]{1,20}\))?"
    + r"\.?[ \t]+\("
    + REFERENCE_DATE_PATTERN
    + r"\)"
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
        cells = [re.sub(r"[ \t]+", " ", cell).strip() for cell in raw_line.split("\t")]
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
    for line in lines[start_idx + 1 :]:
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
        combined = " ".join(parts[idx : min(idx + 3, len(parts))])
        if _looks_like_reference_start(line) or _looks_like_reference_start(combined):
            first_hit = idx
            break
    if first_hit is None:
        return None
    if any(
        BODY_SECTION_HEADING_RE.match(line) and not _looks_like_reference_start(line)
        for line in parts[: first_hit + 1]
    ):
        return None

    hits = sum(1 for line in parts if _looks_like_reference_start(line))
    doi_count = sum(len(DOI_PATTERN.findall(line)) for line in parts)
    if hits == 0 and doi_count == 0:
        return None

    # Density defeats a clean "References" entry in a contents page followed
    # by an entire chapter, while hit count favours the actual bibliography.
    density = hits / max(1, len(parts))
    quality = density * 100 + min(hits, 50) * 5 + min(doi_count, 50) - first_hit * 25
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
        if (
            last_char in ".!?)0123456789"
            or re.search(r"https?://\S+$", before) is not None
            or re.search(r"10\.\d{4,}/\S+$", before) is not None
        ):
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
                wait = 2**attempt
            time.sleep(wait)
            continue

        return response

    # Return last response even if still 429 after retries
    return response


def _normalise(text: str) -> str:
    """Lower-case and collapse whitespace for fuzzy comparison."""
    return re.sub(r"\s+", " ", text.lower().strip())


def _check_partial_match(
    reference: str,
    crossref_title: str | None,
    crossref_authors: list[str] | None = None,
) -> tuple[bool, str | None]:
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
                partial, partial_detail = _check_partial_match(
                    reference, cr_title, cr_authors
                )
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


def _pdf_safe_text(value: str) -> str:
    """Make arbitrary document metadata safe for ReportLab's base font."""
    return value.encode("cp1252", errors="replace").decode("cp1252")


def _wrap_pdf_text(text: str, max_characters: int = 105) -> list[str]:
    """Wrap report text without splitting ordinary words."""
    words: list[str] = []
    for word in _pdf_safe_text(text).split():
        if len(word) <= max_characters:
            words.append(word)
        else:
            words.extend(
                word[start : start + max_characters]
                for start in range(0, len(word), max_characters)
            )
    if not words:
        return [""]
    lines: list[str] = []
    line = words[0]
    for word in words[1:]:
        if len(line) + len(word) + 1 <= max_characters:
            line += " " + word
        else:
            lines.append(line)
            line = word
    lines.append(line)
    return lines


def generate_summary_pdf(
    results: list[ReferenceResult],
    document_metadata: DocumentMetadata | None = None,
) -> bytes:
    """Generate a standalone PDF summary report."""
    summary_buffer = io.BytesIO()
    c = canvas.Canvas(summary_buffer, pagesize=letter)
    _, height = letter

    def new_page(title: str | None = None) -> float:
        c.showPage()
        if title:
            c.setFont("Helvetica-Bold", 11)
            c.drawString(50, height - 45, title)
            c.setFont("Helvetica", 9)
            return height - 65
        c.setFont("Helvetica", 9)
        return height - 50

    def ensure_space(y_position: float, needed: float, title: str) -> float:
        if y_position - needed < 45:
            return new_page(title)
        return y_position

    def draw_pdf_lines(
        lines: Iterable[str],
        x_position: float,
        y_position: float,
        title: str,
        font_name: str = "Helvetica",
        font_size: float = 8,
        leading: float = 10,
        colour: Color | None = None,
    ) -> float:
        """Draw every wrapped line, continuing safely across report pages."""
        line_colour = colour or Color(0, 0, 0)
        for line in lines:
            y_position = ensure_space(y_position, leading + 1, title)
            c.setFont(font_name, font_size)
            c.setFillColor(line_colour)
            c.drawString(x_position, y_position, line)
            y_position -= leading
        return y_position

    c.setTitle("GhostCite Reference Verification Summary")
    c.setAuthor("GhostCite")
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 45, "GhostCite - Reference Verification Summary")
    c.setFont("Helvetica", 8)
    c.setFillColor(Color(0.35, 0.35, 0.35))
    c.drawString(
        50,
        height - 60,
        f"Generated {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}",
    )
    c.setFillColor(Color(0, 0, 0))

    y = height - 85
    if document_metadata:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y, "Document metadata")
        y -= 16
        c.setFont("Helvetica", 8)
        for label, value in document_metadata.display_items():
            lines = _wrap_pdf_text(f"{label}: {value}", max_characters=112)
            y = ensure_space(y, len(lines) * 11 + 3, "Document metadata (continued)")
            for line in lines:
                c.drawString(58, y, line)
                y -= 11
            y -= 2

        y = ensure_space(y, 45, "Provenance and automation signals")
        c.setFont("Helvetica-Bold", 10)
        c.drawString(50, y, "Provenance and automation signals")
        y -= 14
        c.setFont("Helvetica", 8)
        note_lines = _wrap_pdf_text(
            "Observable file evidence only. GhostCite does not decide whether "
            "GenAI was used; a human must interpret every signal in context.",
            max_characters=112,
        )
        for line in note_lines:
            c.drawString(58, y, line)
            y -= 11
        y -= 3

        if document_metadata.provenance_signals:
            for signal in document_metadata.provenance_signals:
                category_lines = _wrap_pdf_text(
                    signal.category.upper(), max_characters=108
                )
                description_lines = _wrap_pdf_text(
                    signal.description, max_characters=108
                )
                evidence_lines = _wrap_pdf_text(
                    f"Evidence: {signal.evidence}", max_characters=108
                )
                needed = (
                    len(category_lines) + len(description_lines) + len(evidence_lines)
                ) * 10 + 8
                y = ensure_space(
                    y,
                    needed,
                    "Provenance and automation signals (continued)",
                )
                c.setFont("Helvetica-Bold", 8)
                for line in category_lines:
                    c.drawString(58, y, line)
                    y -= 10
                c.setFont("Helvetica", 8)
                for line in description_lines:
                    c.drawString(66, y, line)
                    y -= 10
                c.setFillColor(Color(0.35, 0.35, 0.35))
                for line in evidence_lines:
                    c.drawString(66, y, line)
                    y -= 10
                c.setFillColor(Color(0, 0, 0))

                if signal.evidence_items:
                    y = ensure_space(
                        y,
                        22,
                        "Provenance and automation signals (continued)",
                    )
                    c.setFont("Helvetica-Bold", 8)
                    c.drawString(
                        66,
                        y,
                        f"Full occurrence list ({len(signal.evidence_items)}):",
                    )
                    y -= 11
                    for index, item in enumerate(signal.evidence_items, start=1):
                        item_lines = _wrap_pdf_text(
                            f"{index}. {item}", max_characters=104
                        )
                        y = ensure_space(
                            y,
                            min(len(item_lines), 2) * 10 + 2,
                            "Provenance and automation signals (continued)",
                        )
                        y = draw_pdf_lines(
                            item_lines,
                            74,
                            y,
                            "Provenance and automation signals (continued)",
                            colour=Color(0.25, 0.25, 0.25),
                        )
                        y -= 3
                    c.setFillColor(Color(0, 0, 0))
                y -= 8
        else:
            empty_lines = _wrap_pdf_text(
                "No additional provenance or automation markers were detected "
                "in the embedded file data. This does not indicate that GenAI "
                "was absent.",
                max_characters=108,
            )
            y = ensure_space(
                y,
                len(empty_lines) * 10 + 5,
                "Provenance and automation signals (continued)",
            )
            for line in empty_lines:
                c.drawString(58, y, line)
                y -= 10
        y -= 5

    y = ensure_space(y, 30, "Reference results")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, y, "Reference results")
    y -= 18
    c.setFont("Helvetica", 9)
    for result in results:
        status = "VERIFIED" if result.matched else "NOT FOUND"
        if result.partial_match:
            status = (
                "PARTIAL MATCH" if result.matched else "PARTIAL MATCH (NOT VERIFIED)"
            )
        reference_lines = _wrap_pdf_text(result.reference, max_characters=105)
        doi_lines = (
            _wrap_pdf_text(f"DOI: {result.doi}", max_characters=105)
            if result.doi
            else []
        )
        partial_detail_lines = (
            _wrap_pdf_text(
                f"Partial: {result.partial_match_details}", max_characters=105
            )
            if result.partial_match and result.partial_match_details
            else []
        )
        details_height = 27
        details_height += len(doi_lines) * 11 + (2 if doi_lines else 0)
        details_height += len(partial_detail_lines) * 11 + (
            2 if partial_detail_lines else 0
        )
        details_height += max(0, len(reference_lines) - 1) * 11
        y = ensure_space(y, details_height, "Reference results (continued)")

        if result.matched and not result.partial_match:
            color = Color(0, 0.5, 0)
        elif result.partial_match:
            color = Color(0.8, 0.5, 0)
        else:
            color = Color(0.7, 0, 0)
        c.setFillColor(color)
        c.drawString(50, y, f"[{status}]")
        c.setFillColor(Color(0, 0, 0))
        y -= 13

        for reference_line in reference_lines:
            c.drawString(70, y, reference_line)
            y -= 11
        y -= 3

        if doi_lines:
            c.setFillColor(Color(0.3, 0.3, 0.3))
            for line in doi_lines:
                c.drawString(70, y, line)
                y -= 11
            c.setFillColor(Color(0, 0, 0))
            y -= 2

        if partial_detail_lines:
            c.setFillColor(Color(0.8, 0.5, 0))
            for line in partial_detail_lines:
                c.drawString(70, y, line)
                y -= 11
            c.setFillColor(Color(0, 0, 0))
            y -= 2

    c.save()
    return summary_buffer.getvalue()


ALLOWED_EXTENSIONS = {".pdf", ".docx"}


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    results: list[ReferenceResult] = []
    download_id = None
    document_metadata = None

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
                    document_metadata = extract_document_metadata(
                        file_bytes, upload.filename, ext
                    )
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
                        report_pdf = generate_summary_pdf(results, document_metadata)
                        download_id = str(uuid.uuid4())
                        _store_report(download_id, report_pdf)

                except Exception as e:
                    error = f"Could not process this file: {e}"

    return render_template(
        "index.html",
        error=error,
        results=results,
        download_id=download_id,
        document_metadata=document_metadata,
    )


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
