# creepage_voltage_pdf_hybrid.py
#
# Hybrid engineering-drawing extractor for:
#   - TOP creepage distance
#   - BOTTOM creepage distance
#   - Rated voltage (kV)
#   - Basic Impulse Level / BIL (kV)
#
# Workflow:
#   1. Opens File Explorer so the user can select a PDF.
#   2. Uses pdfplumber as the primary extraction method.
#   3. Detects creepage and voltage values through:
#       - exact keyword lines;
#       - nearby continuation lines;
#       - coordinate-based nearby text;
#       - layout-text fallback;
#       - optional OCR fallback.
#   4. Creates a highlighted copy of the selected PDF.
#   5. Highlights creepage, rated-voltage, BIL, and kV text.
#   6. Opens the highlighted PDF for user confirmation.
#   7. Always permits manual entry or N/A.
#   8. Appends confirmed results to CSV.
#
# Install:
#   py -m pip install pdfplumber pymupdf pillow pytesseract
#
# Run:
#   py creepage_voltage_pdf_hybrid.py

import csv
import math
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz
except ImportError:
    fitz = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import pytesseract
except ImportError:
    pytesseract = None


# Canonical filename consumed by automation_application.py. The centralized
# application also recognizes the former drawing_extraction_results.csv name
# so previously collected design folders remain usable.
OUTPUT_CSV = "creepage_distance_results.csv"

# Leave blank when Tesseract is available through PATH.
# Example:
# TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_EXE = r""

OCR_DPI = 300
MAX_CANDIDATES = 30

GEOMETRIC_X_RADIUS = 700
GEOMETRIC_Y_RADIUS = 180

CREEPAGE_KEYWORD_REGEX = re.compile(
    r"""
    \bcreep(?:age)?\s*(?:distance|dist\.?|length)?\b
    |
    \bleakage\s*distance\b
    |
    \bsurface\s*distance\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

RATED_VOLTAGE_KEYWORD_REGEX = re.compile(
    r"""
    \brated\s+voltage\b
    |
    \bvoltage\s+rating\b
    |
    \bnominal\s+voltage\b
    |
    \bsystem\s+voltage\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

BIL_KEYWORD_REGEX = re.compile(
    r"""
    \bbasic\s+impulse\s+level\b
    |
    \bbasic\s+impulse\s+insulation\s+level\b
    |
    \bimpulse\s+level\b
    |
    \bBIL\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

COMBINED_VOLTAGE_KEYWORD_REGEX = re.compile(
    r"""
    \brated\s+voltage\s*(?:&|and|/|-)\s*basic\s+impulse\s+(?:insulation\s+)?level\b
    |
    \brated\s+voltage\s*(?:&|and|/|-)\s*BIL\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

KV_MEASUREMENT_REGEX = re.compile(
    r"""
    (?P<value>\d+(?:\.\d+)?)
    \s*
    (?P<unit>kV|KV|kv)
    """,
    re.VERBOSE,
)

EXPLICIT_MEASUREMENT_REGEX = re.compile(
    r"""
    (?P<value>\d+(?:\.\d+)?)
    \s*
    (?P<unit>
        mm
        |
        millimeters?
        |
        millimetres?
        |
        in
        |
        inch
        |
        inches
        |
        "
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

DUAL_UNIT_REGEX = re.compile(
    r"""
    (?P<inch>\d{1,4}(?:\.\d+)?)
    \s*
    [\[\(\{]
    \s*
    (?P<mm>\d{2,6}(?:\.\d+)?)
    \s*
    [\]\)\}]
    (?:\s*(?P<label>TOP|BOTTOM|UPPER|LOWER))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

LABEL_REGEX = re.compile(
    r"\b(TOP|BOTTOM|UPPER|LOWER)\b",
    re.IGNORECASE,
)


@dataclass
class CreepageCandidate:
    page_number: int
    keyword: str
    reported_text: str
    value_mm: float
    value_in: float
    location_label: str
    context: str
    source: str
    score: float


@dataclass
class VoltageCandidate:
    page_number: int
    keyword: str
    reported_text: str
    value_kv: float
    voltage_type: str
    context: str
    source: str
    score: float


def status(message: str) -> None:
    print(message, flush=True)


def select_pdf() -> Optional[Path]:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()

    selected = filedialog.askopenfilename(
        parent=root,
        title="Select bushing technical drawing PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )

    root.destroy()
    return Path(selected) if selected else None


def open_pdf(pdf_path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(pdf_path))
        elif sys.platform == "darwin":
            os.system(f'open "{pdf_path}"')
        else:
            os.system(f'xdg-open "{pdf_path}" >/dev/null 2>&1 &')
    except Exception as exc:
        status(f"Warning: could not open PDF automatically: {exc}")


def ask_yes_no(title: str, prompt: str) -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()

    answer = messagebox.askyesno(
        title,
        prompt,
        parent=root,
    )

    root.destroy()
    return answer


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_location_label(label: str) -> str:
    value = (label or "").strip().upper()

    if value == "UPPER":
        return "TOP"
    if value == "LOWER":
        return "BOTTOM"

    return value


def cluster_words_into_lines(
    words: List[dict],
    tolerance: float = 5.5,
) -> List[List[dict]]:
    lines: List[List[dict]] = []

    for word in sorted(
        words,
        key=lambda item: (
            float(item.get("top", 0.0)),
            float(item.get("x0", 0.0)),
        ),
    ):
        top = float(word.get("top", 0.0))
        placed = False

        for line in lines:
            average_top = sum(
                float(item.get("top", 0.0))
                for item in line
            ) / len(line)

            if abs(top - average_top) <= tolerance:
                line.append(word)
                placed = True
                break

        if not placed:
            lines.append([word])

    for line in lines:
        line.sort(key=lambda item: float(item.get("x0", 0.0)))

    return lines


def line_text(line: List[dict]) -> str:
    return " ".join(
        str(word.get("text", "")).strip()
        for word in line
        if str(word.get("text", "")).strip()
    )


def word_center(word: dict) -> Tuple[float, float]:
    return (
        (float(word.get("x0", 0.0)) + float(word.get("x1", 0.0))) / 2.0,
        (float(word.get("top", 0.0)) + float(word.get("bottom", 0.0))) / 2.0,
    )


def create_highlighted_copy(
    pdf_path: Path,
) -> Tuple[Path, int]:
    """
    Create a copy highlighting embedded-text or OCR-detected occurrences of:
      creep / creepage
      rated voltage
      basic impulse level / BIL
      kV
    """
    if fitz is None:
        status(
            "Highlighting skipped because PyMuPDF is not installed. "
            "Run: py -m pip install pymupdf"
        )
        return pdf_path, 0

    output_path = Path.cwd() / f"{pdf_path.stem}_drawing_fields_highlighted.pdf"
    counter = 1

    while output_path.exists():
        output_path = Path.cwd() / (
            f"{pdf_path.stem}_drawing_fields_highlighted_{counter}.pdf"
        )
        counter += 1

    highlight_count = 0
    search_terms = (
        "creep",
        "creepage",
        "rated voltage",
        "voltage rating",
        "basic impulse level",
        "basic impulse insulation level",
        "BIL",
        "kV",
    )

    try:
        with fitz.open(pdf_path) as document:
            if document.needs_pass:
                status(
                    "Highlighting skipped because the selected PDF "
                    "requires a password."
                )
                return pdf_path, 0

            for page_number, page in enumerate(document, start=1):
                page_highlights = 0
                seen_rectangles = set()

                def add_rect(rect, method):
                    nonlocal page_highlights, highlight_count

                    rect = fitz.Rect(rect)

                    if rect.is_empty or rect.width <= 0 or rect.height <= 0:
                        return

                    key = (
                        round(rect.x0, 2),
                        round(rect.y0, 2),
                        round(rect.x1, 2),
                        round(rect.y1, 2),
                    )

                    if key in seen_rectangles:
                        return

                    seen_rectangles.add(key)

                    annot = page.add_highlight_annot(rect)
                    annot.set_info(
                        content=(
                            "Automatically highlighted drawing field "
                            f"keyword ({method})."
                        )
                    )
                    annot.update()

                    page_highlights += 1
                    highlight_count += 1

                for term in search_terms:
                    try:
                        for rect in page.search_for(term, quads=False):
                            add_rect(rect, "embedded search")
                    except Exception:
                        pass

                for word_data in page.get_text("words") or []:
                    if len(word_data) < 5:
                        continue

                    token = str(word_data[4]).lower()

                    if (
                        "creep" in token
                        or token == "bil"
                        or token == "kv"
                        or "voltage" in token
                        or "impulse" in token
                    ):
                        add_rect(word_data[:4], "embedded word")

                if (
                    page_highlights == 0
                    and Image is not None
                    and pytesseract is not None
                ):
                    try:
                        matrix = fitz.Matrix(
                            OCR_DPI / 72.0,
                            OCR_DPI / 72.0,
                        )
                        pix = page.get_pixmap(
                            matrix=matrix,
                            alpha=False,
                        )
                        image = Image.frombytes(
                            "RGB",
                            [pix.width, pix.height],
                            pix.samples,
                        )
                        ocr = pytesseract.image_to_data(
                            image,
                            config="--psm 11",
                            output_type=pytesseract.Output.DICT,
                        )

                        x_scale = page.rect.width / float(pix.width)
                        y_scale = page.rect.height / float(pix.height)

                        for index, token in enumerate(ocr.get("text", [])):
                            token_lower = str(token).lower()

                            if not (
                                "creep" in token_lower
                                or token_lower == "bil"
                                or token_lower == "kv"
                                or "voltage" in token_lower
                                or "impulse" in token_lower
                            ):
                                continue

                            left = float(ocr["left"][index]) * x_scale
                            top = float(ocr["top"][index]) * y_scale
                            width = float(ocr["width"][index]) * x_scale
                            height = float(ocr["height"][index]) * y_scale

                            rect = fitz.Rect(
                                max(page.rect.x0, left - 1.5),
                                max(page.rect.y0, top - 1.0),
                                min(page.rect.x1, left + width + 1.5),
                                min(page.rect.y1, top + height + 1.0),
                            )

                            add_rect(rect, "OCR")

                    except Exception as exc:
                        status(
                            f"OCR highlighting could not run on "
                            f"page {page_number}: {exc}"
                        )

                if page_highlights:
                    status(
                        f"Highlighted {page_highlights} field keyword(s) "
                        f"on page {page_number}."
                    )

            if highlight_count == 0:
                status(
                    "No target keyword could be highlighted. "
                    "The original PDF will be opened."
                )
                return pdf_path, 0

            document.save(
                output_path,
                garbage=4,
                deflate=True,
            )

        status(
            f"Highlighted PDF created: {output_path.resolve()}"
        )
        return output_path, highlight_count

    except Exception as exc:
        status(
            f"Could not create the highlighted PDF copy: {exc}"
        )
        return pdf_path, 0


# ---------------------------------------------------------------------------
# Creepage extraction
# ---------------------------------------------------------------------------

def add_dual_unit_candidates(
    candidates: List[CreepageCandidate],
    text: str,
    page_number: int,
    keyword: str,
    source: str,
    base_score: float,
    inherited_label: str = "",
) -> None:
    for match in DUAL_UNIT_REGEX.finditer(text):
        inch_value = float(match.group("inch"))
        mm_value = float(match.group("mm"))

        converted_mm = inch_value * 25.4
        relative_error = abs(converted_mm - mm_value) / max(mm_value, 1.0)

        if relative_error > 0.08:
            continue

        label = normalize_location_label(
            match.group("label") or inherited_label
        )

        candidates.append(
            CreepageCandidate(
                page_number=page_number,
                keyword=keyword,
                reported_text=match.group(0).strip(),
                value_mm=mm_value,
                value_in=inch_value,
                location_label=label,
                context=clean_text(text),
                source=source,
                score=base_score + (75.0 if label else 0.0),
            )
        )


def add_explicit_unit_candidates(
    candidates: List[CreepageCandidate],
    text: str,
    page_number: int,
    keyword: str,
    source: str,
    base_score: float,
    inherited_label: str = "",
) -> None:
    text_label_match = LABEL_REGEX.search(text)
    text_label = (
        normalize_location_label(text_label_match.group(1))
        if text_label_match
        else normalize_location_label(inherited_label)
    )

    for match in EXPLICIT_MEASUREMENT_REGEX.finditer(text):
        value = float(match.group("value"))
        unit = match.group("unit").lower()

        if unit.startswith("mm") or unit.startswith("millim"):
            value_mm = value
            value_in = value / 25.4
        else:
            value_in = value
            value_mm = value * 25.4

        candidates.append(
            CreepageCandidate(
                page_number=page_number,
                keyword=keyword,
                reported_text=match.group(0).strip(),
                value_mm=value_mm,
                value_in=value_in,
                location_label=text_label,
                context=clean_text(text),
                source=source,
                score=base_score + (50.0 if text_label else 0.0),
            )
        )


def add_creepage_geometric_candidates(
    candidates: List[CreepageCandidate],
    words: List[dict],
    keyword_line: List[dict],
    page_number: int,
    keyword: str,
) -> None:
    keyword_words = [
        word
        for word in keyword_line
        if re.search(
            r"creep|distance|dist|leakage|surface",
            str(word.get("text", "")),
            re.IGNORECASE,
        )
    ]

    if not keyword_words:
        keyword_words = keyword_line

    keyword_x0 = min(float(word["x0"]) for word in keyword_words)
    keyword_x1 = max(float(word["x1"]) for word in keyword_words)
    keyword_top = min(float(word["top"]) for word in keyword_words)
    keyword_bottom = max(float(word["bottom"]) for word in keyword_words)

    keyword_center_x = (keyword_x0 + keyword_x1) / 2.0
    keyword_center_y = (keyword_top + keyword_bottom) / 2.0

    nearby_words = []

    for word in words:
        x0 = float(word.get("x0", 0.0))
        x1 = float(word.get("x1", 0.0))
        top = float(word.get("top", 0.0))
        bottom = float(word.get("bottom", 0.0))

        horizontal_match = (
            x0 <= keyword_x1 + GEOMETRIC_X_RADIUS
            and x1 >= keyword_x0 - 100
        )

        vertical_match = (
            top <= keyword_bottom + GEOMETRIC_Y_RADIUS
            and bottom >= keyword_top - GEOMETRIC_Y_RADIUS
        )

        if horizontal_match and vertical_match:
            nearby_words.append(word)

    nearby_lines = cluster_words_into_lines(
        nearby_words,
        tolerance=7.0,
    )

    for line in nearby_lines:
        text = line_text(line)

        if not (
            DUAL_UNIT_REGEX.search(text)
            or EXPLICIT_MEASUREMENT_REGEX.search(text)
        ):
            continue

        centers = [word_center(word) for word in line]

        if centers:
            line_center_x = sum(point[0] for point in centers) / len(centers)
            line_center_y = sum(point[1] for point in centers) / len(centers)
            distance = math.hypot(
                line_center_x - keyword_center_x,
                line_center_y - keyword_center_y,
            )
        else:
            distance = 1000.0

        score = max(800.0, 1900.0 - distance)

        label_match = LABEL_REGEX.search(text)
        inherited_label = (
            normalize_location_label(label_match.group(1))
            if label_match
            else ""
        )

        add_dual_unit_candidates(
            candidates,
            text,
            page_number,
            keyword,
            "pdfplumber geometric nearby line",
            score,
            inherited_label,
        )

        add_explicit_unit_candidates(
            candidates,
            text,
            page_number,
            keyword,
            "pdfplumber geometric nearby line",
            score - 150.0,
            inherited_label,
        )


def extract_creepage_candidates(
    pdf_path: Path,
) -> Tuple[List[CreepageCandidate], List[int]]:
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber is not installed. Run: "
            "py -m pip install pdfplumber"
        )

    candidates: List[CreepageCandidate] = []
    keyword_pages: List[int] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            status(
                f"pdfplumber analyzing creepage on page {page_number} "
                f"of {len(pdf.pages)}..."
            )

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            ) or []

            lines = cluster_words_into_lines(words)

            for line_index, line in enumerate(lines):
                current_text = line_text(line)
                keyword_match = CREEPAGE_KEYWORD_REGEX.search(current_text)

                if not keyword_match:
                    continue

                keyword_pages.append(page_number)
                keyword = keyword_match.group(0)

                status(
                    f"  creepage keyword line found: {current_text}"
                )

                same_line_tail = current_text[keyword_match.end():]

                add_dual_unit_candidates(
                    candidates,
                    same_line_tail,
                    page_number,
                    keyword,
                    "pdfplumber exact keyword line",
                    3200.0,
                )

                add_explicit_unit_candidates(
                    candidates,
                    same_line_tail,
                    page_number,
                    keyword,
                    "pdfplumber exact keyword line",
                    2900.0,
                )

                for offset in (1, 2, 3):
                    next_index = line_index + offset

                    if next_index >= len(lines):
                        break

                    continuation_text = line_text(lines[next_index])

                    if not continuation_text:
                        continue

                    label_match = LABEL_REGEX.search(continuation_text)
                    inherited_label = (
                        normalize_location_label(label_match.group(1))
                        if label_match
                        else ""
                    )

                    add_dual_unit_candidates(
                        candidates,
                        continuation_text,
                        page_number,
                        keyword,
                        f"pdfplumber continuation line +{offset}",
                        2500.0 - offset * 120.0,
                        inherited_label,
                    )

                    add_explicit_unit_candidates(
                        candidates,
                        continuation_text,
                        page_number,
                        keyword,
                        f"pdfplumber continuation line +{offset}",
                        2200.0 - offset * 120.0,
                        inherited_label,
                    )

                add_creepage_geometric_candidates(
                    candidates,
                    words,
                    line,
                    page_number,
                    keyword,
                )

            layout_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
                layout=True,
            ) or ""

            layout_lines = layout_text.splitlines()

            for line_index, raw_line in enumerate(layout_lines):
                keyword_match = CREEPAGE_KEYWORD_REGEX.search(raw_line)

                if not keyword_match:
                    continue

                keyword_pages.append(page_number)
                keyword = keyword_match.group(0)
                same_line_tail = raw_line[keyword_match.end():]

                add_dual_unit_candidates(
                    candidates,
                    same_line_tail,
                    page_number,
                    keyword,
                    "pdfplumber layout exact line",
                    3000.0,
                )

                add_explicit_unit_candidates(
                    candidates,
                    same_line_tail,
                    page_number,
                    keyword,
                    "pdfplumber layout exact line",
                    2700.0,
                )

                for offset in (1, 2, 3):
                    next_index = line_index + offset

                    if next_index >= len(layout_lines):
                        break

                    continuation_text = layout_lines[next_index].strip()

                    if not continuation_text:
                        continue

                    add_dual_unit_candidates(
                        candidates,
                        continuation_text,
                        page_number,
                        keyword,
                        f"pdfplumber layout continuation +{offset}",
                        2300.0 - offset * 120.0,
                    )

                    add_explicit_unit_candidates(
                        candidates,
                        continuation_text,
                        page_number,
                        keyword,
                        f"pdfplumber layout continuation +{offset}",
                        2000.0 - offset * 120.0,
                    )

    return deduplicate_creepage_candidates(candidates), sorted(set(keyword_pages))


def deduplicate_creepage_candidates(
    candidates: List[CreepageCandidate],
) -> List[CreepageCandidate]:
    output: List[CreepageCandidate] = []
    seen = set()

    for candidate in sorted(
        candidates,
        key=lambda item: item.score,
        reverse=True,
    ):
        key = (
            candidate.page_number,
            round(candidate.value_mm, 1),
            candidate.location_label,
        )

        if key in seen:
            continue

        seen.add(key)
        output.append(candidate)

    return output


def print_creepage_candidates(
    candidates: List[CreepageCandidate],
    target_label: str,
) -> None:
    status(
        f"\nPossible {target_label} creepage-distance measurements"
    )
    status("-" * (38 + len(target_label)))

    for index, candidate in enumerate(
        candidates[:MAX_CANDIDATES],
        start=1,
    ):
        label = candidate.location_label or "UNLABELED"

        status(
            f"\n[{index}] Page {candidate.page_number} [{label}]"
        )
        status(f"    Keyword: {candidate.keyword}")
        status(f"    Reported: {candidate.reported_text}")
        status(
            f"    Normalized: {candidate.value_mm:.3f} mm "
            f"({candidate.value_in:.4f} in)"
        )
        status(f"    Source: {candidate.source}")
        status(f"    Context: {candidate.context}")


def parse_manual_creepage(
    text: str,
) -> Optional[Tuple[float, float]]:
    explicit = EXPLICIT_MEASUREMENT_REGEX.search(text)

    if explicit:
        value = float(explicit.group("value"))
        unit = explicit.group("unit").lower()

        if unit.startswith("mm") or unit.startswith("millim"):
            return value, value / 25.4

        return value * 25.4, value

    dual = DUAL_UNIT_REGEX.search(text)

    if dual:
        inch_value = float(dual.group("inch"))
        mm_value = float(dual.group("mm"))

        converted_mm = inch_value * 25.4
        relative_error = abs(converted_mm - mm_value) / max(mm_value, 1.0)

        if relative_error <= 0.08:
            return mm_value, inch_value

    return None


def prompt_for_creepage(
    candidates: List[CreepageCandidate],
    target_label: str,
) -> Optional[Tuple[str, float, float, int, str, str]]:
    while True:
        status(f"\n{target_label} creepage-distance selection")
        status("-" * (29 + len(target_label)))
        status(
            "Enter a candidate number, M for manual entry, "
            "R to redisplay candidates, or N/A if not applicable."
        )

        choice = input(
            f"{target_label} creepage distance: "
        ).strip()

        if choice.lower() in {"n/a", "na", "none"}:
            return None

        if choice.lower() == "r":
            print_creepage_candidates(candidates, target_label)
            continue

        if choice.lower() == "m":
            while True:
                manual_text = input(
                    f"Enter {target_label} creepage distance "
                    '(examples: 1736 mm, 68.35 in, 68.35", '
                    "or type N/A): "
                ).strip()

                if manual_text.lower() in {"n/a", "na", "none"}:
                    return None

                parsed = parse_manual_creepage(manual_text)

                if parsed is None:
                    status(
                        "Could not understand that measurement. "
                        "Please include mm, in, inches, or a quote mark."
                    )
                    continue

                value_mm, value_in = parsed

                return (
                    "manual",
                    value_mm,
                    value_in,
                    0,
                    "",
                    f"User-entered {target_label} value",
                )

        if choice.isdigit():
            index = int(choice)

            if 1 <= index <= len(candidates):
                candidate = candidates[index - 1]

                status("\nSelected candidate:")
                status(f"  Reported: {candidate.reported_text}")
                status(
                    f"  Converted: {candidate.value_mm:.3f} mm "
                    f"({candidate.value_in:.4f} in)"
                )
                status(
                    f"  Detected label: "
                    f"{candidate.location_label or 'UNLABELED'}"
                )
                status(f"  Context: {candidate.context}")

                confirmation = input(
                    f"Use this as the {target_label} creepage distance? "
                    "[y/N]: "
                ).strip().lower()

                if confirmation == "y":
                    return (
                        "automatic candidate",
                        candidate.value_mm,
                        candidate.value_in,
                        candidate.page_number,
                        candidate.keyword,
                        candidate.context,
                    )

                continue

        parsed_direct = parse_manual_creepage(choice)

        if parsed_direct is not None:
            value_mm, value_in = parsed_direct

            return (
                "manual",
                value_mm,
                value_in,
                0,
                "",
                f"User-entered {target_label} value",
            )

        status(
            "Invalid selection. Enter a candidate number, a measurement, "
            "M, R, or N/A."
        )


# ---------------------------------------------------------------------------
# Voltage extraction
# ---------------------------------------------------------------------------

def add_voltage_candidates_from_text(
    candidates: List[VoltageCandidate],
    text: str,
    page_number: int,
    keyword: str,
    voltage_type: str,
    source: str,
    base_score: float,
) -> None:
    for match in KV_MEASUREMENT_REGEX.finditer(text):
        candidates.append(
            VoltageCandidate(
                page_number=page_number,
                keyword=keyword,
                reported_text=match.group(0).strip(),
                value_kv=float(match.group("value")),
                voltage_type=voltage_type,
                context=clean_text(text),
                source=source,
                score=base_score,
            )
        )


def infer_combined_voltage_candidates(
    text: str,
    page_number: int,
    keyword: str,
    source: str,
    base_score: float,
) -> List[VoltageCandidate]:
    """
    When a combined label such as 'Rated Voltage & Basic Impulse Level'
    is followed by two kV values, assign the first to rated voltage and
    the second to BIL. The user must still confirm each result.
    """
    matches = list(KV_MEASUREMENT_REGEX.finditer(text))

    if len(matches) < 2:
        return []

    first = matches[0]
    second = matches[1]

    return [
        VoltageCandidate(
            page_number=page_number,
            keyword=keyword,
            reported_text=first.group(0).strip(),
            value_kv=float(first.group("value")),
            voltage_type="rated_voltage",
            context=clean_text(text),
            source=f"{source} combined field: first kV value",
            score=base_score,
        ),
        VoltageCandidate(
            page_number=page_number,
            keyword=keyword,
            reported_text=second.group(0).strip(),
            value_kv=float(second.group("value")),
            voltage_type="bil_voltage",
            context=clean_text(text),
            source=f"{source} combined field: second kV value",
            score=base_score - 10.0,
        ),
    ]


def add_voltage_geometric_candidates(
    candidates: List[VoltageCandidate],
    words: List[dict],
    keyword_line: List[dict],
    page_number: int,
    keyword: str,
    voltage_type: str,
) -> None:
    keyword_words = [
        word
        for word in keyword_line
        if re.search(
            r"rated|voltage|basic|impulse|level|bil",
            str(word.get("text", "")),
            re.IGNORECASE,
        )
    ]

    if not keyword_words:
        keyword_words = keyword_line

    keyword_x0 = min(float(word["x0"]) for word in keyword_words)
    keyword_x1 = max(float(word["x1"]) for word in keyword_words)
    keyword_top = min(float(word["top"]) for word in keyword_words)
    keyword_bottom = max(float(word["bottom"]) for word in keyword_words)

    keyword_center_x = (keyword_x0 + keyword_x1) / 2.0
    keyword_center_y = (keyword_top + keyword_bottom) / 2.0

    nearby_words = []

    for word in words:
        x0 = float(word.get("x0", 0.0))
        x1 = float(word.get("x1", 0.0))
        top = float(word.get("top", 0.0))
        bottom = float(word.get("bottom", 0.0))

        if (
            x0 <= keyword_x1 + GEOMETRIC_X_RADIUS
            and x1 >= keyword_x0 - 100
            and top <= keyword_bottom + GEOMETRIC_Y_RADIUS
            and bottom >= keyword_top - GEOMETRIC_Y_RADIUS
        ):
            nearby_words.append(word)

    nearby_lines = cluster_words_into_lines(
        nearby_words,
        tolerance=7.0,
    )

    for line in nearby_lines:
        text = line_text(line)

        if not KV_MEASUREMENT_REGEX.search(text):
            continue

        centers = [word_center(word) for word in line]

        if centers:
            line_center_x = sum(point[0] for point in centers) / len(centers)
            line_center_y = sum(point[1] for point in centers) / len(centers)
            distance = math.hypot(
                line_center_x - keyword_center_x,
                line_center_y - keyword_center_y,
            )
        else:
            distance = 1000.0

        score = max(900.0, 2200.0 - distance)

        if voltage_type == "combined":
            candidates.extend(
                infer_combined_voltage_candidates(
                    text,
                    page_number,
                    keyword,
                    "pdfplumber geometric nearby line",
                    score,
                )
            )
        else:
            add_voltage_candidates_from_text(
                candidates,
                text,
                page_number,
                keyword,
                voltage_type,
                "pdfplumber geometric nearby line",
                score,
            )


def extract_voltage_candidates(
    pdf_path: Path,
) -> Tuple[List[VoltageCandidate], List[int]]:
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber is not installed. Run: "
            "py -m pip install pdfplumber"
        )

    candidates: List[VoltageCandidate] = []
    keyword_pages: List[int] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            status(
                f"pdfplumber analyzing voltage fields on page {page_number} "
                f"of {len(pdf.pages)}..."
            )

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            ) or []

            lines = cluster_words_into_lines(words)

            for line_index, line in enumerate(lines):
                current_text = line_text(line)

                combined_match = COMBINED_VOLTAGE_KEYWORD_REGEX.search(
                    current_text
                )
                rated_match = RATED_VOLTAGE_KEYWORD_REGEX.search(current_text)
                bil_match = BIL_KEYWORD_REGEX.search(current_text)

                if combined_match:
                    keyword_match = combined_match
                    voltage_type = "combined"
                elif bil_match:
                    keyword_match = bil_match
                    voltage_type = "bil_voltage"
                elif rated_match:
                    keyword_match = rated_match
                    voltage_type = "rated_voltage"
                else:
                    continue

                keyword_pages.append(page_number)
                keyword = keyword_match.group(0)

                status(
                    f"  voltage keyword line found: {current_text}"
                )

                same_line_tail = current_text[keyword_match.end():]

                if voltage_type == "combined":
                    candidates.extend(
                        infer_combined_voltage_candidates(
                            same_line_tail,
                            page_number,
                            keyword,
                            "pdfplumber exact keyword line",
                            3400.0,
                        )
                    )
                else:
                    add_voltage_candidates_from_text(
                        candidates,
                        same_line_tail,
                        page_number,
                        keyword,
                        voltage_type,
                        "pdfplumber exact keyword line",
                        3300.0,
                    )

                for offset in (1, 2, 3):
                    next_index = line_index + offset

                    if next_index >= len(lines):
                        break

                    continuation_text = line_text(lines[next_index])

                    if not continuation_text:
                        continue

                    if voltage_type == "combined":
                        candidates.extend(
                            infer_combined_voltage_candidates(
                                continuation_text,
                                page_number,
                                keyword,
                                f"pdfplumber continuation line +{offset}",
                                2800.0 - offset * 120.0,
                            )
                        )
                    else:
                        add_voltage_candidates_from_text(
                            candidates,
                            continuation_text,
                            page_number,
                            keyword,
                            voltage_type,
                            f"pdfplumber continuation line +{offset}",
                            2700.0 - offset * 120.0,
                        )

                add_voltage_geometric_candidates(
                    candidates,
                    words,
                    line,
                    page_number,
                    keyword,
                    voltage_type,
                )

            layout_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
                layout=True,
            ) or ""

            layout_lines = layout_text.splitlines()

            for line_index, raw_line in enumerate(layout_lines):
                combined_match = COMBINED_VOLTAGE_KEYWORD_REGEX.search(raw_line)
                rated_match = RATED_VOLTAGE_KEYWORD_REGEX.search(raw_line)
                bil_match = BIL_KEYWORD_REGEX.search(raw_line)

                if combined_match:
                    keyword_match = combined_match
                    voltage_type = "combined"
                elif bil_match:
                    keyword_match = bil_match
                    voltage_type = "bil_voltage"
                elif rated_match:
                    keyword_match = rated_match
                    voltage_type = "rated_voltage"
                else:
                    continue

                keyword_pages.append(page_number)
                keyword = keyword_match.group(0)
                same_line_tail = raw_line[keyword_match.end():]

                if voltage_type == "combined":
                    candidates.extend(
                        infer_combined_voltage_candidates(
                            same_line_tail,
                            page_number,
                            keyword,
                            "pdfplumber layout exact line",
                            3200.0,
                        )
                    )
                else:
                    add_voltage_candidates_from_text(
                        candidates,
                        same_line_tail,
                        page_number,
                        keyword,
                        voltage_type,
                        "pdfplumber layout exact line",
                        3100.0,
                    )

                for offset in (1, 2, 3):
                    next_index = line_index + offset

                    if next_index >= len(layout_lines):
                        break

                    continuation_text = layout_lines[next_index].strip()

                    if not continuation_text:
                        continue

                    if voltage_type == "combined":
                        candidates.extend(
                            infer_combined_voltage_candidates(
                                continuation_text,
                                page_number,
                                keyword,
                                f"pdfplumber layout continuation +{offset}",
                                2600.0 - offset * 120.0,
                            )
                        )
                    else:
                        add_voltage_candidates_from_text(
                            candidates,
                            continuation_text,
                            page_number,
                            keyword,
                            voltage_type,
                            f"pdfplumber layout continuation +{offset}",
                            2500.0 - offset * 120.0,
                        )

    return deduplicate_voltage_candidates(candidates), sorted(set(keyword_pages))


def deduplicate_voltage_candidates(
    candidates: List[VoltageCandidate],
) -> List[VoltageCandidate]:
    output: List[VoltageCandidate] = []
    seen = set()

    for candidate in sorted(
        candidates,
        key=lambda item: item.score,
        reverse=True,
    ):
        key = (
            candidate.page_number,
            round(candidate.value_kv, 3),
            candidate.voltage_type,
        )

        if key in seen:
            continue

        seen.add(key)
        output.append(candidate)

    return output


def print_voltage_candidates(
    candidates: List[VoltageCandidate],
    target_type: str,
    target_label: str,
) -> None:
    filtered = [
        candidate
        for candidate in candidates
        if candidate.voltage_type == target_type
    ]

    status(f"\nPossible {target_label} candidates")
    status("-" * (20 + len(target_label)))

    if not filtered:
        status("No automatic candidates found.")
        return

    for index, candidate in enumerate(
        filtered[:MAX_CANDIDATES],
        start=1,
    ):
        status(f"\n[{index}] Page {candidate.page_number}")
        status(f"    Keyword: {candidate.keyword}")
        status(f"    Reported: {candidate.reported_text}")
        status(f"    Value: {candidate.value_kv:.3f} kV")
        status(f"    Source: {candidate.source}")
        status(f"    Context: {candidate.context}")


def parse_manual_voltage(text: str) -> Optional[float]:
    match = KV_MEASUREMENT_REGEX.search(text)

    if match:
        return float(match.group("value"))

    # Since voltage fields are always in kV, allow a bare number.
    bare = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*", text)

    if bare:
        return float(bare.group(1))

    return None


def prompt_for_voltage(
    candidates: List[VoltageCandidate],
    target_type: str,
    target_label: str,
) -> Optional[Tuple[str, float, int, str, str]]:
    filtered = [
        candidate
        for candidate in candidates
        if candidate.voltage_type == target_type
    ]

    while True:
        status(f"\n{target_label} selection")
        status("-" * (10 + len(target_label)))
        status(
            "Enter a candidate number, M for manual entry, "
            "R to redisplay candidates, or N/A if not applicable."
        )

        choice = input(f"{target_label}: ").strip()

        if choice.lower() in {"n/a", "na", "none"}:
            return None

        if choice.lower() == "r":
            print_voltage_candidates(
                candidates,
                target_type,
                target_label,
            )
            continue

        if choice.lower() == "m":
            while True:
                manual_text = input(
                    f"Enter {target_label} in kV "
                    "(examples: 69 kV, 350 kV, or type N/A): "
                ).strip()

                if manual_text.lower() in {"n/a", "na", "none"}:
                    return None

                value_kv = parse_manual_voltage(manual_text)

                if value_kv is None:
                    status(
                        "Could not understand that voltage. "
                        "Enter a number with or without kV."
                    )
                    continue

                return (
                    "manual",
                    value_kv,
                    0,
                    "",
                    f"User-entered {target_label}",
                )

        if choice.isdigit():
            index = int(choice)

            if 1 <= index <= len(filtered):
                candidate = filtered[index - 1]

                status("\nSelected candidate:")
                status(f"  Reported: {candidate.reported_text}")
                status(f"  Value: {candidate.value_kv:.3f} kV")
                status(f"  Context: {candidate.context}")

                confirmation = input(
                    f"Use this as the {target_label}? [y/N]: "
                ).strip().lower()

                if confirmation == "y":
                    return (
                        "automatic candidate",
                        candidate.value_kv,
                        candidate.page_number,
                        candidate.keyword,
                        candidate.context,
                    )

                continue

        direct_value = parse_manual_voltage(choice)

        if direct_value is not None:
            return (
                "manual",
                direct_value,
                0,
                "",
                f"User-entered {target_label}",
            )

        status(
            "Invalid selection. Enter a candidate number, a kV value, "
            "M, R, or N/A."
        )


# ---------------------------------------------------------------------------
# OCR fallback
# ---------------------------------------------------------------------------

def ocr_pages_for_all_fields(
    pdf_path: Path,
    page_numbers: List[int],
) -> Tuple[List[CreepageCandidate], List[VoltageCandidate]]:
    if fitz is None or Image is None or pytesseract is None:
        raise RuntimeError(
            "OCR requires pymupdf, pillow, and pytesseract. Run: "
            "py -m pip install pymupdf pillow pytesseract"
        )

    if TESSERACT_EXE:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

    creep_candidates: List[CreepageCandidate] = []
    voltage_candidates: List[VoltageCandidate] = []
    matrix = fitz.Matrix(OCR_DPI / 72.0, OCR_DPI / 72.0)

    with fitz.open(pdf_path) as document:
        for page_number in page_numbers:
            status(f"OCR processing page {page_number}...")

            page = document[page_number - 1]
            pixmap = page.get_pixmap(
                matrix=matrix,
                alpha=False,
            )

            image = Image.frombytes(
                "RGB",
                [pixmap.width, pixmap.height],
                pixmap.samples,
            )

            text = pytesseract.image_to_string(
                image,
                config="--psm 11",
            )

            lines = text.splitlines()

            for line_index, raw_line in enumerate(lines):
                creep_match = CREEPAGE_KEYWORD_REGEX.search(raw_line)

                if creep_match:
                    keyword = creep_match.group(0)
                    same_line_tail = raw_line[creep_match.end():]

                    add_dual_unit_candidates(
                        creep_candidates,
                        same_line_tail,
                        page_number,
                        keyword,
                        "OCR exact keyword line",
                        1900.0,
                    )

                    add_explicit_unit_candidates(
                        creep_candidates,
                        same_line_tail,
                        page_number,
                        keyword,
                        "OCR exact keyword line",
                        1700.0,
                    )

                    for offset in (1, 2, 3):
                        next_index = line_index + offset

                        if next_index >= len(lines):
                            break

                        continuation_text = lines[next_index].strip()

                        add_dual_unit_candidates(
                            creep_candidates,
                            continuation_text,
                            page_number,
                            keyword,
                            f"OCR continuation line +{offset}",
                            1500.0 - offset * 100.0,
                        )

                        add_explicit_unit_candidates(
                            creep_candidates,
                            continuation_text,
                            page_number,
                            keyword,
                            f"OCR continuation line +{offset}",
                            1300.0 - offset * 100.0,
                        )

                combined_match = COMBINED_VOLTAGE_KEYWORD_REGEX.search(raw_line)
                rated_match = RATED_VOLTAGE_KEYWORD_REGEX.search(raw_line)
                bil_match = BIL_KEYWORD_REGEX.search(raw_line)

                if combined_match:
                    keyword_match = combined_match
                    voltage_type = "combined"
                elif bil_match:
                    keyword_match = bil_match
                    voltage_type = "bil_voltage"
                elif rated_match:
                    keyword_match = rated_match
                    voltage_type = "rated_voltage"
                else:
                    continue

                keyword = keyword_match.group(0)
                same_line_tail = raw_line[keyword_match.end():]

                if voltage_type == "combined":
                    voltage_candidates.extend(
                        infer_combined_voltage_candidates(
                            same_line_tail,
                            page_number,
                            keyword,
                            "OCR exact keyword line",
                            1800.0,
                        )
                    )
                else:
                    add_voltage_candidates_from_text(
                        voltage_candidates,
                        same_line_tail,
                        page_number,
                        keyword,
                        voltage_type,
                        "OCR exact keyword line",
                        1800.0,
                    )

                for offset in (1, 2, 3):
                    next_index = line_index + offset

                    if next_index >= len(lines):
                        break

                    continuation_text = lines[next_index].strip()

                    if voltage_type == "combined":
                        voltage_candidates.extend(
                            infer_combined_voltage_candidates(
                                continuation_text,
                                page_number,
                                keyword,
                                f"OCR continuation line +{offset}",
                                1500.0 - offset * 100.0,
                            )
                        )
                    else:
                        add_voltage_candidates_from_text(
                            voltage_candidates,
                            continuation_text,
                            page_number,
                            keyword,
                            voltage_type,
                            f"OCR continuation line +{offset}",
                            1500.0 - offset * 100.0,
                        )

    return (
        deduplicate_creepage_candidates(creep_candidates),
        deduplicate_voltage_candidates(voltage_candidates),
    )


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

def append_csv(
    pdf_path: Path,
    top_result: Optional[
        Tuple[str, float, float, int, str, str]
    ],
    bottom_result: Optional[
        Tuple[str, float, float, int, str, str]
    ],
    rated_voltage_result: Optional[
        Tuple[str, float, int, str, str]
    ],
    bil_voltage_result: Optional[
        Tuple[str, float, int, str, str]
    ],
) -> Path:
    output_path = Path(OUTPUT_CSV)
    exists = output_path.exists()

    with output_path.open(
        "a",
        newline="",
        encoding="utf-8-sig",
    ) as file:
        writer = csv.writer(file)

        if not exists:
            writer.writerow([
                "pdf_file",
                "pdf_path",

                "top_status",
                "top_selection_method",
                "top_creepage_distance_mm",
                "top_creepage_distance_in",
                "top_page_number",
                "top_matched_keyword",
                "top_context",

                "bottom_status",
                "bottom_selection_method",
                "bottom_creepage_distance_mm",
                "bottom_creepage_distance_in",
                "bottom_page_number",
                "bottom_matched_keyword",
                "bottom_context",

                "rated_voltage_status",
                "rated_voltage_selection_method",
                "rated_voltage_kv",
                "rated_voltage_page_number",
                "rated_voltage_matched_keyword",
                "rated_voltage_context",

                "bil_voltage_status",
                "bil_voltage_selection_method",
                "bil_voltage_kv",
                "bil_voltage_page_number",
                "bil_voltage_matched_keyword",
                "bil_voltage_context",
            ])

        if top_result is None:
            top_row = ["N/A", "", "", "", "", "", ""]
        else:
            (
                method,
                value_mm,
                value_in,
                page_number,
                keyword,
                context,
            ) = top_result
            top_row = [
                "confirmed",
                method,
                value_mm,
                value_in,
                page_number,
                keyword,
                context,
            ]

        if bottom_result is None:
            bottom_row = ["N/A", "", "", "", "", "", ""]
        else:
            (
                method,
                value_mm,
                value_in,
                page_number,
                keyword,
                context,
            ) = bottom_result
            bottom_row = [
                "confirmed",
                method,
                value_mm,
                value_in,
                page_number,
                keyword,
                context,
            ]

        if rated_voltage_result is None:
            rated_row = ["N/A", "", "", "", "", ""]
        else:
            (
                method,
                value_kv,
                page_number,
                keyword,
                context,
            ) = rated_voltage_result
            rated_row = [
                "confirmed",
                method,
                value_kv,
                page_number,
                keyword,
                context,
            ]

        if bil_voltage_result is None:
            bil_row = ["N/A", "", "", "", "", ""]
        else:
            (
                method,
                value_kv,
                page_number,
                keyword,
                context,
            ) = bil_voltage_result
            bil_row = [
                "confirmed",
                method,
                value_kv,
                page_number,
                keyword,
                context,
            ]

        writer.writerow([
            pdf_path.name,
            str(pdf_path),
            *top_row,
            *bottom_row,
            *rated_row,
            *bil_row,
        ])

    return output_path


def main() -> Optional[Path]:
    status("\nDrawing extractor — creepage and voltage ratings")
    status("------------------------------------------------")

    pdf_path = select_pdf()

    if pdf_path is None:
        status("No PDF selected.")
        return None

    status(f"\nSelected PDF: {pdf_path}")

    try:
        creepage_candidates, creepage_keyword_pages = (
            extract_creepage_candidates(pdf_path)
        )
        voltage_candidates, voltage_keyword_pages = (
            extract_voltage_candidates(pdf_path)
        )
    except Exception as exc:
        status(f"\npdfplumber extraction failed: {exc}")
        input("\nPress Enter to close.")
        return

    if not creepage_candidates or not voltage_candidates:
        relevant_pages = sorted(
            set(creepage_keyword_pages + voltage_keyword_pages)
        )

        if not relevant_pages and fitz is not None:
            try:
                with fitz.open(pdf_path) as document:
                    relevant_pages = list(range(1, len(document) + 1))
            except Exception:
                relevant_pages = []

        run_ocr = ask_yes_no(
            "OCR fallback",
            "One or more fields could not be extracted reliably "
            "with pdfplumber.\n\n"
            "Run OCR on the relevant page(s)?",
        )

        if run_ocr and relevant_pages:
            try:
                ocr_creep, ocr_voltage = ocr_pages_for_all_fields(
                    pdf_path,
                    relevant_pages,
                )

                creepage_candidates = deduplicate_creepage_candidates(
                    creepage_candidates + ocr_creep
                )
                voltage_candidates = deduplicate_voltage_candidates(
                    voltage_candidates + ocr_voltage
                )

            except Exception as exc:
                status(f"OCR failed: {exc}")

    status("\nCreating highlighted PDF for visual confirmation...")
    pdf_to_open, highlight_count = create_highlighted_copy(
        pdf_path
    )

    if highlight_count:
        status(
            f"Opening highlighted copy with {highlight_count} "
            f"highlight(s)..."
        )
    else:
        status(
            "Opening the original PDF because no highlighted copy "
            "was available."
        )

    open_pdf(pdf_to_open)

    if creepage_candidates:
        status(
            "\nThe same creepage candidate list can be used independently "
            "for TOP and BOTTOM selections."
        )
        print_creepage_candidates(
            creepage_candidates,
            "TOP/BOTTOM",
        )
    else:
        status(
            "\nNo automatic creepage candidates were found. "
            "Manual entry remains available."
        )

    if voltage_candidates:
        print_voltage_candidates(
            voltage_candidates,
            "rated_voltage",
            "Rated voltage",
        )
        print_voltage_candidates(
            voltage_candidates,
            "bil_voltage",
            "Basic Impulse Level (BIL)",
        )
    else:
        status(
            "\nNo automatic voltage candidates were found. "
            "Manual entry remains available."
        )

    top_result = prompt_for_creepage(
        creepage_candidates,
        "TOP",
    )

    bottom_result = prompt_for_creepage(
        creepage_candidates,
        "BOTTOM",
    )

    rated_voltage_result = prompt_for_voltage(
        voltage_candidates,
        "rated_voltage",
        "Rated voltage",
    )

    bil_voltage_result = prompt_for_voltage(
        voltage_candidates,
        "bil_voltage",
        "Basic Impulse Level (BIL)",
    )

    output_path = append_csv(
        pdf_path,
        top_result,
        bottom_result,
        rated_voltage_result,
        bil_voltage_result,
    )

    status("\nConfirmed results")
    status("-----------------")

    if top_result is None:
        status("TOP creepage distance: N/A")
    else:
        status(
            f"TOP creepage distance: "
            f"{top_result[1]:.3f} mm "
            f"({top_result[2]:.4f} in)"
        )

    if bottom_result is None:
        status("BOTTOM creepage distance: N/A")
    else:
        status(
            f"BOTTOM creepage distance: "
            f"{bottom_result[1]:.3f} mm "
            f"({bottom_result[2]:.4f} in)"
        )

    if rated_voltage_result is None:
        status("Rated voltage: N/A")
    else:
        status(
            f"Rated voltage: {rated_voltage_result[1]:.3f} kV"
        )

    if bil_voltage_result is None:
        status("Basic Impulse Level (BIL): N/A")
    else:
        status(
            f"Basic Impulse Level (BIL): "
            f"{bil_voltage_result[1]:.3f} kV"
        )

    status(f"\nSaved to: {output_path.resolve()}")
    return output_path.resolve()
    input("\nPress Enter to close.")


if __name__ == "__main__":
    main()