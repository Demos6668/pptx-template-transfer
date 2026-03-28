#!/usr/bin/env python3
"""PPTX Template Transfer — apply one deck's visual design to another's content.

Usage:
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --mode design
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --mode layout
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --verbose

Modes:
    design  — Clone template slides as visual skeletons, inject content text (default
              when template layouts are blank/default).
    layout  — Transfer theme + masters + layouts between files (default when template
              has named layouts with placeholders).
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import sys
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Pt

# ============================================================================
# GLOBALS / CONFIG
# ============================================================================

VERBOSE = False

_NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

_FOOTER_PATTERNS = re.compile(
    r"(?i)(page\s*\d+|confidential|©|\bcopyright\b|\ball rights reserved\b"
    r"|\b\d{4}[-/]\d{2}[-/]\d{2}\b|\b\d{2}/\d{2}/\d{4}\b"
    r"|proprietary|internal use|draft|do not distribute)",
)

_PAGE_NUM_PATTERN = re.compile(r"(?i)page\s*\d+")
_DATE_PATTERN = re.compile(
    r"\b(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}/\d{1,2}/\d{4})\b"
)


def _log(msg: str) -> None:
    if VERBOSE:
        print(msg)


# ============================================================================
# SHARED HELPERS
# ============================================================================

def _text_of(shape) -> str:
    """Extract full text from a shape, or empty string."""
    if not shape.has_text_frame:
        return ""
    return shape.text_frame.text.strip()


def _word_count(text: str) -> int:
    return len(text.split()) if text else 0


def _max_font_pt(shape) -> float:
    """Return the maximum font size in pt found in any run of the shape."""
    mx = 0.0
    if not shape.has_text_frame:
        return mx
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None:
                mx = max(mx, run.font.size.pt)
    return mx


def _min_font_pt(shape) -> float:
    """Return the minimum non-zero font size in pt."""
    mn = 999.0
    if not shape.has_text_frame:
        return 0.0
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None and run.font.size.pt > 0:
                mn = min(mn, run.font.size.pt)
    return mn if mn < 999.0 else 0.0


def _shape_area(shape) -> int:
    w = shape.width or 0
    h = shape.height or 0
    return w * h


def _shape_area_pct(shape, slide_w: int, slide_h: int) -> float:
    total = slide_w * slide_h
    if total == 0:
        return 0.0
    return _shape_area(shape) / total * 100.0


def _is_picture(shape) -> bool:
    try:
        return shape.shape_type == MSO_SHAPE_TYPE.PICTURE or shape.shape_type == 13
    except Exception:
        return False


def _is_table(shape) -> bool:
    return hasattr(shape, "has_table") and shape.has_table


def _is_chart(shape) -> bool:
    return hasattr(shape, "has_chart") and shape.has_chart


def _is_group(shape) -> bool:
    try:
        return shape.shape_type == MSO_SHAPE_TYPE.GROUP
    except Exception:
        return False


def _is_ole_or_embedded(shape) -> bool:
    """Check if shape is an OLE/embedded object."""
    try:
        tag = etree.QName(shape._element.tag).localname
        return tag == "graphicFrame" and not _is_table(shape) and not _is_chart(shape)
    except Exception:
        return False


def _shape_bottom(shape, slide_h: int) -> float:
    """Return shape bottom edge as fraction of slide height."""
    top = shape.top or 0
    height = shape.height or 0
    if slide_h == 0:
        return 0.0
    return (top + height) / slide_h


def _shape_top_frac(shape, slide_h: int) -> float:
    """Return shape top edge as fraction of slide height."""
    if slide_h == 0:
        return 0.0
    return (shape.top or 0) / slide_h


def _shape_left_frac(shape, slide_w: int) -> float:
    if slide_w == 0:
        return 0.0
    return (shape.left or 0) / slide_w


def _is_allcaps_short(text: str) -> bool:
    """Check if text is ALL-CAPS and ≤ 5 words."""
    words = text.split()
    if len(words) == 0 or len(words) > 5:
        return False
    alpha = "".join(c for c in text if c.isalpha())
    if not alpha:
        return False
    return alpha == alpha.upper()


def _dominant_text_color(shape) -> str | None:
    """Extract the most common text color from a shape as hex string."""
    if not shape.has_text_frame:
        return None
    colors: dict[str, int] = {}
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            try:
                c = run.font.color
                if c and c.type is not None and c.rgb:
                    key = str(c.rgb)
                    colors[key] = colors.get(key, 0) + len(run.text)
            except AttributeError:
                pass
    if not colors:
        return None
    return max(colors, key=colors.get)


# ============================================================================
# A. SHAPE ROLE CLASSIFICATION ENGINE
# ============================================================================

def classify_shape_role(
    shape,
    slide,
    slide_w: int,
    slide_h: int,
    *,
    _largest_font: float | None = None,
    _title_assigned: bool = False,
    _body_count: int = 0,
    _info_count: int = 0,
    _similar_shapes: set | None = None,
) -> str:
    """Classify a shape's role on a slide.

    Returns: "title" | "body" | "info" | "decorative" | "footer" | "media"

    Args:
        shape: The shape to classify.
        slide: The slide containing the shape (for context).
        slide_w, slide_h: Slide dimensions in EMU.
        _largest_font: Pre-computed largest font on the slide.
        _title_assigned: Whether a title has already been assigned.
        _body_count: How many body zones already assigned.
        _info_count: How many info zones already assigned.
        _similar_shapes: Set of shape ids that are part of repeated patterns.
    """
    text = _text_of(shape)
    area_pct = _shape_area_pct(shape, slide_w, slide_h)
    font_size = _max_font_pt(shape)
    wc = _word_count(text)
    top_frac = _shape_top_frac(shape, slide_h)
    bottom_frac = _shape_bottom(shape, slide_h)
    left_frac = _shape_left_frac(shape, slide_w)

    # --- MEDIA ---
    if _is_picture(shape) or _is_chart(shape) or _is_table(shape):
        return "media"
    if _is_ole_or_embedded(shape):
        return "media"
    if _is_group(shape):
        # Group with primarily non-text elements → media
        return "media"

    # --- FOOTER / HEADER ---
    # Bottom 10% of slide
    if bottom_frac >= 0.90 and area_pct < 5:
        return "footer"
    # Top 8% and small area
    if top_frac <= 0.08 and area_pct < 3 and wc <= 10:
        return "footer"
    # Footer text patterns
    if text and _FOOTER_PATTERNS.search(text):
        return "footer"

    # --- NO TEXT → DECORATIVE ---
    if not text:
        return "decorative"

    # --- DECORATIVE checks ---
    # Very small shapes
    if area_pct < 2 and wc <= 5:
        return "decorative"
    # All text ≤ 10pt
    if font_size <= 10 and font_size > 0:
        return "decorative"
    # Very few words (labels, numbering)
    if wc <= 3:
        return "decorative"
    # ALL-CAPS short text (branding labels)
    if _is_allcaps_short(text) and area_pct < 5:
        return "decorative"
    # Icon/bullet numbering ("01", "02", etc.)
    if re.match(r"^\d{1,2}$", text.strip()):
        return "decorative"
    # Part of repeated pattern
    if _similar_shapes and id(shape) in _similar_shapes:
        return "decorative"

    # --- TITLE ---
    if (not _title_assigned
            and top_frac < 0.45
            and font_size >= 20
            and _largest_font is not None
            and font_size >= _largest_font - 2  # within 2pt of largest
            and wc <= 20):
        return "title"

    # --- BODY ---
    if _body_count < 2 and area_pct > 4 and wc > 10:
        return "body"
    # Slightly relaxed: decent area and moderate text
    if _body_count < 2 and area_pct > 3 and wc > 5:
        return "body"

    # --- INFO (sidebar/panel) ---
    if (_info_count < 1
            and left_frac >= 0.55
            and 5 <= wc <= 50
            and area_pct > 2):
        return "info"

    # --- DEFAULT: decorative ---
    return "decorative"


def _detect_repeated_patterns(shapes, slide_w: int, slide_h: int) -> set:
    """Find shapes that are part of repeated visual patterns (grids, rows).

    Returns set of shape ids that belong to repeated patterns.
    """
    result = set()
    if len(shapes) < 3:
        return result

    # Group shapes by similar dimensions (width/height within 15%)
    dimension_groups: dict[tuple, list] = {}
    for shape in shapes:
        w = shape.width or 0
        h = shape.height or 0
        if w == 0 or h == 0:
            continue
        # Bucket by rounded dimensions
        bucket_w = round(w / (slide_w * 0.02)) if slide_w else 0
        bucket_h = round(h / (slide_h * 0.02)) if slide_h else 0
        key = (bucket_w, bucket_h)
        dimension_groups.setdefault(key, []).append(shape)

    for group in dimension_groups.values():
        if len(group) < 3:
            continue
        # Check if arranged in a row (similar Y) or grid
        tops = [(shape.top or 0) for shape in group]
        lefts = [(shape.left or 0) for shape in group]

        # Row detection: most shapes at similar Y
        top_buckets: dict[int, int] = {}
        for t in tops:
            bucket = round(t / (slide_h * 0.05)) if slide_h else 0
            top_buckets[bucket] = top_buckets.get(bucket, 0) + 1

        max_in_row = max(top_buckets.values()) if top_buckets else 0
        if max_in_row >= 3:
            for shape in group:
                # Only mark small-text shapes as repeated pattern
                if _word_count(_text_of(shape)) <= 15:
                    result.add(id(shape))

    return result


def classify_all_shapes(slide, slide_w: int, slide_h: int) -> list[tuple[Any, str]]:
    """Classify all shapes on a slide.

    Returns list of (shape, role) tuples.
    """
    shapes = list(slide.shapes)

    # Pre-compute largest font on the slide
    largest_font = 0.0
    for s in shapes:
        fs = _max_font_pt(s)
        if fs > largest_font:
            largest_font = fs

    # Detect repeated patterns
    similar = _detect_repeated_patterns(shapes, slide_w, slide_h)

    # First pass: classify all shapes
    title_assigned = False
    body_count = 0
    info_count = 0

    # Sort shapes by priority for classification: top-to-bottom, then left-to-right
    sorted_shapes = sorted(shapes, key=lambda s: ((s.top or 0), (s.left or 0)))

    results = {}
    for shape in sorted_shapes:
        role = classify_shape_role(
            shape, slide, slide_w, slide_h,
            _largest_font=largest_font,
            _title_assigned=title_assigned,
            _body_count=body_count,
            _info_count=info_count,
            _similar_shapes=similar,
        )
        results[id(shape)] = role
        if role == "title":
            title_assigned = True
        elif role == "body":
            body_count += 1
        elif role == "info":
            info_count += 1

    # Return in original order
    return [(s, results[id(s)]) for s in shapes]


# ============================================================================
# B. CONTENT STRUCTURE EXTRACTOR
# ============================================================================

@dataclass
class ParagraphData:
    text: str
    level: int = 0
    bold: bool = False
    italic: bool = False
    font_size: float = 0.0


@dataclass
class ContentData:
    title: str = ""
    body_paragraphs: list[ParagraphData] = field(default_factory=list)
    tables: list[Any] = field(default_factory=list)  # (table_element, rows_text)
    images: list[Any] = field(default_factory=list)   # (blob, width, height, left, top)
    has_chart: bool = False
    slide_type: str = "content"
    word_count: int = 0
    primary_color: str | None = None


def _classify_slide_type(slide, slide_index: int, total_slides: int) -> str:
    """Classify a slide as: title, content, section, data, closing, image, blank."""
    texts = []
    images = 0
    tables = 0
    charts = 0

    for shape in slide.shapes:
        if _is_picture(shape):
            images += 1
        if _is_table(shape):
            tables += 1
        if _is_chart(shape):
            charts += 1
        t = _text_of(shape)
        if t:
            texts.append({
                "text": t,
                "size": _max_font_pt(shape),
                "words": _word_count(t),
            })

    total_words = sum(t["words"] for t in texts)
    big_texts = [t for t in texts if t["size"] >= 20]

    if not texts and images == 0:
        return "blank"
    if not texts and images > 0:
        return "image"

    if slide_index == 0 and big_texts:
        return "title"
    if total_words <= 20 and big_texts and len(texts) <= 5:
        return "title"

    if slide_index == total_slides - 1 and total_words <= 40:
        return "closing"

    if len(texts) <= 3 and total_words <= 15 and big_texts:
        return "section"

    if tables > 0 or charts > 0:
        return "data"

    if images >= 3 and total_words < 30:
        return "image"

    return "content"


def _extract_paragraphs_from_shape(shape) -> list[ParagraphData]:
    """Extract structured paragraph data from a shape."""
    result = []
    if not shape.has_text_frame:
        return result
    for para in shape.text_frame.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        level = para.level if para.level else 0
        bold = False
        italic = False
        font_size = 0.0
        for run in para.runs:
            if run.font.bold:
                bold = True
            if run.font.italic:
                italic = True
            if run.font.size:
                font_size = max(font_size, run.font.size.pt)
        result.append(ParagraphData(
            text=text, level=level, bold=bold,
            italic=italic, font_size=font_size,
        ))
    return result


def _extract_table_data(shape) -> list[list[str]]:
    """Extract cell text from a table shape as a 2D list."""
    if not _is_table(shape):
        return []
    table = shape.table
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cells.append(cell.text.strip())
        rows.append(cells)
    return rows


def _extract_image_data(shape, slide) -> tuple[bytes, int, int, int, int] | None:
    """Extract image blob and dimensions from a picture shape."""
    if not _is_picture(shape):
        return None
    try:
        image = shape.image
        blob = image.blob
        return (blob, shape.width, shape.height, shape.left or 0, shape.top or 0)
    except Exception:
        return None


def extract_content(
    slide,
    slide_index: int,
    total_slides: int,
    slide_w: int,
    slide_h: int,
) -> ContentData:
    """Extract structured content from a content slide."""
    content = ContentData()
    content.slide_type = _classify_slide_type(slide, slide_index, total_slides)

    shapes = list(slide.shapes)

    # --- Title detection ---
    # Find all text shapes, sort by (font_size DESC, top ASC)
    text_shapes = [(s, _max_font_pt(s), _text_of(s)) for s in shapes if _text_of(s)]
    text_shapes.sort(key=lambda x: (-x[1], (x[0].top or 0)))

    title_shape = None
    for s, fs, txt in text_shapes:
        if _word_count(txt) <= 15 and fs >= 20:
            title_shape = s
            content.title = txt
            break
    # Fallback: topmost shape with ≤ 10 words
    if not title_shape and text_shapes:
        for s, fs, txt in sorted(text_shapes, key=lambda x: (x[0].top or 0)):
            if _word_count(txt) <= 10:
                title_shape = s
                content.title = txt
                break

    # --- Body extraction ---
    # All non-title text shapes, ordered top-to-bottom, left-to-right
    body_shapes = [
        s for s in shapes
        if s is not title_shape and _text_of(s) and not _is_table(s) and not _is_chart(s)
    ]
    body_shapes.sort(key=lambda s: ((s.top or 0), (s.left or 0)))

    for shape in body_shapes:
        paras = _extract_paragraphs_from_shape(shape)
        # Detect subheadings: bold text or font >= 18pt followed by smaller text
        for p in paras:
            if p.bold or (p.font_size >= 18 and _word_count(p.text) <= 10):
                p.bold = True  # Mark as subheading
            content.body_paragraphs.append(p)

    # --- Tables ---
    for shape in shapes:
        if _is_table(shape):
            table_text = _extract_table_data(shape)
            # Also save the element for potential deep-copy later
            content.tables.append({
                "data": table_text,
                "rows": len(table_text),
                "cols": len(table_text[0]) if table_text else 0,
                "element": deepcopy(shape._element),
                "width": shape.width,
                "height": shape.height,
                "left": shape.left,
                "top": shape.top,
            })

    # --- Images ---
    for shape in shapes:
        if _is_picture(shape):
            img_data = _extract_image_data(shape, slide)
            if img_data:
                area_pct = _shape_area_pct(shape, slide_w, slide_h)
                # Only grab content images (>10% of slide, likely charts/screenshots)
                if area_pct > 10:
                    content.images.append(img_data)

    # --- Charts ---
    for shape in shapes:
        if _is_chart(shape):
            content.has_chart = True

    # --- Word count and color ---
    all_text = content.title + " " + " ".join(p.text for p in content.body_paragraphs)
    content.word_count = _word_count(all_text)

    # Dominant text color
    colors: dict[str, int] = {}
    for shape in shapes:
        c = _dominant_text_color(shape)
        if c:
            colors[c] = colors.get(c, 0) + 1
    if colors:
        content.primary_color = max(colors, key=colors.get)

    return content


# ============================================================================
# C. SMART SLIDE MATCHING
# ============================================================================

def _classify_template_structure(
    slide,
    slide_w: int,
    slide_h: int,
    slide_index: int = -1,
    total_slides: int = -1,
) -> str:
    """Classify a template slide's STRUCTURE type.

    Returns: "narrative" | "list" | "grid" | "data" | "visual" | "title" | "section" | "closing"
    """
    shapes = list(slide.shapes)
    text_shapes = [s for s in shapes if _text_of(s)]
    images = [s for s in shapes if _is_picture(s)]
    tables = [s for s in shapes if _is_table(s)]
    charts = [s for s in shapes if _is_chart(s)]

    total_words = sum(_word_count(_text_of(s)) for s in text_shapes)
    big_texts = [s for s in text_shapes if _max_font_pt(s) >= 20]

    # Data slide
    if tables or charts:
        return "data"

    # Visual: image-dominated
    if len(images) >= 2 and total_words < 30:
        return "visual"

    # Detect numbered/sequential items
    numbered_shapes = [s for s in text_shapes if re.match(r"^\d{1,2}$", _text_of(s).strip())]
    if len(numbered_shapes) >= 3:
        similar = _detect_repeated_patterns(shapes, slide_w, slide_h)
        if similar:
            return "grid"
        return "list"

    # Position-based: first slide is almost always a title/cover slide
    if slide_index == 0 and big_texts:
        return "title"

    # Last slide is often closing
    closing_keywords = {"thank", "contact", "questions", "q&a", "reference"}
    all_text_lower = " ".join(_text_of(s) for s in text_shapes).lower()
    if slide_index == total_slides - 1 and total_slides > 1:
        if any(kw in all_text_lower for kw in closing_keywords) or total_words <= 40:
            return "closing"

    # Title slide: big centered text, minimal other content
    if big_texts and total_words <= 20 and len(text_shapes) <= 5:
        return "title"

    # Section: large text, few elements
    if big_texts and len(text_shapes) <= 3 and total_words <= 15:
        return "section"

    # Closing patterns anywhere
    if any(kw in all_text_lower for kw in closing_keywords) and total_words <= 40:
        return "closing"

    # Narrative vs list: check if there's one large body area
    body_shapes = [s for s in text_shapes if _shape_area_pct(s, slide_w, slide_h) > 4 and _word_count(_text_of(s)) > 10]
    if len(body_shapes) >= 1:
        return "narrative"

    if total_words > 20:
        return "narrative"
    return "section"


_TYPE_COMPAT = {
    # (content_type, template_structure) -> score (0-40)
    ("title", "title"): 40,
    ("title", "section"): 20,
    ("content", "narrative"): 40,
    ("content", "list"): 30,
    ("content", "grid"): 25,
    ("content", "data"): 15,
    ("section", "section"): 40,
    ("section", "title"): 25,
    ("data", "data"): 40,
    ("data", "narrative"): 20,
    ("data", "grid"): 25,
    ("closing", "closing"): 40,
    ("closing", "section"): 20,
    ("image", "visual"): 40,
    ("image", "narrative"): 15,
    ("blank", "section"): 10,
}


def _match_score_v2(
    content_type: str,
    template_struct: str,
    content_idx: int,
    template_idx: int,
    content_total: int,
    template_total: int,
    content_words: int,
    template_capacity: int,
    content_has_table: bool,
    template_has_table: bool,
    content_para_count: int,
    template_is_list: bool,
) -> float:
    """Score how well a content slide matches a template slide (v2)."""
    score = 0.0

    # --- Type compatibility (0-40) ---
    key = (content_type, template_struct)
    score += _TYPE_COMPAT.get(key, 5)

    # --- Text density fit (0-25) ---
    if template_capacity > 0 and content_words > 0:
        ratio = min(content_words, template_capacity) / max(content_words, template_capacity)
        score += 25 * ratio
    elif content_words == 0 and template_capacity <= 10:
        score += 20  # Both low-text

    # --- Content structure fit (0-20) ---
    if content_has_table and template_has_table:
        score += 20
    elif content_has_table and not template_has_table:
        score += 5  # Penalty: no table slot
    elif content_para_count >= 5 and template_is_list:
        score += 15
    elif content_para_count >= 3 and template_struct == "narrative":
        score += 15
    elif content_para_count < 3 and template_struct in ("section", "title"):
        score += 10

    # --- Position preference (0-15) ---
    if content_idx == 0 and template_struct == "title":
        score += 15
    elif content_idx == content_total - 1 and template_struct == "closing":
        score += 15
    elif content_total > 1 and template_total > 1:
        c_pos = content_idx / (content_total - 1)
        t_pos = template_idx / (template_total - 1)
        score += 15 * (1 - abs(c_pos - t_pos))

    return score


def build_slide_mapping(
    content_prs: Presentation,
    template_prs: Presentation,
    content_data_list: list[ContentData] | None = None,
) -> list[int]:
    """For each content slide, return the index of the best-matching template slide.

    Enforces variety: no template slide gets more than 40% of mappings.
    """
    sw, sh = template_prs.slide_width, template_prs.slide_height
    ct = len(content_prs.slides)
    tt = len(template_prs.slides)

    # Analyze template slides
    t_info = []
    for i, slide in enumerate(template_prs.slides):
        struct = _classify_template_structure(slide, sw, sh, i, tt)
        words = sum(_word_count(_text_of(s)) for s in slide.shapes)
        has_table = any(_is_table(s) for s in slide.shapes)
        is_list = struct in ("list", "grid")
        t_info.append({
            "struct": struct, "words": words,
            "has_table": has_table, "is_list": is_list,
        })

    # Content data
    if content_data_list is None:
        content_data_list = []
        for i, slide in enumerate(content_prs.slides):
            content_data_list.append(extract_content(
                slide, i, ct,
                content_prs.slide_width, content_prs.slide_height,
            ))

    # Score matrix: content_idx -> [(template_idx, score), ...]
    score_matrix = []
    for ci, cd in enumerate(content_data_list):
        scores = []
        for ti, tinfo in enumerate(t_info):
            sc = _match_score_v2(
                cd.slide_type, tinfo["struct"],
                ci, ti, ct, tt,
                cd.word_count, tinfo["words"],
                len(cd.tables) > 0, tinfo["has_table"],
                len(cd.body_paragraphs), tinfo["is_list"],
            )
            scores.append((ti, sc))
        scores.sort(key=lambda x: -x[1])
        score_matrix.append(scores)

    # Greedy assignment with variety enforcement
    usage_count: dict[int, int] = {i: 0 for i in range(tt)}
    max_per_template = max(2, math.ceil(ct * 0.4))
    # Enforce using more templates when available
    min_distinct = min(tt, max(3, math.ceil(ct / 3)))

    mapping = []
    for ci, scores in enumerate(score_matrix):
        best_idx = scores[0][0]
        best_score = scores[0][1]

        # Check variety constraint
        if usage_count[best_idx] >= max_per_template:
            for ti, sc in scores[1:]:
                if usage_count[ti] < max_per_template:
                    best_idx = ti
                    best_score = sc
                    break

        mapping.append(best_idx)
        usage_count[best_idx] = usage_count.get(best_idx, 0) + 1

    # Second pass: if we haven't used enough distinct templates, redistribute
    used_set = {ti for ti, c in usage_count.items() if c > 0}
    unused = [ti for ti in range(tt) if usage_count[ti] == 0]
    if len(used_set) < min_distinct and unused:
        # Find overused templates and reassign some of their slides
        overused = sorted(
            [(ti, c) for ti, c in usage_count.items() if c > 1],
            key=lambda x: -x[1],
        )
        for u_ti in unused:
            if not overused:
                break
            # Find the best content slide to reassign to this unused template
            donor_ti, donor_count = overused[0]
            # Among slides mapped to donor, find the one with best score for u_ti
            candidates = [(ci, score_matrix[ci]) for ci in range(ct) if mapping[ci] == donor_ti]
            best_reassign = None
            best_reassign_score = -1
            for ci, scores in candidates:
                for ti, sc in scores:
                    if ti == u_ti and sc > best_reassign_score:
                        best_reassign = ci
                        best_reassign_score = sc
            if best_reassign is not None and best_reassign_score > 10:
                old_ti = mapping[best_reassign]
                mapping[best_reassign] = u_ti
                usage_count[old_ti] -= 1
                usage_count[u_ti] = usage_count.get(u_ti, 0) + 1
                # Update overused list
                overused = sorted(
                    [(ti, c) for ti, c in usage_count.items() if c > 1],
                    key=lambda x: -x[1],
                )

    for ci in range(ct):
        cd = content_data_list[ci]
        ti = mapping[ci]
        # Find score for this mapping
        score_val = next((sc for t, sc in score_matrix[ci] if t == ti), 0)
        _log(f"  Content slide {ci+1} ({cd.slide_type}, {cd.word_count}w) "
             f"-> Template slide {ti+1} ({t_info[ti]['struct']}) "
             f"score={score_val:.0f}")

    used = sum(1 for v in usage_count.values() if v > 0)
    _log(f"  Variety: {used}/{tt} template slides used (target: {min_distinct}+)")

    return mapping


# ============================================================================
# SLIDE CLONING
# ============================================================================

def _clone_slide(template_prs: Presentation, src_slide, dst_prs: Presentation):
    """Clone a slide from template_prs into dst_prs, copying all shapes and media.

    Returns the new slide object.
    """
    dst_layout = dst_prs.slide_layouts[0]
    src_layout_name = src_slide.slide_layout.name
    for layout in dst_prs.slide_layouts:
        if layout.name == src_layout_name:
            dst_layout = layout
            break

    new_slide = dst_prs.slides.add_slide(dst_layout)

    # Remove all default placeholder shapes
    spTree = new_slide.shapes._spTree
    for sp in list(spTree):
        tag = etree.QName(sp.tag).localname if isinstance(sp.tag, str) else ""
        if tag in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame"):
            spTree.remove(sp)

    # Deep-copy source shapes
    src_spTree = src_slide.shapes._spTree
    for child in src_spTree:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        if tag in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame"):
            spTree.append(deepcopy(child))

    # Copy background
    src_sld = src_slide._element
    dst_sld = new_slide._element
    src_bg = src_sld.find(f'{{{_NSMAP["p"]}}}bg')
    if src_bg is not None:
        dst_bg = dst_sld.find(f'{{{_NSMAP["p"]}}}bg')
        if dst_bg is not None:
            dst_sld.remove(dst_bg)
        new_bg = deepcopy(src_bg)
        cSld = dst_sld.find(f'{{{_NSMAP["p"]}}}cSld')
        if cSld is not None:
            dst_sld.insert(list(dst_sld).index(cSld), new_bg)
        else:
            dst_sld.insert(0, new_bg)

    # Copy relationships and build rId map
    rid_map = {}
    for rel_key, rel in src_slide.part.rels.items():
        rtype = rel.reltype
        if rtype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        try:
            if rel.is_external:
                new_rid = new_slide.part.rels.get_or_add_ext_rel(rtype, rel.target_ref)
            else:
                new_rid = new_slide.part.rels.get_or_add(rtype, rel.target_part)
            rid_map[rel_key] = new_rid
        except Exception:
            pass

    if rid_map:
        _update_rids_in_tree(spTree, rid_map)
        dst_bg2 = dst_sld.find(f'{{{_NSMAP["p"]}}}bg')
        if dst_bg2 is not None:
            _update_rids_in_tree(dst_bg2, rid_map)

    return new_slide


def _update_rids_in_tree(element, rid_map: dict[str, str]) -> None:
    """Walk an lxml element tree and replace old rIds with new ones."""
    for el in element.iter():
        for attr_name in list(el.attrib.keys()):
            val = el.attrib[attr_name]
            if val in rid_map:
                el.attrib[attr_name] = rid_map[val]


# ============================================================================
# D. TEXT INJECTION ENGINE
# ============================================================================

def _save_paragraph_format(para_element):
    """Save paragraph-level formatting (pPr) from an lxml <a:p> element."""
    ns_a = _NSMAP["a"]
    pPr = para_element.find(f'{{{ns_a}}}pPr')
    return deepcopy(pPr) if pPr is not None else None


def _save_run_format(para_element):
    """Save run-level formatting (rPr) from the first run in an <a:p>."""
    ns_a = _NSMAP["a"]
    for r in para_element.findall(f'{{{ns_a}}}r'):
        rPr = r.find(f'{{{ns_a}}}rPr')
        if rPr is not None:
            return deepcopy(rPr)
    # Try endParaRPr
    endRPr = para_element.find(f'{{{ns_a}}}endParaRPr')
    if endRPr is not None:
        return deepcopy(endRPr)
    return None


def _inject_text_into_shape(shape, text: str) -> None:
    """Replace text in a shape while preserving the template's formatting."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return

    ns_a = _NSMAP["a"]
    first_p = tf.paragraphs[0]._p

    template_pPr = _save_paragraph_format(first_p)
    template_rPr = _save_run_format(first_p)

    txBody = tf._txBody
    for p in list(txBody.findall(f'{{{ns_a}}}p')):
        txBody.remove(p)

    paragraphs = text.split("\n")
    for para_text in paragraphs:
        new_p = etree.SubElement(txBody, f'{{{ns_a}}}p')
        if template_pPr is not None:
            new_p.append(deepcopy(template_pPr))
        if para_text.strip():
            new_r = etree.SubElement(new_p, f'{{{ns_a}}}r')
            if template_rPr is not None:
                new_r.append(deepcopy(template_rPr))
            new_t = etree.SubElement(new_r, f'{{{ns_a}}}t')
            new_t.text = para_text
        else:
            endRPr = etree.SubElement(new_p, f'{{{ns_a}}}endParaRPr')
            if template_rPr is not None:
                for k, v in template_rPr.attrib.items():
                    endRPr.attrib[k] = v


def _inject_structured_text(shape, paragraphs: list[ParagraphData]) -> None:
    """Inject structured paragraph data into a shape, preserving template formatting.

    Handles subheadings (bold), indentation levels, and italic.
    """
    if not shape.has_text_frame or not paragraphs:
        return

    tf = shape.text_frame
    if not tf.paragraphs:
        return

    ns_a = _NSMAP["a"]
    first_p = tf.paragraphs[0]._p

    template_pPr = _save_paragraph_format(first_p)
    template_rPr = _save_run_format(first_p)

    txBody = tf._txBody
    for p in list(txBody.findall(f'{{{ns_a}}}p')):
        txBody.remove(p)

    for pd in paragraphs:
        new_p = etree.SubElement(txBody, f'{{{ns_a}}}p')

        # Paragraph properties
        if template_pPr is not None:
            pPr = deepcopy(template_pPr)
            # Set indent level
            if pd.level > 0:
                pPr.set("lvl", str(pd.level))
            new_p.append(pPr)
        elif pd.level > 0:
            pPr = etree.SubElement(new_p, f'{{{ns_a}}}pPr')
            pPr.set("lvl", str(pd.level))

        if pd.text.strip():
            new_r = etree.SubElement(new_p, f'{{{ns_a}}}r')
            if template_rPr is not None:
                rPr = deepcopy(template_rPr)
            else:
                rPr = etree.SubElement(new_r, f'{{{ns_a}}}rPr')
                rPr = None  # Will be created fresh below if needed

            if template_rPr is not None:
                # Apply bold/italic overrides
                if pd.bold:
                    rPr.set("b", "1")
                if pd.italic:
                    rPr.set("i", "1")
                new_r.append(rPr)

            new_t = etree.SubElement(new_r, f'{{{ns_a}}}t')
            new_t.text = pd.text
        else:
            endRPr = etree.SubElement(new_p, f'{{{ns_a}}}endParaRPr')
            if template_rPr is not None:
                for k, v in template_rPr.attrib.items():
                    endRPr.attrib[k] = v


def _clear_shape_text(shape) -> None:
    """Remove all text from a shape but keep the shape itself."""
    if not shape.has_text_frame:
        return
    ns_a = _NSMAP["a"]
    txBody = shape.text_frame._txBody
    for p in list(txBody.findall(f'{{{ns_a}}}p')):
        txBody.remove(p)
    new_p = etree.SubElement(txBody, f'{{{ns_a}}}p')
    etree.SubElement(new_p, f'{{{ns_a}}}endParaRPr')


def inject_content(
    cloned_slide,
    content_data: ContentData,
    slide_w: int,
    slide_h: int,
) -> dict[str, Any]:
    """Inject content into a cloned template slide using shape role classification.

    Returns a diagnostic dict for verbose output.
    """
    diag = {
        "shapes": [],
        "injected_title": None,
        "injected_body": None,
        "protected_count": 0,
    }

    classifications = classify_all_shapes(cloned_slide, slide_w, slide_h)

    title_shape = None
    body_shapes = []
    info_shape = None

    for shape, role in classifications:
        diag["shapes"].append({
            "name": shape.name,
            "role": role,
            "area_pct": round(_shape_area_pct(shape, slide_w, slide_h), 1),
            "top_pct": round(_shape_top_frac(shape, slide_h) * 100, 0),
            "text_preview": _text_of(shape)[:40],
        })

        if role == "title" and title_shape is None:
            title_shape = shape
        elif role == "body":
            body_shapes.append(shape)
        elif role == "info" and info_shape is None:
            info_shape = shape
        else:
            diag["protected_count"] += 1

    # --- Inject title ---
    if content_data.title and title_shape:
        _inject_text_into_shape(title_shape, content_data.title)
        diag["injected_title"] = content_data.title[:50]
    elif title_shape:
        _clear_shape_text(title_shape)

    # --- Inject body ---
    if content_data.body_paragraphs and body_shapes:
        if len(body_shapes) == 1:
            _inject_structured_text(body_shapes[0], content_data.body_paragraphs)
            total_words = sum(_word_count(p.text) for p in content_data.body_paragraphs)
            diag["injected_body"] = f"{total_words} words -> 1 zone"
        else:
            # Split paragraphs across body zones
            per_zone = max(1, len(content_data.body_paragraphs) // len(body_shapes))
            idx = 0
            for i, zone in enumerate(body_shapes):
                if i == len(body_shapes) - 1:
                    chunk = content_data.body_paragraphs[idx:]
                else:
                    chunk = content_data.body_paragraphs[idx:idx + per_zone]
                    idx += per_zone
                if chunk:
                    _inject_structured_text(zone, chunk)
                else:
                    _clear_shape_text(zone)
            total_words = sum(_word_count(p.text) for p in content_data.body_paragraphs)
            diag["injected_body"] = f"{total_words} words -> {len(body_shapes)} zones"
    elif body_shapes:
        for s in body_shapes:
            _clear_shape_text(s)

    # --- Inject info (sidebar summary) ---
    if info_shape:
        if content_data.body_paragraphs:
            # First 2-3 paragraphs as summary
            summary = content_data.body_paragraphs[:3]
            _inject_structured_text(info_shape, summary)
        else:
            _clear_shape_text(info_shape)

    return diag


# ============================================================================
# E. TABLE & IMAGE HANDLING
# ============================================================================

def _handle_tables(
    cloned_slide,
    content_data: ContentData,
    slide_w: int,
    slide_h: int,
) -> None:
    """Handle table transfer between content and template slides."""
    if not content_data.tables:
        return

    # Find existing tables in the cloned slide
    template_tables = [s for s in cloned_slide.shapes if _is_table(s)]
    content_table = content_data.tables[0]  # Use first content table

    if template_tables:
        # Case 1: Template has a table — fill it with content data
        tmpl_table_shape = template_tables[0]
        tmpl_table = tmpl_table_shape.table
        c_data = content_table["data"]

        if not c_data:
            return

        c_rows = len(c_data)
        c_cols = len(c_data[0]) if c_data else 0
        t_rows = len(tmpl_table.rows)
        t_cols = len(tmpl_table.columns)

        # Fill cells: min of available dimensions
        for ri in range(min(c_rows, t_rows)):
            for ci in range(min(c_cols, t_cols)):
                try:
                    cell = tmpl_table.cell(ri, ci)
                    cell.text = c_data[ri][ci]
                except Exception:
                    pass

        # Clear extra template cells
        for ri in range(c_rows, t_rows):
            for ci in range(t_cols):
                try:
                    tmpl_table.cell(ri, ci).text = ""
                except Exception:
                    pass
    else:
        # Case 2: No table in template — add the content table element
        try:
            spTree = cloned_slide.shapes._spTree
            table_el = content_table["element"]
            # Position in lower half of slide
            ns_a = _NSMAP["a"]
            # Just append the element
            spTree.append(deepcopy(table_el))
        except Exception:
            pass


def _handle_images(
    cloned_slide,
    content_data: ContentData,
    slide_w: int,
    slide_h: int,
    dst_prs: Presentation,
) -> None:
    """Add content images to the cloned slide if space is available."""
    if not content_data.images:
        return

    # Find occupied areas on the cloned slide
    occupied = []
    for shape in cloned_slide.shapes:
        occupied.append((
            shape.left or 0,
            shape.top or 0,
            (shape.left or 0) + (shape.width or 0),
            (shape.top or 0) + (shape.height or 0),
        ))

    for blob, orig_w, orig_h, orig_left, orig_top in content_data.images:
        # Try to find space in the lower-right area
        # Simple heuristic: place below existing content
        max_bottom = max((o[3] for o in occupied), default=int(slide_h * 0.3))
        available_top = min(max_bottom + Emu(Pt(10).emu), int(slide_h * 0.85))
        available_height = slide_h - available_top

        if available_height < int(slide_h * 0.1):
            # Not enough space
            continue

        # Scale image to fit available space
        target_w = min(orig_w, int(slide_w * 0.6))
        target_h = min(orig_h, available_height)
        # Maintain aspect ratio
        if orig_w > 0 and orig_h > 0:
            scale = min(target_w / orig_w, target_h / orig_h)
            target_w = int(orig_w * scale)
            target_h = int(orig_h * scale)

        # Center horizontally
        target_left = (slide_w - target_w) // 2

        try:
            from io import BytesIO
            img_stream = BytesIO(blob)
            cloned_slide.shapes.add_picture(
                img_stream, target_left, available_top, target_w, target_h,
            )
        except Exception:
            pass


# ============================================================================
# F. POST-PROCESSING
# ============================================================================

def _post_process(output_prs: Presentation) -> None:
    """Run post-processing on all slides in the output presentation."""
    for slide_idx, slide in enumerate(output_prs.slides):
        slide_num = slide_idx + 1

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text

            # Update page numbers
            if _PAGE_NUM_PATTERN.search(text):
                new_text = _PAGE_NUM_PATTERN.sub(f"Page {slide_num:02d}", text)
                if new_text != text:
                    _inject_text_into_shape(shape, new_text)

            # Update dates in footer areas
            sw = output_prs.slide_width or 1
            sh = output_prs.slide_height or 1
            bottom_frac = _shape_bottom(shape, sh)
            if bottom_frac >= 0.90:
                match = _DATE_PATTERN.search(text)
                if match:
                    today_str = date.today().strftime("%Y-%m-%d")
                    new_text = _DATE_PATTERN.sub(today_str, text)
                    if new_text != text:
                        _inject_text_into_shape(shape, new_text)


# ============================================================================
# G. VERBOSE DIAGNOSTICS
# ============================================================================

def _print_slide_diagnostic(
    slide_idx: int,
    total: int,
    content_data: ContentData,
    template_idx: int,
    template_struct: str,
    match_score: float,
    injection_diag: dict,
) -> None:
    """Print detailed diagnostic for one slide."""
    print(f"\nSlide {slide_idx+1}/{total}:")
    print(f"  Content type: {content_data.slide_type} "
          f"({content_data.word_count} words, "
          f"{len(content_data.tables)} table(s), "
          f"{len(content_data.images)} image(s))")
    print(f"  Template match: slide {template_idx+1} "
          f"(score={match_score:.0f}, type={template_struct})")

    if injection_diag.get("shapes"):
        print("  Shape classifications:")
        for s in injection_diag["shapes"]:
            preview = s["text_preview"]
            if preview:
                preview = f' "{preview}"'
            else:
                preview = ""
            print(f'    Shape "{s["name"]}" ({s["area_pct"]}% area, '
                  f'top {s["top_pct"]:.0f}%){preview} -> {s["role"]}')

    if injection_diag.get("injected_title"):
        print(f'  Injected: title="{injection_diag["injected_title"]}"')
    if injection_diag.get("injected_body"):
        print(f'  Injected: body ({injection_diag["injected_body"]})')
    print(f'  Protected: {injection_diag.get("protected_count", 0)} shapes untouched')


# ============================================================================
# DESIGN MODE ORCHESTRATOR
# ============================================================================

def apply_design(
    template_path: Path,
    content_path: Path,
    output_path: Path,
    slide_map_path: Path | None = None,
) -> None:
    """Design mode: clone template slides as skeletons, inject content."""
    print("\n[design] Loading presentations...")
    template_prs = Presentation(str(template_path))
    content_prs = Presentation(str(content_path))

    sw = template_prs.slide_width
    sh = template_prs.slide_height
    print(f"  Template: {len(template_prs.slides)} slides, "
          f"{Emu(sw).inches:.1f}\" x {Emu(sh).inches:.1f}\"")
    print(f"  Content:  {len(content_prs.slides)} slides")

    # Step 1: Extract content from all slides
    print("\n[design] Extracting content structure...")
    content_data_list = []
    for i, slide in enumerate(content_prs.slides):
        cd = extract_content(
            slide, i, len(content_prs.slides),
            content_prs.slide_width, content_prs.slide_height,
        )
        content_data_list.append(cd)
        _log(f"  Slide {i+1}: type={cd.slide_type}, words={cd.word_count}, "
             f"title=\"{cd.title[:40]}\", "
             f"paras={len(cd.body_paragraphs)}, "
             f"tables={len(cd.tables)}, images={len(cd.images)}")

    # Step 2: Build slide mapping
    print("\n[design] Mapping content slides to template slides...")
    if slide_map_path and slide_map_path.exists():
        raw_map = json.loads(slide_map_path.read_text())
        mapping = [int(raw_map.get(str(i+1), 1)) - 1 for i in range(len(content_data_list))]
        print("  Using manual slide mapping from", slide_map_path)
    else:
        mapping = build_slide_mapping(content_prs, template_prs, content_data_list)

    # Pre-compute template structure types for diagnostics
    t_structs = []
    for i, slide in enumerate(template_prs.slides):
        t_structs.append(_classify_template_structure(slide, sw, sh, i, len(template_prs.slides)))

    # Step 3: Create output from template
    print("\n[design] Building output presentation...")
    output_prs = Presentation(str(template_path))

    # Delete all existing slides
    prs_element = output_prs.slides._sldIdLst
    for sldId in list(prs_element):
        rId = sldId.get(f'{{{_NSMAP["r"]}}}id')
        if rId:
            try:
                output_prs.part.drop_rel(rId)
            except Exception:
                pass
        prs_element.remove(sldId)

    # Step 4: Clone and inject
    print("\n[design] Cloning and injecting content...")
    for ci, cd in enumerate(content_data_list):
        ti = mapping[ci]
        src_slide = template_prs.slides[ti]

        # Clone template slide
        new_slide = _clone_slide(template_prs, src_slide, output_prs)

        # Inject content
        diag = inject_content(new_slide, cd, sw, sh)

        # Handle tables
        _handle_tables(new_slide, cd, sw, sh)

        # Handle images
        _handle_images(new_slide, cd, sw, sh, output_prs)

        # Print status
        title_preview = cd.title[:50] if cd.title else "(no title)"
        if VERBOSE:
            # Compute score for diagnostic
            tinfo = {
                "struct": t_structs[ti],
                "words": sum(_word_count(_text_of(s)) for s in src_slide.shapes),
                "has_table": any(_is_table(s) for s in src_slide.shapes),
                "is_list": t_structs[ti] in ("list", "grid"),
            }
            score = _match_score_v2(
                cd.slide_type, tinfo["struct"],
                ci, ti, len(content_data_list), len(template_prs.slides),
                cd.word_count, tinfo["words"],
                len(cd.tables) > 0, tinfo["has_table"],
                len(cd.body_paragraphs), tinfo["is_list"],
            )
            _print_slide_diagnostic(
                ci, len(content_data_list), cd,
                ti, t_structs[ti], score, diag,
            )
        else:
            print(f"  Slide {ci+1}/{len(content_data_list)}: "
                  f'[{cd.slide_type}] "{title_preview}" '
                  f"<- template slide {ti+1} ({t_structs[ti]})")

    # Step 5: Post-processing
    print("\n[design] Post-processing...")
    _post_process(output_prs)

    # Step 6: Save
    print(f"\n[design] Saving to {output_path}...")
    output_prs.save(str(output_path))
    print(f"[design] Done! {len(content_data_list)} slides created.")


# ============================================================================
# LAYOUT MODE (backward-compatible)
# ============================================================================

def apply_layout(
    template_path: Path,
    content_path: Path,
    output_path: Path,
    layout_map_path: Path | None = None,
) -> None:
    """Layout mode: transfer theme + masters + layouts between files."""
    print("\n[layout] Loading presentations...")
    template_prs = Presentation(str(template_path))
    content_prs = Presentation(str(content_path))

    print(f"  Template layouts: {[l.name for l in template_prs.slide_layouts]}")
    print(f"  Content slides:   {len(content_prs.slides)}")
    print("  [layout] Falling back to design-mode pipeline (python-pptx limitation)")
    apply_design(template_path, content_path, output_path)


# ============================================================================
# AUTO-DETECTION
# ============================================================================

def detect_mode(template_path: Path) -> str:
    """Auto-detect whether to use 'design' or 'layout' mode."""
    prs = Presentation(str(template_path))
    generic_names = {"default", "blank", "empty", "custom", "custom layout", ""}

    has_named_layouts = False
    has_placeholders = False

    for layout in prs.slide_layouts:
        name = layout.name.strip().lower()
        if name not in generic_names:
            has_named_layouts = True
        if len(layout.placeholders) > 0:
            has_placeholders = True

    if has_named_layouts and has_placeholders:
        return "layout"
    return "design"


# ============================================================================
# CLI
# ============================================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="PPTX Template Transfer - apply one deck's visual design to another's content."
    )
    parser.add_argument("template_pptx", type=Path,
                        help="Template PPTX (design/theme source)")
    parser.add_argument("content_pptx", type=Path,
                        help="Content PPTX (text/data to transfer)")
    parser.add_argument("output_pptx", type=Path,
                        help="Output PPTX path")
    parser.add_argument("--mode", choices=["design", "layout"], default=None,
                        help="Transfer mode (auto-detected if not specified)")
    parser.add_argument("--slide-map", type=Path, default=None,
                        help='Manual slide mapping JSON: {"1": 3, "2": 1, ...}')
    parser.add_argument("--layout-map", type=Path, default=None,
                        help="Manual layout mapping JSON (backward compat)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Print detailed shape classification diagnostics")

    args = parser.parse_args()

    global VERBOSE
    VERBOSE = args.verbose

    # Ensure stdout handles unicode
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace",
        )

    for f, name in [(args.template_pptx, "Template"), (args.content_pptx, "Content")]:
        if not f.exists():
            print(f"Error: {name} not found: {f}", file=sys.stderr)
            sys.exit(1)

    mode = args.mode
    if mode is None:
        mode = detect_mode(args.template_pptx)
        print(f"Auto-detected mode: {mode}")

    if mode == "design":
        apply_design(args.template_pptx, args.content_pptx, args.output_pptx,
                     args.slide_map)
    else:
        apply_layout(args.template_pptx, args.content_pptx, args.output_pptx,
                     args.layout_map)


if __name__ == "__main__":
    main()
