#!/usr/bin/env python3
"""Generate a layout mapping between source template and target slides.

Usage:
    python map_layouts.py <source_unpacked_dir> <target_unpacked_dir> [--output mapping.json]

Produces a JSON mapping file that maps each target slide to the best-matching
source layout based on name, synonym, slide content analysis, placeholder
signature, or fallback heuristics.
"""

import argparse
import json
import re
import sys
from pathlib import Path

from defusedxml.minidom import parse as parse_xml


def _find_all(parent, tag_local):
    results = []
    if parent is None:
        return results
    for child in parent.getElementsByTagName("*"):
        local = child.localName or child.tagName.split(":")[-1]
        if local == tag_local:
            results.append(child)
    return results


def _attr(node, attr, default=""):
    if node is None:
        return default
    return node.getAttribute(attr) if node.hasAttribute(attr) else default


# Common layout name synonyms for fuzzy matching
SYNONYMS = {
    "two content": {"two column", "two_content", "2 content", "2 column"},
    "two column": {"two content", "two_content", "2 content", "2 column"},
    "title only": {"title_only", "titleonly"},
    "title slide": {"title_slide", "titleslide"},
    "section header": {"section_header", "sectionheader", "section title"},
    "section title": {"section_header", "sectionheader", "section header"},
    "comparison": {"compare"},
    "content with caption": {"content_with_caption", "caption"},
    "picture with caption": {"picture_with_caption", "pic caption"},
    "blank": {"empty"},
}

# Layout names considered generic (trigger content-aware analysis)
GENERIC_LAYOUT_NAMES = {
    "default", "blank", "empty", "custom", "custom layout", "layout",
    "1_default", "2_default", "default design",
}


def _normalize_name(name):
    """Normalize layout name for comparison."""
    return name.strip().lower().replace("_", " ").replace("-", " ")


def _get_layout_info(layouts_dir: Path):
    """Get info for all layouts in a directory: {filename: {name, placeholders}}."""
    layouts = {}
    if not layouts_dir.exists():
        return layouts

    for lf in sorted(layouts_dir.glob("slideLayout*.xml")):
        info = {"file": lf.name, "name": "", "placeholders": [], "placeholder_types": set()}
        try:
            doc = parse_xml(str(lf))
            csld = _find_all(doc, "cSld")
            if csld:
                info["name"] = _attr(csld[0], "name", "")

            for ph in _find_all(doc, "ph"):
                ph_type = _attr(ph, "type", "body")
                info["placeholders"].append(ph_type)
                info["placeholder_types"].add(ph_type)
        except Exception:
            pass

        layouts[lf.name] = info
    return layouts


def _get_slides_with_layouts(slides_dir: Path):
    """Get list of slides with their current layout references."""
    slides = []
    if not slides_dir.exists():
        return slides

    for sf in sorted(slides_dir.glob("slide*.xml")):
        if not sf.name.startswith("slide") or sf.suffix != ".xml":
            continue

        info = {"file": sf.name, "layout": None, "layout_name": ""}
        rels_path = slides_dir / "_rels" / f"{sf.name}.rels"
        if rels_path.exists():
            try:
                rdoc = parse_xml(str(rels_path))
                for rel in _find_all(rdoc, "Relationship"):
                    if "slideLayout" in _attr(rel, "Type", ""):
                        info["layout"] = Path(_attr(rel, "Target", "")).name
                        break
            except Exception:
                pass
        slides.append(info)
    return slides


# ---------------------------------------------------------------------------
# Content-aware slide analysis
# ---------------------------------------------------------------------------

def _get_text_from_node(node):
    """Recursively extract all text from a node tree."""
    parts = []
    if node.nodeType == node.TEXT_NODE:
        return node.nodeValue or ""
    for child in node.childNodes:
        parts.append(_get_text_from_node(child))
    return "".join(parts)


def _analyze_slide_content(slide_path: Path) -> dict:
    """Analyze a slide's XML to infer its intended layout type.

    Returns dict with text_boxes, max_font_size, has_title_like, has_body_text,
    has_images, is_first_slide, inferred_type.
    """
    result = {
        "text_boxes": 0,
        "max_font_size": 0,          # hundredths of a point
        "has_title_like": False,
        "has_body_text": False,
        "has_images": False,
        "is_first_slide": slide_path.name == "slide1.xml",
        "inferred_type": "unknown",
    }

    try:
        doc = parse_xml(str(slide_path))
    except Exception:
        return result

    # Collect shape info
    shapes = []
    for sp in _find_all(doc, "sp"):
        shape_info = {
            "has_text": False,
            "text": "",
            "max_font_size": 0,
            "is_bold": False,
            "top": 0,
            "is_centered_h": False,
            "ph_type": None,
        }

        # Check for placeholder type
        for ph in _find_all(sp, "ph"):
            shape_info["ph_type"] = _attr(ph, "type", "body")

        # Position (top in EMU)
        for off in _find_all(sp, "off"):
            top = _attr(off, "y", "0")
            try:
                shape_info["top"] = int(top)
            except (ValueError, TypeError):
                pass

        # Extent (width) — check if horizontally centered-ish
        for ext_elem in _find_all(sp, "ext"):
            cx = _attr(ext_elem, "cx", "0")
            try:
                # > 60% of standard slide width (12192000 EMU) = roughly centered
                if int(cx) > 7_000_000:
                    shape_info["is_centered_h"] = True
            except (ValueError, TypeError):
                pass

        # Text content and font properties
        for txBody in _find_all(sp, "txBody"):
            full_text = ""
            for p_elem in _find_all(txBody, "p"):
                for r_elem in _find_all(p_elem, "r"):
                    t_elem = None
                    for child in r_elem.childNodes:
                        local = getattr(child, "localName", None) or ""
                        if local == "t":
                            t_elem = child
                            break
                    if t_elem:
                        full_text += _get_text_from_node(t_elem)

                    # Font size from run properties
                    for rPr in _find_all(r_elem, "rPr"):
                        sz = _attr(rPr, "sz", "")
                        if sz:
                            try:
                                fs = int(sz)
                                shape_info["max_font_size"] = max(shape_info["max_font_size"], fs)
                                result["max_font_size"] = max(result["max_font_size"], fs)
                            except (ValueError, TypeError):
                                pass
                        b_val = _attr(rPr, "b", "")
                        if b_val == "1" or b_val == "true":
                            shape_info["is_bold"] = True

                # Also check paragraph-level default run properties
                for defRPr in _find_all(p_elem, "defRPr"):
                    sz = _attr(defRPr, "sz", "")
                    if sz:
                        try:
                            fs = int(sz)
                            shape_info["max_font_size"] = max(shape_info["max_font_size"], fs)
                            result["max_font_size"] = max(result["max_font_size"], fs)
                        except (ValueError, TypeError):
                            pass

            shape_info["text"] = full_text.strip()
            if full_text.strip():
                shape_info["has_text"] = True

        if shape_info["has_text"]:
            shapes.append(shape_info)

    # Check for images (pic elements)
    for _pic in _find_all(doc, "pic"):
        result["has_images"] = True
        break

    result["text_boxes"] = len(shapes)

    # Classify shapes
    # A "title-like" element: large font (>= 2400 hundredths = 24pt), in top portion
    # Top portion = top < 2500000 EMU (~2.5 inches from top)
    title_shapes = []
    body_shapes = []
    for s in shapes:
        is_large = s["max_font_size"] >= 2400
        is_top = s["top"] < 2_500_000
        if is_large and is_top:
            title_shapes.append(s)
            result["has_title_like"] = True
        elif s["has_text"] and not is_large:
            body_shapes.append(s)
            result["has_body_text"] = True

    # Infer type
    num_text = result["text_boxes"]

    if num_text == 0:
        if result["has_images"]:
            result["inferred_type"] = "picture"
        else:
            result["inferred_type"] = "blank"
    elif result["is_first_slide"] and num_text <= 2 and result["has_title_like"]:
        result["inferred_type"] = "title"
    elif result["has_title_like"] and result["has_body_text"]:
        result["inferred_type"] = "content"
    elif num_text <= 2 and result["has_title_like"] and not result["has_body_text"]:
        # Large centered text, no body — section header
        centered = any(s["is_centered_h"] for s in title_shapes)
        if centered:
            result["inferred_type"] = "section"
        else:
            result["inferred_type"] = "content"
    elif result["has_title_like"]:
        result["inferred_type"] = "content"
    elif num_text == 1 and shapes[0]["max_font_size"] >= 2000:
        result["inferred_type"] = "section"
    else:
        result["inferred_type"] = "content"

    return result


# Layout type keywords for matching inferred types to layout names
_LAYOUT_TYPE_KEYWORDS = {
    "title": ["title slide", "cover", "title page", "branded_title", "branded title"],
    "content": ["content", "body", "text", "title and content", "title, content",
                 "branded_content", "branded content"],
    "section": ["section", "divider", "break", "header"],
    "picture": ["picture", "image", "photo", "media"],
    "blank": ["blank", "empty"],
}

# Layout names to EXCLUDE for certain types (avoid mapping "title" to "title and content")
_LAYOUT_TYPE_EXCLUDES = {
    "title": ["title and content", "title, content", "title and body",
              "title only", "title and text"],
}


def _find_best_layout_for_type(inferred_type: str, source_layouts: dict,
                               template_dir: Path = None) -> tuple:
    """Find the best source layout for an inferred slide type.

    Returns (layout_filename, match_type, confidence).
    """
    if inferred_type == "unknown":
        return None, "none", "none"

    keywords = _LAYOUT_TYPE_KEYWORDS.get(inferred_type, [])
    excludes = _LAYOUT_TYPE_EXCLUDES.get(inferred_type, [])

    # Search layout names for keyword matches
    best = None
    for fname, info in source_layouts.items():
        lname = _normalize_name(info["name"])
        if not lname:
            continue

        # Check excludes first
        excluded = False
        for ex in excludes:
            if ex in lname:
                excluded = True
                break
        if excluded:
            continue

        for kw in keywords:
            if kw in lname:
                best = fname
                return best, "content_analysis", "high"

    # Fallback chain: section → content, picture → content
    _FALLBACK_TYPES = {"section": "content", "picture": "content"}
    fallback_type = _FALLBACK_TYPES.get(inferred_type)
    if fallback_type:
        fb_keywords = _LAYOUT_TYPE_KEYWORDS.get(fallback_type, [])
        fb_excludes = _LAYOUT_TYPE_EXCLUDES.get(fallback_type, [])
        for fname, info in source_layouts.items():
            lname = _normalize_name(info["name"])
            if not lname:
                continue
            excluded = any(ex in lname for ex in fb_excludes)
            if excluded:
                continue
            for kw in fb_keywords:
                if kw in lname:
                    return fname, "content_analysis_type_fallback", "medium"

    # Positional fallback: for title type, use first layout used by template's slide1
    # For content type, use first layout used by template's non-first slides
    if template_dir is not None:
        template_slides_dir = template_dir / "ppt" / "slides"
        if template_slides_dir.exists():
            template_slides = _get_slides_with_layouts(template_slides_dir)
            if inferred_type == "title":
                for ts in template_slides:
                    if ts["file"] == "slide1.xml" and ts["layout"]:
                        return ts["layout"], "content_analysis_fallback", "medium"
            else:
                for ts in template_slides:
                    if ts["file"] != "slide1.xml" and ts["layout"]:
                        return ts["layout"], "content_analysis_fallback", "medium"

    # Placeholder-based fallback for content type
    if inferred_type == "content":
        for fname, info in source_layouts.items():
            ph = info.get("placeholder_types", set())
            if "body" in ph or "title" in ph:
                return fname, "content_analysis_placeholder", "medium"

    return None, "none", "none"


def _match_score(source_layout, target_layout_name, target_ph_types):
    """Compute match score between a source layout and a target layout description."""
    src_name = _normalize_name(source_layout["name"])
    tgt_name = _normalize_name(target_layout_name)

    # Exact name match
    if src_name == tgt_name:
        return 100, "exact_name", "high"

    # Synonym match
    src_synonyms = SYNONYMS.get(src_name, set())
    tgt_synonyms = SYNONYMS.get(tgt_name, set())
    if tgt_name in src_synonyms or src_name in tgt_synonyms:
        return 90, "synonym_name", "high"

    # Partial name match (one contains the other)
    if src_name and tgt_name:
        if src_name in tgt_name or tgt_name in src_name:
            return 70, "partial_name", "medium"

    # Placeholder signature match
    src_ph = source_layout.get("placeholder_types", set())
    if src_ph and target_ph_types:
        intersection = src_ph & target_ph_types
        union = src_ph | target_ph_types
        if union:
            jaccard = len(intersection) / len(union)
            if jaccard > 0.7:
                return int(60 * jaccard), "placeholder_match", "medium"
            elif jaccard > 0.3:
                return int(40 * jaccard), "placeholder_partial", "low"

    return 0, "none", "none"


def _is_generic_layout_name(name: str) -> bool:
    """Check if a layout name is generic/default (triggers content analysis)."""
    normalized = _normalize_name(name)
    if not normalized:
        return True
    return normalized in GENERIC_LAYOUT_NAMES


def map_layouts(source_dir: Path, target_dir: Path):
    """Generate layout mapping between source and target."""
    src_layouts = _get_layout_info(source_dir / "ppt" / "slideLayouts")
    tgt_layouts = _get_layout_info(target_dir / "ppt" / "slideLayouts")
    tgt_slides = _get_slides_with_layouts(target_dir / "ppt" / "slides")

    # Enrich slide info with layout names
    for slide in tgt_slides:
        if slide["layout"] and slide["layout"] in tgt_layouts:
            slide["layout_name"] = tgt_layouts[slide["layout"]]["name"]

    # Find blank/fallback layout in source
    blank_layout = None
    first_layout = None
    for fname, info in src_layouts.items():
        if first_layout is None:
            first_layout = fname
        if _normalize_name(info["name"]) in ("blank", "empty"):
            blank_layout = fname
            break

    fallback = blank_layout or first_layout

    mappings = []
    unmapped = []

    for slide in tgt_slides:
        current_layout = slide["layout"] or ""
        current_name = slide.get("layout_name", "")
        current_ph = tgt_layouts.get(current_layout, {}).get("placeholder_types", set())

        best_match = None
        best_score = -1
        best_match_type = "none"
        best_confidence = "none"

        # Step 1-2: Name-based matching (exact + synonym)
        for src_fname, src_info in src_layouts.items():
            score, match_type, confidence = _match_score(src_info, current_name, current_ph)
            if score > best_score:
                best_score = score
                best_match = src_fname
                best_match_type = match_type
                best_confidence = confidence

        # Step 3: Content-aware analysis — when current layout is generic or match is weak
        use_content_analysis = False
        if best_score <= 0:
            use_content_analysis = True
        elif _is_generic_layout_name(current_name):
            # Current layout is generic (e.g., DEFAULT) — always try content analysis
            use_content_analysis = True
        elif best_match_type in ("exact_name", "synonym_name"):
            # Check if the match was to a generic name (e.g., DEFAULT→DEFAULT)
            matched_name = _normalize_name(src_layouts.get(best_match, {}).get("name", ""))
            if _is_generic_layout_name(matched_name):
                use_content_analysis = True

        if use_content_analysis:
            slide_path = target_dir / "ppt" / "slides" / slide["file"]
            analysis = _analyze_slide_content(slide_path)
            inferred = analysis["inferred_type"]

            if inferred != "unknown":
                ca_match, ca_type, ca_conf = _find_best_layout_for_type(
                    inferred, src_layouts, source_dir
                )
                ca_score = {"high": 80, "medium": 60, "low": 40}.get(ca_conf, 0)
                if ca_match and ca_score >= best_score:
                    best_match = ca_match
                    best_match_type = ca_type
                    best_confidence = ca_conf
                    best_score = ca_score

        # Step 5: Final fallback
        if best_score <= 0 or best_match is None:
            # First slide → title-like layout, others → content-like layout
            slide_path = target_dir / "ppt" / "slides" / slide["file"]
            if slide["file"] == "slide1.xml":
                fb_match, fb_type, fb_conf = _find_best_layout_for_type(
                    "title", src_layouts, source_dir
                )
                if fb_match:
                    best_match = fb_match
                    best_match_type = "positional_fallback"
                    best_confidence = "low"
                else:
                    best_match = fallback
                    best_match_type = "fallback"
                    best_confidence = "low"
            else:
                fb_match, fb_type, fb_conf = _find_best_layout_for_type(
                    "content", src_layouts, source_dir
                )
                if fb_match:
                    best_match = fb_match
                    best_match_type = "positional_fallback"
                    best_confidence = "low"
                else:
                    best_match = fallback
                    best_match_type = "fallback"
                    best_confidence = "low"

        entry = {
            "target_slide": slide["file"],
            "current_layout": current_layout,
            "current_layout_name": current_name,
            "suggested_layout": best_match,
            "suggested_layout_name": src_layouts.get(best_match, {}).get("name", ""),
            "match_type": best_match_type,
            "confidence": best_confidence,
        }
        mappings.append(entry)

        if best_confidence == "none" or best_match is None:
            unmapped.append(slide["file"])

    result = {
        "mappings": mappings,
        "unmapped_slides": unmapped,
        "available_source_layouts": [
            {"file": fname, "name": info["name"],
             "placeholders": list(info.get("placeholder_types", set()))}
            for fname, info in sorted(src_layouts.items())
        ],
    }
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Generate layout mapping between source template and target slides."
    )
    parser.add_argument("source_unpacked_dir", type=Path,
                        help="Path to source (template) unpacked PPTX")
    parser.add_argument("target_unpacked_dir", type=Path,
                        help="Path to target (content) unpacked PPTX")
    parser.add_argument("--output", type=Path, default=None,
                        help="Output mapping JSON file (default: stdout)")
    args = parser.parse_args()

    if not args.source_unpacked_dir.exists():
        print(f"Error: Source not found: {args.source_unpacked_dir}", file=sys.stderr)
        sys.exit(1)
    if not args.target_unpacked_dir.exists():
        print(f"Error: Target not found: {args.target_unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    result = map_layouts(args.source_unpacked_dir, args.target_unpacked_dir)

    output_str = json.dumps(result, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output_str)
        print(f"Mapping written to {args.output}")
        # Summary
        high = sum(1 for m in result["mappings"] if m["confidence"] == "high")
        med = sum(1 for m in result["mappings"] if m["confidence"] == "medium")
        low = sum(1 for m in result["mappings"] if m["confidence"] == "low")
        print(f"  {len(result['mappings'])} slides mapped: {high} high, {med} medium, {low} low confidence")
        if result["unmapped_slides"]:
            print(f"  Unmapped: {', '.join(result['unmapped_slides'])}")
    else:
        print(output_str)


if __name__ == "__main__":
    main()
