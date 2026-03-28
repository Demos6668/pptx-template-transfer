#!/usr/bin/env python3
"""Adapt text colors to avoid dark-on-dark or light-on-light after template transfer.

Usage:
    python adapt_text_colors.py <unpacked_dir> [--dry-run]

For each slide, determines the effective background color (from slide, layout, or
master), then checks all hardcoded text colors for contrast. Flips dark text to
white on dark backgrounds and light text to dark on light backgrounds.

Only modifies explicit srgbClr values. Theme-referenced colors (schemeClr) are
left alone since they adapt automatically through the theme.
"""

import argparse
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


def _write_xml(doc, path):
    """Write XML document with correct OOXML declaration."""
    xml_str = doc.toxml()
    xml_str = xml_str.replace(
        '<?xml version="1.0" ?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)


def _hex_to_rgb(hex_color: str) -> tuple:
    """Convert 6-char hex string to (R, G, B) tuple (0-255)."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return (128, 128, 128)  # neutral gray fallback
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    except ValueError:
        return (128, 128, 128)


def _luminance(r: int, g: int, b: int) -> float:
    """Compute perceptual luminance (0.0 = black, 1.0 = white)."""
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def _is_dark(hex_color: str) -> bool:
    """Return True if color is dark (luminance < 0.5)."""
    r, g, b = _hex_to_rgb(hex_color)
    return _luminance(r, g, b) < 0.5


def _get_solid_fill_color(parent_elem):
    """Extract hex color from a solidFill element under parent.

    Looks for <a:solidFill><a:srgbClr val="XXXXXX"/> patterns.
    Returns hex string (no #) or None.
    """
    for sf in _find_all(parent_elem, "solidFill"):
        for child in sf.childNodes:
            if getattr(child, "nodeType", None) != 1:
                continue
            local = child.localName or child.tagName.split(":")[-1]
            if local == "srgbClr":
                return _attr(child, "val", "")
        # Skip schemeClr, sysClr, etc. — they adapt via theme
    return None


def _get_background_color_from_xml(doc):
    """Try to extract a solid background color from a parsed XML document.

    Checks <p:bg> -> <p:bgPr> -> <a:solidFill> -> <a:srgbClr>
    Also checks <p:bg> -> <p:bgRef> with embedded solidFill.
    Returns hex string (no #) or None.
    """
    for bg in _find_all(doc, "bg"):
        # bgPr path
        for bgPr in _find_all(bg, "bgPr"):
            color = _get_solid_fill_color(bgPr)
            if color:
                return color

        # bgRef path (sometimes has embedded fill)
        for bgRef in _find_all(bg, "bgRef"):
            color = _get_solid_fill_color(bgRef)
            if color:
                return color

    return None


def get_effective_background(slide_path: Path, unpacked_dir: Path) -> str:
    """Determine effective background color for a slide.

    Checks slide -> layout -> master hierarchy.
    Returns hex color string (no #), defaults to "FFFFFF" (white).
    """
    slides_dir = unpacked_dir / "ppt" / "slides"
    layouts_dir = unpacked_dir / "ppt" / "slideLayouts"
    masters_dir = unpacked_dir / "ppt" / "slideMasters"

    # 1. Check slide itself
    try:
        slide_doc = parse_xml(str(slide_path))
        color = _get_background_color_from_xml(slide_doc)
        if color:
            return color
    except Exception:
        pass

    # 2. Find the slide's layout via .rels
    layout_file = None
    rels_path = slides_dir / "_rels" / f"{slide_path.name}.rels"
    if rels_path.exists():
        try:
            rdoc = parse_xml(str(rels_path))
            for rel in _find_all(rdoc, "Relationship"):
                if "slideLayout" in _attr(rel, "Type", ""):
                    layout_file = Path(_attr(rel, "Target", "")).name
                    break
        except Exception:
            pass

    # Check layout background
    if layout_file:
        layout_path = layouts_dir / layout_file
        if layout_path.exists():
            try:
                layout_doc = parse_xml(str(layout_path))
                color = _get_background_color_from_xml(layout_doc)
                if color:
                    return color
            except Exception:
                pass

            # 3. Find layout's parent master via .rels
            master_file = None
            layout_rels = layouts_dir / "_rels" / f"{layout_file}.rels"
            if layout_rels.exists():
                try:
                    lrdoc = parse_xml(str(layout_rels))
                    for rel in _find_all(lrdoc, "Relationship"):
                        if "slideMaster" in _attr(rel, "Type", ""):
                            master_file = Path(_attr(rel, "Target", "")).name
                            break
                except Exception:
                    pass

            if master_file:
                master_path = masters_dir / master_file
                if master_path.exists():
                    try:
                        master_doc = parse_xml(str(master_path))
                        color = _get_background_color_from_xml(master_doc)
                        if color:
                            return color
                    except Exception:
                        pass

    # Default: assume white
    return "FFFFFF"


def adapt_slide_colors(slide_path: Path, bg_color: str, dry_run: bool = False) -> list:
    """Check and fix text colors in a single slide for contrast against background.

    Returns list of changes made (or would-be-made in dry_run mode).
    """
    changes = []
    bg_dark = _is_dark(bg_color)

    try:
        doc = parse_xml(str(slide_path))
    except Exception as e:
        print(f"  Warning: Could not parse {slide_path.name}: {e}")
        return changes

    modified = False

    # Find all text run properties with explicit colors
    for rPr in _find_all(doc, "rPr"):
        for sf in _find_all(rPr, "solidFill"):
            for child in sf.childNodes:
                if getattr(child, "nodeType", None) != 1:
                    continue
                local = child.localName or child.tagName.split(":")[-1]
                if local != "srgbClr":
                    continue  # Skip schemeClr, sysClr — they adapt via theme

                old_color = _attr(child, "val", "")
                if not old_color or len(old_color) != 6:
                    continue

                text_dark = _is_dark(old_color)

                new_color = None
                if bg_dark and text_dark:
                    # Dark text on dark background → flip to white
                    new_color = "FFFFFF"
                elif not bg_dark and not text_dark:
                    # Light text on light background → flip to dark
                    new_color = "333333"

                if new_color and new_color.upper() != old_color.upper():
                    change = {
                        "slide": slide_path.name,
                        "old_color": f"#{old_color}",
                        "new_color": f"#{new_color}",
                        "reason": "dark_on_dark" if bg_dark else "light_on_light",
                        "background": f"#{bg_color}",
                    }
                    changes.append(change)

                    if not dry_run:
                        child.setAttribute("val", new_color)
                        modified = True

    if modified and not dry_run:
        _write_xml(doc, slide_path)

    return changes


def adapt_text_colors(unpacked_dir: Path, dry_run: bool = False) -> list:
    """Scan all slides and adapt text colors for contrast.

    Returns list of all changes made.
    """
    slides_dir = unpacked_dir / "ppt" / "slides"
    if not slides_dir.exists():
        print("  No slides directory found.")
        return []

    all_changes = []
    slide_files = sorted(slides_dir.glob("slide*.xml"))

    for slide_path in slide_files:
        if not slide_path.name.startswith("slide") or slide_path.suffix != ".xml":
            continue

        bg_color = get_effective_background(slide_path, unpacked_dir)
        bg_type = "dark" if _is_dark(bg_color) else "light"
        print(f"  {slide_path.name}: bg=#{bg_color} ({bg_type})")

        changes = adapt_slide_colors(slide_path, bg_color, dry_run)
        all_changes.extend(changes)

        for c in changes:
            action = "Would flip" if dry_run else "Flipped"
            print(f"    {action} {c['old_color']} -> {c['new_color']} ({c['reason']})")

    if not all_changes:
        print("  No color adaptations needed.")
    else:
        action = "would adapt" if dry_run else "adapted"
        print(f"\n  Total: {action} {len(all_changes)} text color(s)")

    return all_changes


def main():
    parser = argparse.ArgumentParser(
        description="Adapt text colors to avoid contrast issues after template transfer."
    )
    parser.add_argument("unpacked_dir", type=Path,
                        help="Path to unpacked PPTX directory")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print what would change without modifying files")
    args = parser.parse_args()

    if not args.unpacked_dir.exists():
        print(f"Error: Directory not found: {args.unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    mode = "DRY RUN" if args.dry_run else "LIVE"
    print(f"Adapting text colors ({mode})...")
    adapt_text_colors(args.unpacked_dir, args.dry_run)


if __name__ == "__main__":
    main()
