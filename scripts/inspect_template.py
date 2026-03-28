#!/usr/bin/env python3
"""Analyze and report the full template structure of an unpacked PPTX.

Usage:
    python inspect_template.py <unpacked_dir>

Reports theme colors, fonts, slide masters, layouts, placeholder types,
background definitions, and media references.
"""

import argparse
import json
import sys
from pathlib import Path

from defusedxml.minidom import parse as parse_xml, parseString as parse_xml_string


# OOXML namespace prefixes used throughout
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _text(node):
    """Get text content of a node."""
    if node is None:
        return ""
    return node.firstChild.nodeValue if node.firstChild else ""


def _attr(node, attr, default=""):
    """Get attribute value or default."""
    if node is None:
        return default
    return node.getAttribute(attr) if node.hasAttribute(attr) else default


def _find_all(parent, tag_local):
    """Find all descendant elements matching a local tag name (ignoring namespace prefix)."""
    results = []
    if parent is None:
        return results
    for child in parent.getElementsByTagName("*"):
        local = child.localName or child.tagName.split(":")[-1]
        if local == tag_local:
            results.append(child)
    return results


def _find_first(parent, tag_local):
    """Find first descendant matching local tag name."""
    matches = _find_all(parent, tag_local)
    return matches[0] if matches else None


def _parse_color_element(elem):
    """Extract hex color from a color element (srgbClr, sysClr, etc.)."""
    if elem is None:
        return None
    for child in elem.childNodes:
        if getattr(child, "nodeType", None) != 1:
            continue
        local = child.localName or child.tagName.split(":")[-1]
        if local == "srgbClr":
            return "#" + _attr(child, "val", "000000")
        if local == "sysClr":
            last_clr = _attr(child, "lastClr", "")
            return f"sys:{_attr(child, 'val', '')}(#{last_clr})" if last_clr else f"sys:{_attr(child, 'val', '')}"
    return None


def inspect_theme(unpacked_dir: Path):
    """Inspect theme files and return theme info dict."""
    theme_dir = unpacked_dir / "ppt" / "theme"
    themes = []
    if not theme_dir.exists():
        return themes

    for theme_file in sorted(theme_dir.glob("theme*.xml")):
        try:
            doc = parse_xml(str(theme_file))
        except Exception as e:
            themes.append({"file": theme_file.name, "error": str(e)})
            continue

        info = {"file": theme_file.name, "colors": {}, "fonts": {}, "name": ""}

        # Theme name
        theme_elem = _find_first(doc, "theme")
        if theme_elem:
            info["name"] = _attr(theme_elem, "name", "")

        # Color scheme
        clr_scheme = _find_first(doc, "clrScheme")
        if clr_scheme:
            info["color_scheme_name"] = _attr(clr_scheme, "name", "")
            color_names = ["dk1", "lt1", "dk2", "lt2",
                           "accent1", "accent2", "accent3", "accent4",
                           "accent5", "accent6", "hlink", "folHlink"]
            for cname in color_names:
                elem = _find_first(clr_scheme, cname)
                info["colors"][cname] = _parse_color_element(elem)

        # Font scheme
        font_scheme = _find_first(doc, "fontScheme")
        if font_scheme:
            info["font_scheme_name"] = _attr(font_scheme, "name", "")
            major = _find_first(font_scheme, "majorFont")
            minor = _find_first(font_scheme, "minorFont")
            if major:
                latin = _find_first(major, "latin")
                info["fonts"]["major"] = _attr(latin, "typeface", "") if latin else ""
            if minor:
                latin = _find_first(minor, "latin")
                info["fonts"]["minor"] = _attr(latin, "typeface", "") if latin else ""

        themes.append(info)

    return themes


def inspect_rels(rels_path: Path):
    """Parse a .rels file and return list of {id, type, target}."""
    if not rels_path.exists():
        return []
    try:
        doc = parse_xml(str(rels_path))
    except Exception:
        return []
    rels = []
    for rel in _find_all(doc, "Relationship"):
        rels.append({
            "id": _attr(rel, "Id"),
            "type": _attr(rel, "Type").split("/")[-1],
            "target": _attr(rel, "Target"),
        })
    return rels


def inspect_masters_and_layouts(unpacked_dir: Path):
    """Inspect slide masters and their child layouts."""
    masters = []
    masters_dir = unpacked_dir / "ppt" / "slideMasters"
    layouts_dir = unpacked_dir / "ppt" / "slideLayouts"

    if not masters_dir.exists():
        return masters

    for master_file in sorted(masters_dir.glob("slideMaster*.xml")):
        master_info = {
            "file": master_file.name,
            "layouts": [],
            "background": None,
            "media": [],
        }

        # Parse master rels to find child layouts and media
        rels_path = masters_dir / "_rels" / f"{master_file.name}.rels"
        rels = inspect_rels(rels_path)
        layout_files = []
        for r in rels:
            if r["type"] == "slideLayout":
                layout_files.append(r["target"].split("/")[-1])
            elif r["type"] in ("image", "oleObject"):
                master_info["media"].append(r["target"])

        # Parse master XML for background
        try:
            doc = parse_xml(str(master_file))
            bg = _find_first(doc, "bg")
            if bg:
                master_info["background"] = "defined"
        except Exception:
            pass

        # Inspect each child layout
        for lf in sorted(layout_files):
            layout_path = layouts_dir / lf
            layout_info = {"file": lf, "name": "", "placeholders": [], "used_by_slides": []}

            if layout_path.exists():
                try:
                    ldoc = parse_xml(str(layout_path))
                    # Layout name from cSld
                    csld = _find_first(ldoc, "cSld")
                    if csld:
                        layout_info["name"] = _attr(csld, "name", "")

                    # Also check the layout type attribute
                    for elem in _find_all(ldoc, "cSld"):
                        layout_info["name"] = _attr(elem, "name", layout_info["name"])

                    # Placeholders
                    for ph in _find_all(ldoc, "ph"):
                        ph_type = _attr(ph, "type", "body")
                        ph_idx = _attr(ph, "idx", "")
                        layout_info["placeholders"].append({
                            "type": ph_type,
                            "idx": ph_idx,
                        })
                except Exception:
                    pass

            master_info["layouts"].append(layout_info)

        masters.append(master_info)

    return masters


def inspect_slides(unpacked_dir: Path):
    """Return list of slides with their layout references."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    slides = []
    if not slides_dir.exists():
        return slides

    for slide_file in sorted(slides_dir.glob("slide*.xml")):
        if slide_file.name.startswith("slide") and slide_file.suffix == ".xml":
            info = {"file": slide_file.name, "layout": None}
            rels_path = slides_dir / "_rels" / f"{slide_file.name}.rels"
            for r in inspect_rels(rels_path):
                if r["type"] == "slideLayout":
                    info["layout"] = r["target"].split("/")[-1]
                    break
            slides.append(info)

    return slides


def map_slides_to_layouts(slides, masters):
    """Add used_by_slides info to layouts in masters."""
    slide_layout_map = {}
    for s in slides:
        if s["layout"]:
            slide_layout_map.setdefault(s["layout"], []).append(s["file"])

    for master in masters:
        for layout in master["layouts"]:
            layout["used_by_slides"] = slide_layout_map.get(layout["file"], [])


def inspect(unpacked_dir: Path):
    """Run full template inspection and return results dict."""
    results = {
        "unpacked_dir": str(unpacked_dir),
        "themes": inspect_theme(unpacked_dir),
        "masters": inspect_masters_and_layouts(unpacked_dir),
        "slides": inspect_slides(unpacked_dir),
    }
    map_slides_to_layouts(results["slides"], results["masters"])
    return results


def print_report(results):
    """Print a human-readable report."""
    print("=" * 70)
    print(f"PPTX Template Inspection: {results['unpacked_dir']}")
    print("=" * 70)

    # Themes
    for theme in results["themes"]:
        print(f"\nTheme: {theme.get('name', 'N/A')} ({theme['file']})")
        if "error" in theme:
            print(f"  ERROR: {theme['error']}")
            continue

        print(f"  Color Scheme: {theme.get('color_scheme_name', 'N/A')}")
        for cname, cval in theme.get("colors", {}).items():
            print(f"    {cname:12s} = {cval}")

        print(f"  Font Scheme: {theme.get('font_scheme_name', 'N/A')}")
        for ftype, fval in theme.get("fonts", {}).items():
            print(f"    {ftype:12s} = {fval}")

    # Masters and layouts
    print(f"\nSlide Masters: {len(results['masters'])}")
    for master in results["masters"]:
        print(f"\n  {master['file']}")
        if master["background"]:
            print(f"    Background: {master['background']}")
        if master["media"]:
            print(f"    Media: {', '.join(master['media'])}")

        print(f"    Layouts ({len(master['layouts'])}):")
        for layout in master["layouts"]:
            ph_types = [p["type"] for p in layout["placeholders"]]
            used = layout.get("used_by_slides", [])
            print(f"      {layout['file']:30s} name=\"{layout['name']}\"")
            print(f"        Placeholders: {', '.join(ph_types) if ph_types else 'none'}")
            if used:
                print(f"        Used by: {', '.join(used)}")

    # Slides
    print(f"\nSlides: {len(results['slides'])}")
    for slide in results["slides"]:
        print(f"  {slide['file']:20s} -> {slide['layout']}")

    print()


def main():
    parser = argparse.ArgumentParser(
        description="Analyze the template structure of an unpacked PPTX directory."
    )
    parser.add_argument("unpacked_dir", type=Path, help="Path to unpacked PPTX directory")
    parser.add_argument("--json", action="store_true", help="Output as JSON instead of human-readable")
    args = parser.parse_args()

    if not args.unpacked_dir.exists():
        print(f"Error: Directory not found: {args.unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    results = inspect(args.unpacked_dir)

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        print_report(results)


if __name__ == "__main__":
    main()
