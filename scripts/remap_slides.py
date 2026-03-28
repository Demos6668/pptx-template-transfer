#!/usr/bin/env python3
"""Remap all slides in a target PPTX to use new layouts based on a mapping file.

Usage:
    python remap_slides.py <target_unpacked_dir> <mapping.json>

Updates each slide's .rels to point to the new layout while preserving content.
"""

import argparse
import json
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


def _get_placeholder_types(layout_path: Path):
    """Get set of placeholder types from a layout file."""
    if not layout_path.exists():
        return set()
    try:
        doc = parse_xml(str(layout_path))
        return {_attr(ph, "type", "body") for ph in _find_all(doc, "ph")}
    except Exception:
        return set()


def _get_slide_placeholder_types(slide_path: Path):
    """Get set of placeholder types used by a slide."""
    if not slide_path.exists():
        return set()
    try:
        doc = parse_xml(str(slide_path))
        return {_attr(ph, "type", "body") for ph in _find_all(doc, "ph")}
    except Exception:
        return set()


def remap_slides(target_dir: Path, mapping_path: Path):
    """Remap slides to new layouts based on mapping file."""
    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping_data = json.load(f)

    mappings = mapping_data.get("mappings", [])
    if not mappings:
        print("No mappings found in mapping file.")
        return

    slides_dir = target_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    layouts_dir = target_dir / "ppt" / "slideLayouts"

    if not rels_dir.exists():
        print("Error: No slides/_rels/ directory found.", file=sys.stderr)
        sys.exit(1)

    remapped = 0
    warnings = []

    for entry in mappings:
        slide_file = entry["target_slide"]
        new_layout = entry.get("suggested_layout")

        if not new_layout:
            print(f"  Skipping {slide_file}: no suggested layout")
            continue

        rels_path = rels_dir / f"{slide_file}.rels"
        if not rels_path.exists():
            print(f"  Warning: No rels file for {slide_file}")
            continue

        # Check placeholder compatibility
        slide_path = slides_dir / slide_file
        new_layout_path = layouts_dir / new_layout
        slide_ph = _get_slide_placeholder_types(slide_path)
        layout_ph = _get_placeholder_types(new_layout_path)

        if slide_ph and layout_ph:
            missing = slide_ph - layout_ph
            if missing:
                msg = (f"  Warning: {slide_file} uses placeholder types "
                       f"{missing} not in {new_layout}")
                warnings.append(msg)
                print(msg)

        # Update rels file
        try:
            rdoc = parse_xml(str(rels_path))
            updated = False

            for rel in _find_all(rdoc, "Relationship"):
                rel_type = _attr(rel, "Type", "")
                if "slideLayout" in rel_type:
                    old_target = _attr(rel, "Target", "")
                    new_target = f"../slideLayouts/{new_layout}"
                    if old_target != new_target:
                        rel.setAttribute("Target", new_target)
                        updated = True
                        print(f"  {slide_file}: {old_target} -> {new_target}")
                    break

            if updated:
                _write_xml(rdoc, rels_path)
                remapped += 1
            else:
                print(f"  {slide_file}: already using correct layout")

        except Exception as e:
            print(f"  Error processing {slide_file}: {e}", file=sys.stderr)

    print(f"\nRemapped {remapped}/{len(mappings)} slides")
    if warnings:
        print(f"Warnings: {len(warnings)} placeholder mismatches")


def main():
    parser = argparse.ArgumentParser(
        description="Remap slides to new layouts based on a mapping file."
    )
    parser.add_argument("target_unpacked_dir", type=Path,
                        help="Path to target unpacked PPTX directory")
    parser.add_argument("mapping_json", type=Path,
                        help="Path to layout mapping JSON file")
    args = parser.parse_args()

    if not args.target_unpacked_dir.exists():
        print(f"Error: Target not found: {args.target_unpacked_dir}", file=sys.stderr)
        sys.exit(1)
    if not args.mapping_json.exists():
        print(f"Error: Mapping file not found: {args.mapping_json}", file=sys.stderr)
        sys.exit(1)

    print(f"Remapping slides in {args.target_unpacked_dir}...")
    remap_slides(args.target_unpacked_dir, args.mapping_json)


if __name__ == "__main__":
    main()
