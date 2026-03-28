#!/usr/bin/env python3
"""Scan an unpacked PPTX for relationship ID collisions and fix them.

Usage:
    python reconcile_rids.py <unpacked_dir>

Finds duplicate rId values within the same .rels scope and renumbers them,
updating all references in corresponding XML files.
"""

import argparse
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


def _write_xml(doc, path):
    """Write XML document with correct OOXML declaration."""
    xml_str = doc.toxml()
    xml_str = xml_str.replace(
        '<?xml version="1.0" ?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)


def _find_corresponding_xml(rels_path: Path, unpacked_dir: Path):
    """Find the XML file that a .rels file belongs to.

    E.g. ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
         ppt/_rels/presentation.xml.rels -> ppt/presentation.xml
    """
    # The rels file is in _rels/ subdir, named <filename>.rels
    rels_name = rels_path.name  # e.g. slide1.xml.rels
    if rels_name.endswith(".rels"):
        xml_name = rels_name[:-5]  # e.g. slide1.xml
    else:
        return None

    parent_dir = rels_path.parent.parent  # go up from _rels/
    xml_path = parent_dir / xml_name
    if xml_path.exists():
        return xml_path
    return None


def _collect_rid_references(doc):
    """Collect all r:id attribute values from an XML document."""
    refs = []
    for elem in doc.getElementsByTagName("*"):
        for i in range(elem.attributes.length):
            attr = elem.attributes.item(i)
            if attr.name == "r:id" or (attr.name.endswith(":id") and attr.value.startswith("rId")):
                refs.append((elem, attr.name, attr.value))
            # Also check r:embed, r:link, and other relationship attributes
            elif attr.name in ("r:embed", "r:link", "r:href", "r:dm", "r:cs", "r:lo", "r:qs"):
                if attr.value.startswith("rId"):
                    refs.append((elem, attr.name, attr.value))
    return refs


def reconcile_rels_file(rels_path: Path, unpacked_dir: Path):
    """Check a single .rels file for duplicate rIds and fix them.

    Returns number of fixes applied.
    """
    try:
        rels_doc = parse_xml(str(rels_path))
    except Exception as e:
        print(f"  Warning: Could not parse {rels_path}: {e}")
        return 0

    relationships = _find_all(rels_doc, "Relationship")
    if not relationships:
        return 0

    # Check for duplicates
    seen_ids = {}
    duplicates = []
    for rel in relationships:
        rid = _attr(rel, "Id", "")
        if rid in seen_ids:
            duplicates.append(rel)
        else:
            seen_ids[rid] = rel

    if not duplicates:
        return 0

    # Find the corresponding XML file to update references
    xml_path = _find_corresponding_xml(rels_path, unpacked_dir)

    # Compute max existing rId number
    max_rid = 0
    for rel in relationships:
        rid = _attr(rel, "Id", "")
        m = re.search(r"(\d+)", rid)
        if m:
            num = int(m.group(1))
            if num > max_rid:
                max_rid = num

    # Renumber duplicates
    renames = {}  # old_id -> new_id
    next_num = max_rid + 1

    for rel in duplicates:
        old_id = _attr(rel, "Id", "")
        new_id = f"rId{next_num}"
        next_num += 1

        rel.setAttribute("Id", new_id)
        renames[old_id] = new_id
        print(f"  {rels_path.name}: {old_id} -> {new_id} (target: {_attr(rel, 'Target', '')})")

    _write_xml(rels_doc, rels_path)

    # Update references in the corresponding XML
    if xml_path and renames:
        try:
            xml_doc = parse_xml(str(xml_path))
            refs = _collect_rid_references(xml_doc)
            updated = False

            for elem, attr_name, old_val in refs:
                if old_val in renames:
                    elem.setAttribute(attr_name, renames[old_val])
                    updated = True

            if updated:
                _write_xml(xml_doc, xml_path)
                print(f"  Updated references in {xml_path.name}")
        except Exception as e:
            print(f"  Warning: Could not update {xml_path}: {e}")

    return len(duplicates)


def reconcile_rids(unpacked_dir: Path):
    """Scan all .rels files and fix duplicate rIds."""
    total_fixes = 0

    # Find all .rels files
    rels_files = list(unpacked_dir.rglob("*.rels"))
    print(f"Scanning {len(rels_files)} .rels files...")

    for rels_path in sorted(rels_files):
        fixes = reconcile_rels_file(rels_path, unpacked_dir)
        total_fixes += fixes

    if total_fixes == 0:
        print("No relationship ID collisions found.")
    else:
        print(f"\nFixed {total_fixes} relationship ID collisions.")

    return total_fixes


def main():
    parser = argparse.ArgumentParser(
        description="Scan and fix relationship ID collisions in an unpacked PPTX."
    )
    parser.add_argument("unpacked_dir", type=Path,
                        help="Path to unpacked PPTX directory")
    args = parser.parse_args()

    if not args.unpacked_dir.exists():
        print(f"Error: Directory not found: {args.unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    reconcile_rids(args.unpacked_dir)


if __name__ == "__main__":
    main()
