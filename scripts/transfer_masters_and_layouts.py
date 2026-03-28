#!/usr/bin/env python3
"""Copy slide masters and their child layouts from source to target PPTX.

Usage:
    python transfer_masters_and_layouts.py <source_unpacked_dir> <target_unpacked_dir> [--replace | --merge]

Modes:
    --replace (default): Remove existing masters/layouts, replace with source.
    --merge: Add source masters/layouts alongside existing ones.
"""

import argparse
import re
import shutil
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


def _next_id(existing_ids):
    """Return next available numeric ID (for sldMasterIdLst, etc.)."""
    max_id = max(existing_ids) if existing_ids else 2147483647
    return max_id + 1


def _extract_number(filename):
    """Extract number from filename like slideMaster3.xml -> 3."""
    m = re.search(r"(\d+)", filename)
    return int(m.group(1)) if m else 0


def _clear_directory_contents(dir_path: Path, pattern: str):
    """Remove files matching pattern in directory."""
    if dir_path.exists():
        for f in dir_path.glob(pattern):
            f.unlink()
        rels_dir = dir_path / "_rels"
        if rels_dir.exists():
            for f in rels_dir.glob(pattern + ".rels"):
                f.unlink()


def _get_max_existing_number(dir_path: Path, prefix: str):
    """Get the highest number from files like slideLayout12.xml."""
    max_num = 0
    if dir_path.exists():
        for f in dir_path.glob(f"{prefix}*.xml"):
            num = _extract_number(f.name)
            if num > max_num:
                max_num = num
    return max_num


def transfer_masters_and_layouts(source_dir: Path, target_dir: Path, mode: str = "replace"):
    """Transfer slide masters and layouts from source to target."""
    src_masters_dir = source_dir / "ppt" / "slideMasters"
    src_layouts_dir = source_dir / "ppt" / "slideLayouts"
    tgt_masters_dir = target_dir / "ppt" / "slideMasters"
    tgt_layouts_dir = target_dir / "ppt" / "slideLayouts"

    if not src_masters_dir.exists():
        print("Error: No slideMasters/ in source.", file=sys.stderr)
        sys.exit(1)

    # Determine numbering offset
    master_offset = 0
    layout_offset = 0

    if mode == "replace":
        print("  Mode: REPLACE - removing existing masters and layouts")
        _clear_directory_contents(tgt_masters_dir, "slideMaster*.xml")
        _clear_directory_contents(tgt_layouts_dir, "slideLayout*.xml")
    else:
        print("  Mode: MERGE - adding alongside existing")
        master_offset = _get_max_existing_number(tgt_masters_dir, "slideMaster")
        layout_offset = _get_max_existing_number(tgt_layouts_dir, "slideLayout")

    tgt_masters_dir.mkdir(parents=True, exist_ok=True)
    tgt_layouts_dir.mkdir(parents=True, exist_ok=True)
    (tgt_masters_dir / "_rels").mkdir(exist_ok=True)
    (tgt_layouts_dir / "_rels").mkdir(exist_ok=True)

    # Build layout renumber map: old_name -> new_name
    layout_map = {}  # e.g. "slideLayout1.xml" -> "slideLayout5.xml"
    master_map = {}  # e.g. "slideMaster1.xml" -> "slideMaster3.xml"

    # Collect source layouts referenced by masters
    src_master_files = sorted(src_masters_dir.glob("slideMaster*.xml"))

    # First pass: build layout mapping
    all_src_layouts = set()
    for master_file in src_master_files:
        rels_path = src_masters_dir / "_rels" / f"{master_file.name}.rels"
        if rels_path.exists():
            rdoc = parse_xml(str(rels_path))
            for rel in _find_all(rdoc, "Relationship"):
                target = _attr(rel, "Target", "")
                if "slideLayout" in target:
                    layout_name = Path(target).name
                    all_src_layouts.add(layout_name)

    # Assign new layout numbers
    next_layout_num = layout_offset + 1
    for layout_name in sorted(all_src_layouts, key=_extract_number):
        new_name = f"slideLayout{next_layout_num}.xml"
        layout_map[layout_name] = new_name
        next_layout_num += 1

    # Also map any layouts found in source dir but not referenced by masters
    if src_layouts_dir.exists():
        for lf in sorted(src_layouts_dir.glob("slideLayout*.xml"), key=lambda p: _extract_number(p.name)):
            if lf.name not in layout_map:
                new_name = f"slideLayout{next_layout_num}.xml"
                layout_map[lf.name] = new_name
                next_layout_num += 1

    # Assign new master numbers
    next_master_num = master_offset + 1
    for master_file in src_master_files:
        new_name = f"slideMaster{next_master_num}.xml"
        master_map[master_file.name] = new_name
        next_master_num += 1

    print(f"  Masters to copy: {len(master_map)}")
    print(f"  Layouts to copy: {len(layout_map)}")

    # Copy and rewrite layout files, tracking media references
    layout_media_to_copy = set()
    for old_name, new_name in layout_map.items():
        src_path = src_layouts_dir / old_name
        if not src_path.exists():
            print(f"  Warning: Source layout {old_name} not found, skipping")
            continue

        # Copy layout XML
        shutil.copy2(src_path, tgt_layouts_dir / new_name)

        # Copy and rewrite layout rels
        src_rels = src_layouts_dir / "_rels" / f"{old_name}.rels"
        if src_rels.exists():
            rdoc = parse_xml(str(src_rels))
            for rel in _find_all(rdoc, "Relationship"):
                target = _attr(rel, "Target", "")
                rel_type = _attr(rel, "Type", "").split("/")[-1]

                # Rewrite master references
                for old_master, new_master in master_map.items():
                    if old_master in target:
                        new_target = target.replace(old_master, new_master)
                        rel.setAttribute("Target", new_target)
                        break

                # Track media referenced by layouts (backgrounds, logos)
                if rel_type in ("image", "oleObject", "audio", "video"):
                    media_name = Path(target).name
                    layout_media_to_copy.add(media_name)

            _write_xml(rdoc, tgt_layouts_dir / "_rels" / f"{new_name}.rels")
        print(f"  Copied layout {old_name} -> {new_name}")

    # Copy and rewrite master files
    media_to_copy = set()
    for old_name, new_name in master_map.items():
        src_path = src_masters_dir / old_name
        shutil.copy2(src_path, tgt_masters_dir / new_name)

        # Copy and rewrite master rels
        src_rels = src_masters_dir / "_rels" / f"{old_name}.rels"
        if src_rels.exists():
            rdoc = parse_xml(str(src_rels))
            for rel in _find_all(rdoc, "Relationship"):
                target = _attr(rel, "Target", "")
                rel_type = _attr(rel, "Type", "").split("/")[-1]

                # Rewrite layout references
                for old_layout, new_layout in layout_map.items():
                    if old_layout in target:
                        new_target = target.replace(old_layout, new_layout)
                        rel.setAttribute("Target", new_target)
                        break

                # Track media to copy
                if rel_type in ("image", "oleObject", "audio", "video"):
                    media_name = Path(target).name
                    media_to_copy.add(media_name)

            _write_xml(rdoc, tgt_masters_dir / "_rels" / f"{new_name}.rels")
        print(f"  Copied master {old_name} -> {new_name}")

    # Copy referenced media (from both masters and layouts)
    all_media_to_copy = media_to_copy | layout_media_to_copy
    src_media = source_dir / "ppt" / "media"
    tgt_media = target_dir / "ppt" / "media"
    if all_media_to_copy and src_media.exists():
        tgt_media.mkdir(parents=True, exist_ok=True)
        for media_name in all_media_to_copy:
            src_media_path = src_media / media_name
            if src_media_path.exists():
                shutil.copy2(src_media_path, tgt_media / media_name)
                print(f"  Copied media/{media_name}")

    # Update presentation.xml - sldMasterIdLst
    pres_path = target_dir / "ppt" / "presentation.xml"
    if pres_path.exists():
        pdoc = parse_xml(str(pres_path))

        # Find or create sldMasterIdLst
        master_id_list = _find_all(pdoc, "sldMasterIdLst")
        if master_id_list:
            master_id_list = master_id_list[0]
        else:
            # Create it - insert before sldIdLst if exists
            pres_elem = pdoc.documentElement
            sld_id_list = _find_all(pdoc, "sldIdLst")
            master_id_list = pdoc.createElement("p:sldMasterIdLst")
            if sld_id_list:
                pres_elem.insertBefore(master_id_list, sld_id_list[0])
            else:
                pres_elem.appendChild(master_id_list)

        if mode == "replace":
            # Remove existing master id entries
            while master_id_list.firstChild:
                master_id_list.removeChild(master_id_list.firstChild)

        # Collect existing IDs to avoid collisions
        existing_ids = set()
        for mid in _find_all(pdoc, "sldMasterId"):
            id_val = _attr(mid, "id", "")
            if id_val:
                existing_ids.add(int(id_val))

        next_id = max(existing_ids) + 1 if existing_ids else 2147483648

        # Add new master references
        pres_rels_path = target_dir / "ppt" / "_rels" / "presentation.xml.rels"
        if pres_rels_path.exists():
            pres_rels_doc = parse_xml(str(pres_rels_path))
        else:
            # Create minimal rels doc
            pres_rels_doc = parse_xml(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )

        # Get next rId
        existing_rids = set()
        for rel in _find_all(pres_rels_doc, "Relationship"):
            rid = _attr(rel, "Id", "")
            m = re.search(r"(\d+)", rid)
            if m:
                existing_rids.add(int(m.group(1)))
        next_rid_num = max(existing_rids) + 1 if existing_rids else 1

        if mode == "replace":
            # Remove existing master rels
            rels_root = pres_rels_doc.documentElement
            to_remove = []
            for rel in _find_all(pres_rels_doc, "Relationship"):
                rel_type = _attr(rel, "Type", "")
                if "slideMaster" in rel_type:
                    to_remove.append(rel)
            for rel in to_remove:
                rel.parentNode.removeChild(rel)

        rels_root = pres_rels_doc.documentElement
        for old_name, new_name in sorted(master_map.items(), key=lambda x: _extract_number(x[1])):
            rid = f"rId{next_rid_num}"
            next_rid_num += 1

            # Add to sldMasterIdLst
            master_id_elem = pdoc.createElement("p:sldMasterId")
            master_id_elem.setAttribute("id", str(next_id))
            master_id_elem.setAttribute("r:id", rid)
            master_id_list.appendChild(master_id_elem)
            next_id += 1

            # Add relationship
            rel_elem = pres_rels_doc.createElement("Relationship")
            rel_elem.setAttribute("Id", rid)
            rel_elem.setAttribute("Type",
                                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
            rel_elem.setAttribute("Target", f"slideMasters/{new_name}")
            rels_root.appendChild(rel_elem)

            print(f"  Added master reference: {rid} -> {new_name}")

        _write_xml(pdoc, pres_path)
        _write_xml(pres_rels_doc, pres_rels_path)

    # Update [Content_Types].xml
    ct_path = target_dir / "[Content_Types].xml"
    if ct_path.exists():
        ct_doc = parse_xml(str(ct_path))
        ct_root = ct_doc.documentElement

        # Collect existing overrides
        existing_overrides = set()
        for override in _find_all(ct_doc, "Override"):
            existing_overrides.add(_attr(override, "PartName", ""))

        # Add overrides for new masters
        master_ct = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
        layout_ct = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"

        for new_name in master_map.values():
            part_name = f"/ppt/slideMasters/{new_name}"
            if part_name not in existing_overrides:
                ov = ct_doc.createElement("Override")
                ov.setAttribute("PartName", part_name)
                ov.setAttribute("ContentType", master_ct)
                ct_root.appendChild(ov)

        for new_name in layout_map.values():
            part_name = f"/ppt/slideLayouts/{new_name}"
            if part_name not in existing_overrides:
                ov = ct_doc.createElement("Override")
                ov.setAttribute("PartName", part_name)
                ov.setAttribute("ContentType", layout_ct)
                ct_root.appendChild(ov)

        # If replace mode, remove overrides for old files that no longer exist
        if mode == "replace":
            to_remove = []
            for override in _find_all(ct_doc, "Override"):
                pn = _attr(override, "PartName", "")
                if "/ppt/slideMasters/" in pn or "/ppt/slideLayouts/" in pn:
                    filename = Path(pn).name
                    if filename not in master_map.values() and filename not in layout_map.values():
                        # Check if file actually exists (could be from slides)
                        actual_path = target_dir / pn.lstrip("/")
                        if not actual_path.exists():
                            to_remove.append(override)
            for ov in to_remove:
                ov.parentNode.removeChild(ov)

        _write_xml(ct_doc, ct_path)
        print("  Updated [Content_Types].xml")

    print(f"\n  Transfer complete: {len(master_map)} masters, {len(layout_map)} layouts")
    return master_map, layout_map


def main():
    parser = argparse.ArgumentParser(
        description="Copy slide masters and layouts from source to target PPTX."
    )
    parser.add_argument("source_unpacked_dir", type=Path,
                        help="Path to source unpacked PPTX directory")
    parser.add_argument("target_unpacked_dir", type=Path,
                        help="Path to target unpacked PPTX directory")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--replace", action="store_const", const="replace", dest="mode",
                       help="Replace existing masters/layouts (default)")
    group.add_argument("--merge", action="store_const", const="merge", dest="mode",
                       help="Merge: add alongside existing")
    parser.set_defaults(mode="replace")
    args = parser.parse_args()

    if not args.source_unpacked_dir.exists():
        print(f"Error: Source not found: {args.source_unpacked_dir}", file=sys.stderr)
        sys.exit(1)
    if not args.target_unpacked_dir.exists():
        print(f"Error: Target not found: {args.target_unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Transferring masters and layouts ({args.mode} mode)...")
    transfer_masters_and_layouts(args.source_unpacked_dir, args.target_unpacked_dir, args.mode)


if __name__ == "__main__":
    main()
