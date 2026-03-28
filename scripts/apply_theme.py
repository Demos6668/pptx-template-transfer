#!/usr/bin/env python3
"""Apply an extracted theme bundle to a target unpacked PPTX.

Usage:
    python apply_theme.py <theme_bundle_dir> <target_unpacked_dir>

Replaces theme XML, copies media, fixes rels and content types.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

from defusedxml.minidom import parse as parse_xml


def _find_all(parent, tag_local):
    """Find all descendant elements matching a local tag name."""
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
    # minidom produces '<?xml version="1.0" ?>' but OOXML requires encoding+standalone
    xml_str = xml_str.replace(
        '<?xml version="1.0" ?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)


def update_content_types(target_dir: Path, new_media_files: list):
    """Add content type entries for new media types if needed."""
    ct_path = target_dir / "[Content_Types].xml"
    if not ct_path.exists():
        print("  Warning: [Content_Types].xml not found")
        return

    doc = parse_xml(str(ct_path))
    types_elem = doc.documentElement

    # Collect existing extensions
    existing_exts = set()
    for ext_elem in _find_all(doc, "Default"):
        existing_exts.add(_attr(ext_elem, "Extension", "").lower())

    # Map of common extensions to content types
    ext_content_types = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tiff": "image/tiff",
        "tif": "image/tiff",
        "svg": "image/svg+xml",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
        "wdp": "image/vnd.ms-photo",
    }

    added = []
    for media_file in new_media_files:
        ext = Path(media_file).suffix.lstrip(".").lower()
        if ext and ext not in existing_exts and ext in ext_content_types:
            new_default = doc.createElement("Default")
            new_default.setAttribute("Extension", ext)
            new_default.setAttribute("ContentType", ext_content_types[ext])
            types_elem.appendChild(new_default)
            existing_exts.add(ext)
            added.append(ext)

    if added:
        _write_xml(doc, ct_path)
        print(f"  Added content types for: {', '.join(added)}")


def apply_theme(bundle_dir: Path, target_dir: Path):
    """Apply theme bundle to target unpacked PPTX."""
    manifest_path = bundle_dir / "manifest.json"
    if not manifest_path.exists():
        print("Error: manifest.json not found in bundle.", file=sys.stderr)
        sys.exit(1)

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    target_theme_dir = target_dir / "ppt" / "theme"
    target_theme_dir.mkdir(parents=True, exist_ok=True)

    # Check for slide-level theme overrides
    slides_dir = target_dir / "ppt" / "slides"
    if slides_dir.exists():
        for rels_file in (slides_dir / "_rels").glob("*.rels") if (slides_dir / "_rels").exists() else []:
            try:
                rdoc = parse_xml(str(rels_file))
                for rel in _find_all(rdoc, "Relationship"):
                    rel_type = _attr(rel, "Type", "").split("/")[-1]
                    if rel_type == "theme":
                        print(f"  Warning: {rels_file.stem} has a slide-level theme override")
            except Exception:
                pass

    # Copy theme XML files
    bundle_theme_dir = bundle_dir / "theme"
    if bundle_theme_dir.exists():
        for theme_file in sorted(bundle_theme_dir.glob("theme*.xml")):
            dest = target_theme_dir / theme_file.name
            shutil.copy2(theme_file, dest)
            print(f"  Replaced {theme_file.name}")

        # Copy theme rels
        bundle_rels_dir = bundle_theme_dir / "_rels"
        if bundle_rels_dir.exists():
            target_rels_dir = target_theme_dir / "_rels"
            target_rels_dir.mkdir(exist_ok=True)
            for rels_file in bundle_rels_dir.glob("*.rels"):
                shutil.copy2(rels_file, target_rels_dir / rels_file.name)
                print(f"  Replaced _rels/{rels_file.name}")

    # Copy media files
    bundle_media_dir = bundle_dir / "media"
    new_media = []
    if bundle_media_dir.exists():
        target_media_dir = target_dir / "ppt" / "media"
        target_media_dir.mkdir(parents=True, exist_ok=True)
        for media_file in bundle_media_dir.iterdir():
            shutil.copy2(media_file, target_media_dir / media_file.name)
            new_media.append(media_file.name)
            print(f"  Copied media/{media_file.name}")

    # Update content types for new media
    if new_media:
        update_content_types(target_dir, new_media)

    print("  Theme applied successfully.")


def main():
    parser = argparse.ArgumentParser(
        description="Apply an extracted theme bundle to a target unpacked PPTX."
    )
    parser.add_argument("theme_bundle_dir", type=Path,
                        help="Path to theme bundle directory")
    parser.add_argument("target_unpacked_dir", type=Path,
                        help="Path to target unpacked PPTX directory")
    args = parser.parse_args()

    if not args.theme_bundle_dir.exists():
        print(f"Error: Bundle not found: {args.theme_bundle_dir}", file=sys.stderr)
        sys.exit(1)
    if not args.target_unpacked_dir.exists():
        print(f"Error: Target not found: {args.target_unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Applying theme bundle to {args.target_unpacked_dir}...")
    apply_theme(args.theme_bundle_dir, args.target_unpacked_dir)


if __name__ == "__main__":
    main()
