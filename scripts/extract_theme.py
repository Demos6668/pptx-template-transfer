#!/usr/bin/env python3
"""Extract the complete theme from an unpacked PPTX into a portable bundle.

Usage:
    python extract_theme.py <source_unpacked_dir> <output_bundle_dir>

Extracts theme XML, referenced media, and creates a manifest.json.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

from defusedxml.minidom import parse as parse_xml

from inspect_template import inspect_theme, inspect_rels


def extract_theme(source_dir: Path, output_dir: Path):
    """Extract theme files and media into a bundle directory."""
    output_dir.mkdir(parents=True, exist_ok=True)

    theme_src = source_dir / "ppt" / "theme"
    if not theme_src.exists():
        print("Error: No ppt/theme/ directory found in source.", file=sys.stderr)
        sys.exit(1)

    manifest = {
        "source": str(source_dir),
        "themes": [],
        "media_files": [],
    }

    # Copy theme XML files
    themes_out = output_dir / "theme"
    themes_out.mkdir(exist_ok=True)

    for theme_file in sorted(theme_src.glob("theme*.xml")):
        shutil.copy2(theme_file, themes_out / theme_file.name)
        print(f"  Copied {theme_file.name}")

        # Copy theme rels if present
        rels_src = theme_src / "_rels" / f"{theme_file.name}.rels"
        if rels_src.exists():
            rels_out = themes_out / "_rels"
            rels_out.mkdir(exist_ok=True)
            shutil.copy2(rels_src, rels_out / rels_src.name)
            print(f"  Copied _rels/{rels_src.name}")

            # Extract media referenced by theme
            rels = inspect_rels(rels_src)
            for r in rels:
                if r["type"] in ("image", "oleObject", "audio", "video"):
                    # Target is relative like ../media/image1.png
                    media_rel = r["target"]
                    media_name = Path(media_rel).name
                    media_src_path = source_dir / "ppt" / "media" / media_name
                    if media_src_path.exists():
                        media_out = output_dir / "media"
                        media_out.mkdir(exist_ok=True)
                        shutil.copy2(media_src_path, media_out / media_name)
                        manifest["media_files"].append(media_name)
                        print(f"  Copied media/{media_name}")
                    else:
                        print(f"  Warning: Referenced media not found: {media_src_path}")

    # Build theme info for manifest
    theme_info = inspect_theme(source_dir)
    for ti in theme_info:
        manifest["themes"].append({
            "file": ti["file"],
            "name": ti.get("name", ""),
            "colors": ti.get("colors", {}),
            "fonts": ti.get("fonts", {}),
            "color_scheme_name": ti.get("color_scheme_name", ""),
            "font_scheme_name": ti.get("font_scheme_name", ""),
        })

    # Write manifest
    manifest_path = output_dir / "manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"  Wrote manifest.json")

    return manifest


def main():
    parser = argparse.ArgumentParser(
        description="Extract theme from an unpacked PPTX into a portable bundle."
    )
    parser.add_argument("source_unpacked_dir", type=Path,
                        help="Path to source unpacked PPTX directory")
    parser.add_argument("output_bundle_dir", type=Path,
                        help="Path to output bundle directory")
    args = parser.parse_args()

    if not args.source_unpacked_dir.exists():
        print(f"Error: Source not found: {args.source_unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Extracting theme from {args.source_unpacked_dir}...")
    manifest = extract_theme(args.source_unpacked_dir, args.output_bundle_dir)
    print(f"\nTheme bundle created at {args.output_bundle_dir}")
    print(f"  Themes: {len(manifest['themes'])}")
    print(f"  Media files: {len(manifest['media_files'])}")


if __name__ == "__main__":
    main()
