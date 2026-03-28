#!/usr/bin/env python3
"""One-command template transfer orchestrator.

Usage:
    python apply_template.py <template.pptx> <content.pptx> <output.pptx> [--layout-map mapping.json]

Takes a template PPTX (design source) and a content PPTX (slides with text/images),
produces output.pptx with the template's design and the content's text.

Workflow:
1. Unpack both PPTX files
2. Inspect both (for logging)
3. Extract theme from template
4. Apply theme to content
5. Map layouts (before transfer to preserve original names)
6. Transfer masters and layouts from template to content
7. Translate mapping filenames to post-transfer names
8. Remap slides to new layouts
9. Adapt text colors for contrast (dark-on-dark / light-on-light)
10. Reconcile relationship IDs
11. Clean orphaned files and pack result
"""

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

# Import sibling modules
SCRIPTS_DIR = Path(__file__).parent

# Add scripts dir to path for imports
sys.path.insert(0, str(SCRIPTS_DIR))

from inspect_template import inspect, print_report
from extract_theme import extract_theme
from apply_theme import apply_theme
from transfer_masters_and_layouts import transfer_masters_and_layouts
from map_layouts import map_layouts
from remap_slides import remap_slides
from reconcile_rids import reconcile_rids
from adapt_text_colors import adapt_text_colors


def _unpack_pptx(pptx_path: Path, output_dir: Path):
    """Unpack a PPTX file (it's just a ZIP)."""
    import zipfile
    output_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(pptx_path), "r") as zf:
        zf.extractall(str(output_dir))
    print(f"  Unpacked {pptx_path.name} to {output_dir}")


def _pack_pptx(source_dir: Path, output_path: Path):
    """Pack a directory back into a PPTX (ZIP) file."""
    import zipfile
    with zipfile.ZipFile(str(output_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_file():
                arcname = file_path.relative_to(source_dir)
                zf.write(str(file_path), str(arcname))
    print(f"  Packed result to {output_path}")


def _clean_orphans(unpacked_dir: Path):
    """Remove files not referenced by any .rels or [Content_Types].xml.

    Simple heuristic: collect all Target values from .rels files and
    all PartName values from [Content_Types].xml, then check for unreferenced
    files in ppt/media/, ppt/embeddings/, etc.
    """
    from defusedxml.minidom import parse as parse_xml

    referenced = set()

    # Collect from .rels files
    for rels_file in unpacked_dir.rglob("*.rels"):
        try:
            doc = parse_xml(str(rels_file))
            for elem in doc.getElementsByTagName("*"):
                local = elem.localName or elem.tagName.split(":")[-1]
                if local == "Relationship":
                    target = elem.getAttribute("Target") if elem.hasAttribute("Target") else ""
                    if target:
                        # Resolve relative path
                        rels_parent = rels_file.parent.parent  # up from _rels/
                        resolved = (rels_parent / target).resolve()
                        referenced.add(resolved)
        except Exception:
            pass

    # Collect from [Content_Types].xml
    ct_path = unpacked_dir / "[Content_Types].xml"
    if ct_path.exists():
        try:
            doc = parse_xml(str(ct_path))
            for elem in doc.getElementsByTagName("*"):
                local = elem.localName or elem.tagName.split(":")[-1]
                if local == "Override":
                    part = elem.getAttribute("PartName") if elem.hasAttribute("PartName") else ""
                    if part:
                        resolved = (unpacked_dir / part.lstrip("/")).resolve()
                        referenced.add(resolved)
        except Exception:
            pass

    # Check media directory for orphans
    media_dir = unpacked_dir / "ppt" / "media"
    removed = 0
    if media_dir.exists():
        for media_file in media_dir.iterdir():
            if media_file.resolve() not in referenced:
                # Double check: search all XML files for the filename
                found = False
                for xml_file in unpacked_dir.rglob("*.xml"):
                    try:
                        content = xml_file.read_text(encoding="utf-8", errors="ignore")
                        if media_file.name in content:
                            found = True
                            break
                    except Exception:
                        pass
                for rels_file in unpacked_dir.rglob("*.rels"):
                    try:
                        content = rels_file.read_text(encoding="utf-8", errors="ignore")
                        if media_file.name in content:
                            found = True
                            break
                    except Exception:
                        pass

                if not found:
                    media_file.unlink()
                    removed += 1
                    print(f"  Removed orphan: media/{media_file.name}")

    if removed:
        print(f"  Cleaned {removed} orphaned media files")
    else:
        print("  No orphaned files found")


def apply_template_pipeline(
    template_pptx: Path,
    content_pptx: Path,
    output_pptx: Path,
    layout_map_path: Path = None,
):
    """Run the full template transfer pipeline."""
    # Create temp working directory
    work_dir = Path(tempfile.mkdtemp(prefix="pptx_transfer_"))
    template_dir = work_dir / "template_unpacked"
    content_dir = work_dir / "content_unpacked"
    theme_bundle_dir = work_dir / "theme_bundle"
    mapping_path = layout_map_path or (work_dir / "layout_mapping.json")

    success = False
    try:
        # Step 1: Unpack
        print("\n[1/11] Unpacking PPTX files...")
        _unpack_pptx(template_pptx, template_dir)
        _unpack_pptx(content_pptx, content_dir)

        # Step 2: Inspect (for logging)
        print("\n[2/11] Inspecting template structure...")
        template_info = inspect(template_dir)
        print_report(template_info)

        print("Inspecting content structure...")
        content_info = inspect(content_dir)
        print_report(content_info)

        # Step 3: Extract theme
        print("\n[3/11] Extracting theme from template...")
        extract_theme(template_dir, theme_bundle_dir)

        # Step 4: Apply theme
        print("\n[4/11] Applying theme to content...")
        apply_theme(theme_bundle_dir, content_dir)

        # Step 5: Map layouts BEFORE transfer (while original layout names still exist)
        print("\n[5/11] Mapping layouts...")
        if layout_map_path and layout_map_path.exists():
            print(f"  Using provided layout map: {layout_map_path}")
            if layout_map_path != mapping_path:
                shutil.copy2(layout_map_path, mapping_path)
            with open(mapping_path, "r", encoding="utf-8") as f:
                mapping_result = json.load(f)
        else:
            mapping_result = map_layouts(template_dir, content_dir)
            with open(mapping_path, "w", encoding="utf-8") as f:
                json.dump(mapping_result, f, indent=2)
            high = sum(1 for m in mapping_result["mappings"] if m["confidence"] == "high")
            med = sum(1 for m in mapping_result["mappings"] if m["confidence"] == "medium")
            low = sum(1 for m in mapping_result["mappings"] if m["confidence"] == "low")
            print(f"  Auto-generated mapping with {len(mapping_result['mappings'])} entries")
            print(f"  Confidence: {high} high, {med} medium, {low} low")

        # Step 6: Transfer masters and layouts
        print("\n[6/11] Transferring masters and layouts...")
        _master_map, layout_rename_map = transfer_masters_and_layouts(
            template_dir, content_dir, mode="replace"
        )

        # Step 7: Translate mapping filenames to post-transfer names
        # layout_rename_map: {"slideLayout1.xml" (template) -> "slideLayout1.xml" (target)}
        print("\n[7/11] Translating layout mapping to new filenames...")
        for entry in mapping_result.get("mappings", []):
            old_suggested = entry.get("suggested_layout", "")
            if old_suggested in layout_rename_map:
                entry["suggested_layout"] = layout_rename_map[old_suggested]
                if old_suggested != layout_rename_map[old_suggested]:
                    print(f"  {entry['target_slide']}: {old_suggested} -> {layout_rename_map[old_suggested]}")
        with open(mapping_path, "w", encoding="utf-8") as f:
            json.dump(mapping_result, f, indent=2)

        # Step 8: Remap slides to new layouts
        print("\n[8/11] Remapping slides to new layouts...")
        remap_slides(content_dir, mapping_path)

        # Step 9: Adapt text colors for contrast
        print("\n[9/11] Adapting text colors for background contrast...")
        adapt_text_colors(content_dir)

        # Step 10: Reconcile rIds
        print("\n[10/11] Reconciling relationship IDs...")
        reconcile_rids(content_dir)

        # Step 10: Clean orphans and pack
        print("\n[11/11] Cleaning orphaned files and packing...")
        _clean_orphans(content_dir)
        _pack_pptx(content_dir, output_pptx)

        success = True
        print(f"\nTemplate transfer complete! Output: {output_pptx}")

    except Exception as e:
        print(f"\nError during template transfer: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        print(f"\nIntermediate files preserved at: {work_dir}", file=sys.stderr)
        print("You can inspect them for debugging.", file=sys.stderr)
        sys.exit(1)

    finally:
        if success:
            # Clean up temp dir on success
            try:
                shutil.rmtree(work_dir)
            except Exception:
                pass
        else:
            print(f"\nWork directory preserved for debugging: {work_dir}")


def main():
    parser = argparse.ArgumentParser(
        description="Apply a template PPTX's design to a content PPTX."
    )
    parser.add_argument("template_pptx", type=Path,
                        help="Path to template PPTX (design source)")
    parser.add_argument("content_pptx", type=Path,
                        help="Path to content PPTX (slides with text)")
    parser.add_argument("output_pptx", type=Path,
                        help="Path for output PPTX")
    parser.add_argument("--layout-map", type=Path, default=None,
                        help="Optional pre-built layout mapping JSON")
    args = parser.parse_args()

    if not args.template_pptx.exists():
        print(f"Error: Template not found: {args.template_pptx}", file=sys.stderr)
        sys.exit(1)
    if not args.content_pptx.exists():
        print(f"Error: Content not found: {args.content_pptx}", file=sys.stderr)
        sys.exit(1)

    apply_template_pipeline(
        args.template_pptx,
        args.content_pptx,
        args.output_pptx,
        args.layout_map,
    )


if __name__ == "__main__":
    main()
