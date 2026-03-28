---
name: pptx-template-transfer
description: Transfer themes, slide masters, slide layouts, and design elements from one PowerPoint file to another while preserving content. Use when the user asks to apply a template/design/theme from one PPTX to another, rebrand a presentation, or match styling between decks.
---

# PPTX Template Transfer

Transfer themes, slide masters, slide layouts, and design elements from one PowerPoint file to another while preserving content.

## When to Use

**Trigger:** User asks to apply a template/design/theme from one PPTX to another, or asks to rebrand a presentation, or wants to change the look of slides to match a corporate template.

**Keywords:** "apply template", "transfer design", "rebrand presentation", "use this theme on", "match the style of", "apply corporate template"

## Quick One-Liner

```bash
python ~/.claude/skills/pptx-template-transfer/scripts/apply_template.py template.pptx content.pptx output.pptx
```

This runs the full 11-step pipeline automatically: unpack, inspect, extract theme, apply theme, smart layout mapping (content-aware), transfer masters/layouts, translate mapping, remap slides, adapt text colors for contrast, reconcile IDs, clean and pack.

## Step-by-Step Manual Approach

When you need finer control or debugging:

### 1. Inspect both files
```bash
# Understand what you're working with
python scripts/inspect_template.py /tmp/template_unpacked
python scripts/inspect_template.py /tmp/content_unpacked
python scripts/inspect_template.py /tmp/content_unpacked --json > content_info.json
```

### 2. Extract and apply theme (colors, fonts)
```bash
python scripts/extract_theme.py /tmp/template_unpacked /tmp/theme_bundle
python scripts/apply_theme.py /tmp/theme_bundle /tmp/content_unpacked
```

### 3. Transfer slide masters and layouts
```bash
# Replace mode (default): wipe existing masters/layouts, use template's
python scripts/transfer_masters_and_layouts.py /tmp/template_unpacked /tmp/content_unpacked --replace

# Merge mode: add template's masters alongside existing ones
python scripts/transfer_masters_and_layouts.py /tmp/template_unpacked /tmp/content_unpacked --merge
```

### 4. Map and remap layouts
```bash
# Auto-generate mapping
python scripts/map_layouts.py /tmp/template_unpacked /tmp/content_unpacked --output mapping.json

# Review mapping.json, edit if needed, then apply
python scripts/remap_slides.py /tmp/content_unpacked mapping.json
```

### 5. Adapt text colors (after remap, before reconcile)
```bash
# Fix dark-on-dark and light-on-light text after background changes
python scripts/adapt_text_colors.py /tmp/content_unpacked

# Preview changes without modifying files
python scripts/adapt_text_colors.py /tmp/content_unpacked --dry-run
```

### 6. Fix relationship IDs
```bash
python scripts/reconcile_rids.py /tmp/content_unpacked
```

### 7. Clean and pack
Use the existing PPTX skill's `clean.py` and `pack.py`, or the orchestrator handles this automatically.

## With a Custom Layout Map

If auto-mapping isn't correct, create a `mapping.json` manually:

```json
{
  "mappings": [
    {
      "target_slide": "slide1.xml",
      "suggested_layout": "slideLayout1.xml"
    },
    {
      "target_slide": "slide2.xml",
      "suggested_layout": "slideLayout3.xml"
    }
  ]
}
```

Then pass it to the orchestrator:

```bash
python scripts/apply_template.py template.pptx content.pptx output.pptx --layout-map mapping.json
```

## Scripts Reference

| Script | Purpose |
|--------|---------|
| `inspect_template.py` | Analyze template structure (themes, masters, layouts, placeholders) |
| `extract_theme.py` | Extract theme to portable bundle (XML + media + manifest) |
| `apply_theme.py` | Apply theme bundle to target PPTX |
| `transfer_masters_and_layouts.py` | Copy masters/layouts with --replace or --merge mode |
| `map_layouts.py` | Smart layout matching: name, synonym, content analysis, placeholders |
| `remap_slides.py` | Update slide .rels to point to new layouts |
| `adapt_text_colors.py` | Fix dark-on-dark / light-on-light text after background changes |
| `reconcile_rids.py` | Fix duplicate relationship ID collisions |
| `apply_template.py` | Full 11-step pipeline orchestrator |

## Technical Notes

- **XML parsing:** All scripts use `defusedxml.minidom` exclusively. Do NOT use `xml.etree.ElementTree` as it corrupts OOXML namespace declarations.
- **Dependencies:** Only `defusedxml` and Python stdlib (zipfile, json, pathlib, shutil, argparse, tempfile).
- All scripts work on unpacked PPTX directories (ZIP-extracted).
- All scripts are importable as Python modules and runnable as CLI tools.
- Relationship paths in OOXML are relative (e.g., `../media/image1.png`).

## Common Failure Modes

### Output opens but slides are blank
The layout remap pointed slides to layouts that don't exist. Run `inspect_template.py` on the output to check that all referenced layouts exist.

### "Repair" dialog in PowerPoint
Usually a relationship ID collision or missing `[Content_Types].xml` entry. Run `reconcile_rids.py` and verify content types include all file extensions.

### Colors/fonts don't change
The theme was applied but slides have hardcoded colors (srgbClr values instead of theme references). Theme transfer only affects theme-referenced colors. The `adapt_text_colors.py` script handles the most common case (contrast flipping), but full recoloring of hardcoded values requires manual XML editing.

### Text invisible after transfer (dark on dark / light on light)
The `adapt_text_colors.py` step (step 9 in the pipeline) should handle this automatically. If it missed some text, run it standalone with `--dry-run` first to see what it would change. It only modifies explicit `srgbClr` values, not theme-referenced colors (`schemeClr`), and skips gradients.

### Missing backgrounds/logos
Masters reference media files that weren't copied. Run `inspect_template.py` and check the media references, then ensure they exist in `ppt/media/`.

### Layout mapping is wrong
Use `map_layouts.py --output mapping.json`, review the JSON, manually correct entries, then pass it to `remap_slides.py` or `apply_template.py --layout-map`.

## Integration with Existing PPTX Skill

After template transfer, use the existing `/mnt/skills/public/pptx/` skill to:
- Edit text content on individual slides
- Duplicate slides
- Add/remove slides
- Clean orphaned files with `clean.py`
- Pack with `pack.py`

The template transfer scripts handle the design layer; the content editing skill handles the content layer.
