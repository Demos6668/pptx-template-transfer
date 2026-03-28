# PPTX Template Transfer

One-command template/theme transfer between PowerPoint files. Applies a template PPTX's design (theme, colors, fonts, slide masters, layouts) to a content PPTX while preserving all text and images.

## Features

- **Full theme transfer** - colors, fonts, format schemes
- **Smart layout mapping** - content-aware analysis infers slide type (title/content/section/blank) and maps to the best branded layout, even when source slides use generic layouts like "DEFAULT"
- **Text color adaptation** - automatically flips dark-on-dark and light-on-light text after background changes
- **Relationship ID reconciliation** - fixes rId collisions that cause PowerPoint repair dialogs
- **Single command** - 11-step pipeline runs automatically

## Requirements

- Python 3.10+
- `defusedxml` (`pip install defusedxml`)
- No other dependencies (stdlib only)

## Quick Start

```bash
pip install defusedxml

python scripts/apply_template.py template.pptx content.pptx output.pptx
```

This runs the full pipeline:
1. Unpack both PPTX files
2. Inspect structures
3. Extract theme from template
4. Apply theme to content
5. Map layouts (content-aware smart matching)
6. Transfer masters and layouts
7. Translate layout filenames
8. Remap slides to new layouts
9. Adapt text colors for contrast
10. Reconcile relationship IDs
11. Clean orphaned files and pack

## Custom Layout Map

If auto-mapping isn't correct, provide a manual mapping:

```bash
# Generate mapping, review it, then apply
python scripts/map_layouts.py template_unpacked/ content_unpacked/ --output mapping.json

# Use custom mapping with the orchestrator
python scripts/apply_template.py template.pptx content.pptx output.pptx --layout-map mapping.json
```

## Scripts

| Script | Purpose |
|--------|---------|
| `apply_template.py` | Full 11-step pipeline orchestrator |
| `inspect_template.py` | Analyze PPTX structure (themes, masters, layouts, placeholders) |
| `extract_theme.py` | Extract theme to portable bundle (XML + media + manifest) |
| `apply_theme.py` | Apply theme bundle to target PPTX |
| `transfer_masters_and_layouts.py` | Copy masters/layouts with `--replace` or `--merge` mode |
| `map_layouts.py` | Smart layout matching: name, synonym, content analysis, placeholders |
| `remap_slides.py` | Update slide .rels to point to new layouts |
| `adapt_text_colors.py` | Fix dark-on-dark / light-on-light text after background changes |
| `reconcile_rids.py` | Fix duplicate relationship ID collisions |

All scripts are importable as Python modules and runnable as CLI tools.

## Using with Claude on claude.ai

This repo is designed to work as a skill for Claude. When Claude needs to apply a template to a PPTX:

1. Clone this repo into the working directory
2. Run `apply_template.py` with the template and content files
3. If the auto-mapping needs adjustment, generate and edit `mapping.json`

See the `SKILL.md` file for Claude-specific integration instructions.

## How It Works

PPTX files are ZIP archives containing XML files following the OOXML standard. The pipeline:

1. **Theme transfer** replaces the color scheme, font scheme, and format scheme
2. **Layout mapping** analyzes each slide's content (font sizes, positions, text count) to infer whether it's a title slide, content slide, section header, etc., then maps to the best matching branded layout
3. **Text color adaptation** walks the slide -> layout -> master hierarchy to determine effective background color, then flips hardcoded text colors that would be invisible (dark text on dark background or light text on light background)

## Technical Notes

- All XML parsing uses `defusedxml.minidom` exclusively (not `xml.etree.ElementTree`, which corrupts OOXML namespace declarations)
- Only modifies explicit `srgbClr` values for color adaptation; theme-referenced colors (`schemeClr`) adapt automatically
- Relationship paths in OOXML are relative (e.g., `../media/image1.png`)

## License

MIT
