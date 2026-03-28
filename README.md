# PPTX Template Transfer

One-command template/theme transfer between PowerPoint files. Applies a template PPTX's design (theme, colors, fonts, slide masters, layouts) to a content PPTX while preserving all text and images.

**Single file. Zero config.**

## Features

- **Full theme transfer** — colors, fonts, format schemes
- **Smart layout mapping** — content-aware analysis infers slide type (title/content/section/blank) and maps to the best branded layout, even when source slides use generic layouts like "DEFAULT"
- **Text color adaptation** — automatically flips dark-on-dark and light-on-light text after background changes
- **Relationship ID reconciliation** — fixes rId collisions that cause PowerPoint repair dialogs
- **Single command** — full pipeline runs automatically

## Requirements

- Python 3.10+
- `defusedxml` (`pip install defusedxml`)
- No other dependencies (stdlib only)

## Usage

```bash
pip install defusedxml

python3 pptx_template_transfer.py template.pptx content.pptx output.pptx
```

With a custom layout mapping:

```bash
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --layout-map mapping.json
```

## Using with Claude on claude.ai

This repo is designed to work as a tool for Claude. When Claude needs to apply a template to a PPTX:

1. Clone this repo into the working directory
2. Run `python3 pptx_template_transfer.py template.pptx content.pptx output.pptx`
3. If the auto-mapping needs adjustment, create a `mapping.json` and pass `--layout-map mapping.json`

## How It Works

PPTX files are ZIP archives containing XML files following the OOXML standard. The pipeline:

1. **Theme transfer** — replaces the color scheme, font scheme, and format scheme
2. **Layout mapping** — analyzes each slide's content (font sizes, positions, text count) to infer whether it's a title slide, content slide, section header, etc., then maps to the best matching branded layout
3. **Text color adaptation** — walks the slide -> layout -> master hierarchy to determine effective background color, then flips hardcoded text colors that would be invisible

## Technical Notes

- All XML parsing uses `defusedxml.minidom` exclusively (not `xml.etree.ElementTree`, which corrupts OOXML namespace declarations)
- Only modifies explicit `srgbClr` values for color adaptation; theme-referenced colors (`schemeClr`) adapt automatically
- Relationship paths in OOXML are relative (e.g., `../media/image1.png`)

## License

MIT
