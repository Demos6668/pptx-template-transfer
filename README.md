# PPTX Template Transfer

Apply one deck's visual design — logos, shapes, backgrounds, branding — to another deck's text content. Built on `python-pptx`.

**Single file. One command.**

## The Problem

Real-world branded PPTX files store all visual design (logos, decorative shapes, watermarks, backgrounds) as shapes inside individual slides, NOT in slide masters/layouts. The layouts are often just "BLANK." Traditional layout/master transfer does nothing useful for these files.

## The Solution

**Design mode** clones template slides as visual skeletons, then injects content text into them. The result has the template's complete visual identity with the content's text.

## Requirements

- Python 3.10+
- `python-pptx` and `Pillow`

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Auto-detect mode (design mode if layouts are blank, layout mode otherwise)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx

# Explicit design mode (clone template slides, inject content text)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --mode design

# Explicit layout mode (transfer theme + masters + layouts)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --mode layout

# Verbose diagnostics (see every shape classification decision)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --verbose

# Manual slide mapping (JSON: content slide number → template slide number)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --slide-map mapping.json
```

## How Design Mode Works

1. **Extract** structured content from each content slide (title, body paragraphs, tables, images)
2. **Classify** each template slide's structure: title, narrative, list, grid, data, visual, section, closing
3. **Match** each content slide to the best template slide using type compatibility, text density, content structure, and position preference — with variety enforcement so the output uses diverse template slides
4. **Clone** the matched template slide into the output (preserving ALL decorative elements, images, shapes, backgrounds)
5. **Classify** every shape on the cloned slide as: title | body | info | decorative | footer | media
6. **Inject** content text ONLY into title/body/info zones — all other shapes are protected
7. **Post-process** page numbers and dates

### Shape Role Classification

Each shape on a template slide is classified into one of six roles:

| Role | Description | Action |
|------|-------------|--------|
| **title** | Primary heading (largest font, top of slide) | Text replaced with content title |
| **body** | Main content area (large area, substantial text) | Text replaced with content paragraphs |
| **info** | Sidebar/panel (right side, moderate text) | Text replaced with content summary |
| **footer** | Bottom/top edge, page numbers, dates, legal text | Protected (page numbers auto-updated) |
| **media** | Pictures, charts, tables, OLE objects, groups | Protected |
| **decorative** | Logos, labels, diagrams, small shapes, patterns | Protected |

Classification uses position, font size, area, word count, ALL-CAPS detection, footer text patterns, and repeated pattern detection (for diagram/grid elements).

### Content Structure Extraction

Each content slide is parsed into:
- **Title**: Largest-font short text near the top
- **Body paragraphs**: Ordered text with subheading detection (bold/large font), indentation levels
- **Tables**: Cell text matrices with dimensions
- **Images**: Content images (>10% of slide area) for potential transfer

### Slide Matching Score

```
Score = type_compatibility (0-40) + text_density_fit (0-25)
      + content_structure_fit (0-20) + position_preference (0-15)
```

**Variety enforcement**: No template slide gets more than 40% of mappings. Unused templates are redistributed to ensure visual diversity.

### Table Handling

- Template has table + content has table → content data fills template table (preserving template styling)
- Content has table but template doesn't → table element is added to the output slide
- Template has table but content doesn't → table contents are cleared

### Image Handling

- Template images (logos, backgrounds) are always preserved
- Content images >10% of slide area are placed in available whitespace on the output slide

### Post-Processing

- Page numbers (`Page 01`, `Page 02`, ...) are updated to match actual slide position
- Dates in footer areas are updated to today's date

## Verbose Mode

Use `--verbose` (or `-v`) to see detailed diagnostics for every slide:

```
Slide 4/16:
  Content type: content (79 words, 0 table(s), 1 image(s))
  Template match: slide 6 (score=75, type=narrative)
  Shape classifications:
    Shape "Shape 0" (4.6% area, top 67%) -> footer
    Shape "Text 3" (6.8% area, top 11%) "Day2.Work™ ..." -> title
    Shape "Text 7" (7.9% area, top 42%) "• 24x7 ..." -> body
    ...
  Injected: title="Deployment Overview"
  Injected: body (75 words -> 2 zones)
  Protected: 27 shapes untouched
```

## Using with Claude

```bash
git clone https://github.com/Demos6668/pptx-template-transfer.git
pip install -r pptx-template-transfer/requirements.txt
python3 pptx-template-transfer/pptx_template_transfer.py template.pptx content.pptx output.pptx
```

## Technical Notes

- Built on `python-pptx` with `lxml` for low-level XML manipulation
- Slide cloning deep-copies all shape elements and remaps relationship IDs for media
- Text injection preserves template run properties (font, size, color, bold) — only text nodes are replaced
- Structured injection preserves paragraph hierarchy: subheadings as bold, indentation levels maintained
- Template images are always preserved (they're branding); content images transferred where space allows
- Output uses the template's slide dimensions

## License

MIT
