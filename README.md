# PPTX Template Transfer

Apply one deck's visual design — logos, shapes, backgrounds, branding — to another deck's text content. Built on `python-pptx`.

**Single file. One command. Any PPTX pair.**

## The Problem

Real-world branded PPTX files store all visual design (logos, decorative shapes, watermarks, backgrounds) as shapes inside individual slides, NOT in slide masters/layouts. The layouts are often just "BLANK." Traditional layout/master transfer does nothing useful for these files.

## The Solution

**Design mode** clones template slides as visual skeletons, classifies every shape's role, and injects content text ONLY into the right zones — leaving all branding untouched.

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

# Explicit design mode
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --mode design

# Verbose diagnostics (every shape classification decision)
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --verbose

# JSON report for automation
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --report report.json

# Manual slide mapping
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --slide-map mapping.json

# Skip speaker notes transfer
python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --no-notes
```

## Programmatic API

```python
from pptx_template_transfer import transfer, TransferConfig, Thresholds

# Basic usage
report = transfer(
    template=Path("template.pptx"),
    content=Path("content.pptx"),
    output=Path("output.pptx"),
)

# Custom config
config = TransferConfig(
    mode="design",
    verbose=True,
    preserve_notes=True,
    thresholds=Thresholds(title_min_font_pt=18, body_max_zones=3),
)
report = transfer(template, content, output, config)
```

## How Design Mode Works

1. **Validate** input files (valid ZIP, has slides, correct format)
2. **Extract** structured content from each slide (title, body paragraphs with subheading detection, tables, images, charts, speaker notes)
3. **Classify** each template slide's structure: title, narrative, list, grid, data, visual, section, closing
4. **Match** each content slide to the best template using type compatibility, text density, content structure, and position — with variety enforcement
5. **Clone** the matched template slide (preserving ALL shapes, backgrounds, transitions, media)
6. **Classify** every shape on the cloned slide: title | body | info | decorative | footer | media — with confidence scoring
7. **Inject** content ONLY into title/body/info zones — all other shapes protected
8. **Handle** tables (fill with content data, expand rows if needed), charts (best-effort cloning), and images (place in available space)
9. **Transfer** speaker notes from content to output
10. **Post-process** page numbers and dates (both text patterns and XML placeholders)
11. **Validate** output for orphaned shapes or missing content

### Shape Role Classification

Each shape is classified using multiple signals:

| Signal | Priority | Description |
|--------|----------|-------------|
| **Placeholder type** | Highest | PowerPoint's `PP_PLACEHOLDER.TITLE/BODY/FOOTER` |
| **Shape name** | High | "title", "body", "content" in auto-generated names |
| **Position + area** | Medium | Footer zone (bottom 10%), header zone (top 8%) |
| **Font size** | Medium | Adaptive thresholds based on slide-level statistics |
| **Text patterns** | Medium | "Page XX", "Confidential", dates, ALL-CAPS labels |
| **Repeated patterns** | Medium | 3+ shapes with similar size in a row/grid → decorative |
| **Word count** | Lower | ≤3 words → likely decorative, >10 words → likely body |

Each classification includes a **confidence score** (0.0-1.0) visible in verbose mode.

| Role | Max per slide | Action |
|------|--------------|--------|
| **title** | 1 | Text replaced with content title |
| **body** | 2 | Text replaced with content paragraphs (multi-level format preserved) |
| **info** | 1 | Text replaced with content summary (first 3 paragraphs) |
| **footer** | unlimited | Protected (page numbers auto-updated) |
| **media** | unlimited | Protected (pictures, charts, tables, groups, OLE) |
| **decorative** | unlimited | Protected (logos, labels, diagrams, branding) |

### Content Structure Extraction

Each content slide is parsed into:
- **Title**: Largest-font short text near the top
- **Body paragraphs**: Ordered text with subheading detection (bold/large font), indentation levels, per-run hyperlinks
- **Tables**: Cell text matrices with formatting
- **Images**: Content images >10% of slide area
- **Charts**: Chart elements for best-effort cloning
- **Speaker notes**: Full text from notes pane

### Formatting Preservation

- **Multi-level format**: Each indent level (0-8) gets its own saved formatting from the template
- **Bullet styles**: `buChar`, `buAutoNum`, `buFont` preserved from template paragraph properties
- **Bold/italic subheadings**: Detected in content and applied during injection
- **Overflow prevention**: Estimates shape text capacity; truncates with "..." if content exceeds it

### Slide Matching

```
Score = type_compatibility (0-40) + text_density_fit (0-25)
      + content_structure_fit (0-20) + position_preference (0-15)
```

**Variety enforcement**: No template slide gets >40% of mappings. Unused templates are redistributed.

### Table Handling

- Template has table + content has table → content fills template cells (preserving cell formatting)
- Content table has more rows → template table is expanded (row cloning)
- No table in template → content table element added to slide

### Robustness

- **Per-slide error isolation**: One failed slide doesn't abort the entire deck
- **Input validation**: Checks for valid ZIP, Content_Types.xml, non-zero slides
- **Broken relationship recovery**: Logs broken rIds instead of silent failure
- **Large deck support**: Pre-computed ShapeInfo cache, single-pass property extraction

## Verbose Mode

```
Slide 4/16:
  Content type: content (79 words, 0 table(s), 1 image(s))
  Template match: slide 6 (score=75, type=narrative)
  Shape classifications:
    Shape "Shape 0" (4.6% area, top 67%, conf=0.85) -> footer
    Shape "Text 3" (6.8% area, top 11%, conf=0.85) "Day2.Work..." -> title
    Shape "Text 7" (7.9% area, top 42%, conf=0.80) "24x7 ..." -> body
    ...
  Injected: title="Deployment Overview"
  Injected: body (75 words -> 2 zones)
  Protected: 27 shapes untouched
```

## JSON Report

Use `--report report.json` for machine-readable output:

```json
{
  "mode": "design",
  "slides": [
    {
      "index": 1,
      "content_type": "title",
      "template_slide": 1,
      "template_type": "title",
      "title": "Managed Detection and Response (MDR) Report",
      "word_count": 17,
      "status": "ok",
      "protected_shapes": 27,
      "classifications": [...]
    }
  ],
  "warnings": [],
  "errors": []
}
```

## Configuration

All classification thresholds are configurable via the `Thresholds` dataclass:

```python
from pptx_template_transfer import Thresholds

# For templates with smaller fonts
th = Thresholds(title_min_font_pt=16, decorative_max_font_pt=8)

# For templates with more body zones
th = Thresholds(body_max_zones=3, body_min_area_pct=3.0)
```

## Technical Notes

- Built on `python-pptx` with `lxml` for low-level XML manipulation
- Slide cloning deep-copies all shape elements and remaps relationship IDs
- Transitions are preserved from template slides
- Speaker notes transferred from content to output
- Text injection preserves template run properties per indent level
- Table cell formatting preserved during data fill (not stripped by `cell.text = ...`)
- Chart transfer is best-effort (copies chart part + element)
- Output uses the template's slide dimensions

## License

MIT
