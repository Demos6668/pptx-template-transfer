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
```

## How Design Mode Works

1. **Classify** each slide in both decks by type: title, content, section, data, closing, image
2. **Match** each content slide to the best template slide using type similarity, text density, position in deck, and shape count
3. **Clone** the matched template slide into the output (preserving ALL decorative elements, images, shapes, backgrounds)
4. **Inject** the content text into the clone's "content zones" (large text shapes), preserving the template's formatting
5. **Output** has template's theme, masters, layouts, AND per-slide decorative elements — with the content's text

### Shape Classification

Each shape in a template slide is classified as:
- **Design** (preserved as-is): images, logos, decorative shapes, branding labels, page numbers, small-font elements
- **Content zones** (text replaced): large text shapes — title, subtitle, body text areas

### Slide Matching Score

```
Score = type_match(50) + text_density(20) + deck_position(15) + shape_count(15)
```

A single template slide can be reused for multiple content slides.

## Using with Claude

```bash
git clone https://github.com/Demos6668/pptx-template-transfer.git
pip install -r pptx-template-transfer/requirements.txt
python3 pptx-template-transfer/pptx_template_transfer.py template.pptx content.pptx output.pptx
```

## Technical Notes

- Built on `python-pptx` with `lxml` for low-level XML manipulation
- Slide cloning deep-copies all shape elements and remaps relationship IDs for media
- Text injection preserves template run properties (font, size, color, bold) — only `<a:t>` text nodes are replaced
- Template images are always preserved (they're branding); content images are not transferred (they'd conflict with the template layout)
- Output uses the template's slide dimensions

## License

MIT
