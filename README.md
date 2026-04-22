# pptx_template_converter

Apply a corporate PowerPoint template (theme, colors, fonts, slide master branding)
to an existing `.pptx` deck — for example, an AI-generated deck — without ever
uploading the corporate template to any third-party service.

The converter runs entirely on your local machine. The corporate template file
stays on disk.

## Directory layout

```
input/      drop the source .pptx here   (gitignored; one sample fixture committed)
template/   drop the corporate template  (gitignored; one sample fixture committed)
output/     converted deck is written here (gitignored entirely)
scripts/    helper scripts (fixture generator)
```

Real `.pptx` / `.potx` files are gitignored globally — only the two synthetic
fixtures are tracked so the repo clones cleanly with something to test against.

## Quickstart (once the converter is implemented)

```bash
pip install -r requirements.txt

# Minimal theme swap — carries over theme colors, fonts, and master branding:
python convert.py \
  --input    input/sample_input.pptx \
  --template template/sample_template.pptx \
  --output   output/sample_merged.pptx

# With a color map — remaps source hardcoded RGBs onto template theme slots.
# Essential for AI-generated decks that bake explicit colors into every shape.
python convert.py \
  --input    input/sample_input.pptx \
  --template template/sample_template.pptx \
  --output   output/sample_merged.pptx \
  --color-map mappings/keyhole_to_corporate.json
```

### Color map format

`mappings/*.json`:

```json
{
  "description": "...",
  "map": {
    "1A1A2E": "lt1",
    "00D4FF": "accent1",
    "FF4444": "#D32F2F"
  }
}
```

Keys are 6-hex source colors (no `#`). Values are either a theme-slot name
(`dk1`, `lt1`, `dk2`, `lt2`, `accent1`–`accent6`) or a `#RRGGBB` literal.
Theme-slot targets emit `<a:schemeClr>` references, so the color becomes
theme-aware and will follow the template's theme if you later swap it.

The color report at the end of each run lists every source color and whether
it was remapped — iterate your mapping until the "still hardcoded" list is
empty (or limited to colors you deliberately want to preserve).

### Title and subtitle recoloring

AI-generated decks almost always use absolutely-positioned text boxes rather
than title placeholders, so the template's `titleStyle` / `bodyStyle` colors
never get a chance to apply. `convert.py` detects the title and subtitle
heuristically (topmost wide shape = title; next wide shape still near the top
= subtitle) and rewrites their text color to match the template master.

Override per-run with `--title-color` / `--subtitle-color`. Accepted values:

- `auto` (default) — reads from template master `titleStyle` / `bodyStyle`.
- `none` — skip that pass entirely.
- a theme slot name: `dk1`, `lt1`, `dk2`, `lt2`, `tx1`, `tx2`, `bg1`, `bg2`,
  `accent1` .. `accent6`, `hlink`, `folHlink`.
- a `#RRGGBB` hex literal.

The report at the end of each run lists the detected title and subtitle text
for every slide so you can verify the heuristic didn't misfire.

## Sample fixtures

- `input/sample_input.pptx` — a five-slide deck in the style of an AI-generated
  presentation: 16:9, "Blank" layout throughout, absolutely-positioned text
  boxes and shapes, some hardcoded RGB colors.
- `template/sample_template.pptx` — a synthetic "Acme Corporation" template
  with a custom theme color scheme, theme fonts, and branded slide master.

Regenerate both with:

```bash
python scripts/make_samples.py
```

## Approach

**Strategy A — Theme swap** (what this tool does): graft the template's theme,
slide master, and slide layouts onto the source deck. Anything in the source
that references theme colors/fonts picks up the corporate branding. Absolutely-
positioned shapes and hardcoded colors in the source are preserved as-is.

Hardcoded colors (explicit RGB) will not be re-themed — the tool reports them
per slide so you can decide whether to touch up manually.

## Privacy

No telemetry, no uploads, no cloud calls. The script opens the template and
source files with `python-pptx` and writes a merged `.pptx` to `output/`.
