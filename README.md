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

# Convert the bundled sample fixtures:
python convert.py \
  --input    input/sample_input.pptx \
  --template template/sample_template.pptx \
  --output   output/sample_merged.pptx
```

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
