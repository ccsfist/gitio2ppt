# gitio2ppt

Convert a GitHub.io / MkDocs Markdown `.zip` archive into a styled PowerPoint deck (`.pptx`) and a matching self-contained HTML presentation (`.html`) with live embedded Streamlit dashboards.

## Features

- **Clean PowerPoint Output:** Strips unused `"Click to add..."` placeholders, calculates element heights dynamically to prevent overlapping text, and automatically splits crowded slides onto `(Cont.)` slides.
- **Live HTML Deck:** Generates a matching 16:9 browser presentation where exercise results slides embed live interactive Streamlit `<iframe>` dashboards instead of static screenshots.
- **Template Support:** Optionally inherits slide masters, branding, and logos from an existing `.pptx` template.

## Installation

```bash
pip install python-pptx Pillow lxml pyyaml
```

*(Optional — only needed if using `--capture-live` to capture fresh Streamlit screenshots for the `.pptx`:)*
```bash
pip install playwright && playwright install chromium
```

## Usage

```bash
python github_zip_to_pptx.py hoopavalley.zip \
  -o HoopaValley_Workshop.pptx \
  --template "Scenario Course_NTICCReno_October 2026.pptx"
```

This outputs both `HoopaValley_Workshop.pptx` and `HoopaValley_Workshop.html`.

### Options

| Flag | Description |
| :--- | :--- |
| `-o`, `--output` | Output `.pptx` file path (default: `HoopaValley_Workshop.pptx`) |
| `--html` | Custom output `.html` file path (defaults to matching the `-o` filename) |
| `-t`, `--template` | Reference `.pptx` file to inherit Slide Master, theme, and logos from |
| `--capture-live` | Capture fresh screenshots of embedded Streamlit apps via Playwright for the `.pptx` |
