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

# The first figure in index.md becomes the title slide image

### 1. Basic Conversion (Simplest)
Convert a GitHub Pages / MkDocs `.zip` archive using default settings. This generates `HoopaValley_Workshop.pptx` and `HoopaValley_Workshop.html` in the current directory:

```bash
python github_zip_to_pptx.py hoopavalley.zip
```

### 2. Custom Output Filenames
Specify a custom PowerPoint filename with `-o` (the matching `.html` presentation automatically uses the same base name unless `--html` is also set):

```bash
python github_zip_to_pptx.py hoopavalley.zip -o Workshop_Deck.pptx
```

To save the `.pptx` and `.html` files to separate paths:

```bash
python github_zip_to_pptx.py hoopavalley.zip -o dist/slides.pptx --html dist/interactive.html
```

### 3. Apply a Branded PowerPoint Template
Inherit the Slide Master, color theme, and logos from an existing `.pptx` template:

```bash
python github_zip_to_pptx.py hoopavalley.zip -o Workshop_Deck.pptx \
  --template "Scenario Course_NTICCReno_October 2026.pptx"
```

### 4. Typography, Cover Image & Live Dashboard Capture
Adjust text auto-fitting bounds, override the Slide 1 cover graphic, or capture fresh screenshots of embedded Streamlit dashboards (requires Playwright):

```bash
python github_zip_to_pptx.py hoopavalley.zip -o Workshop_Deck.pptx \
  --template "Scenario Course_NTICCReno_October 2026.pptx" \
  --font-size 13 \
  --min-font-size 10 \
  --cover-image hero_banner.png \
  --capture-live
```

---

## Command-Line Options

| Option | Default | Description |
| :--- | :--- | :--- |
| `zip_file` | *(Required)* | Path to the GitHub.io / MkDocs `.zip` archive (e.g., `hoopavalley.zip`). |
| `-o`, `--output` | `HoopaValley_Workshop.pptx` | Output path for the generated PowerPoint (`.pptx`) deck. |
| `--html` | `<output>.html` | Output path for the self-contained HTML presentation. Defaults to the `-o` filename with an `.html` extension. |
| `-t`, `--template` | `None` | Reference `.pptx` template used to inherit Slide Master layouts, theme colors, and logos. |
| `--font-size` | `12.5` | Preferred body text size (in points) before auto-fit scaling is applied. |
| `--min-font-size` | `9.5` | Minimum body text size (in points) allowed when shrinking dense content to fit on one slide before spilling onto a `(Cont.)` slide. |
| `--cover-image` | First image in `index.md` | Custom image path or filename inside `docs/` to use on the Slide 1 cover. |
| `--capture-live` | `False` | Uses Playwright to render and capture fresh PNG screenshots of live Streamlit `<iframe>` embeds for the `.pptx` slides. |
