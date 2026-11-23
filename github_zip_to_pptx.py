#!/usr/bin/env python3
"""
Convert a GitHub.io / MkDocs Markdown ZIP archive into:
  1. A PowerPoint deck (.pptx) matching Scenario Course_NTICCReno_October 2026.pptx
     with zero empty/master placeholders, deterministic 1-Markdown-page-per-slide
     mapping, full preservation of Markdown/HTML/CSS/LaTeX page breaks, calibrated
     Open Sans auto-fit scaling, proportional tables, adaptive image placement,
     and static dashboard screenshots.
  2. A matching self-contained HTML presentation (.html) with the same visual
     slide chrome, layout rules, print page-break CSS, embedded images, and
     LIVE interactive Streamlit iframes on the results slides.

Usage:
    python github_zip_to_pptx.py hoopavalley.zip -o HoopaValley_Workshop.pptx \
        --template "Scenario Course_NTICCReno_October 2026.pptx"
"""

import argparse
import base64
import html
import math
import os
import re
import shutil
import tempfile
import urllib.parse
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import yaml
except ImportError:
    yaml = None

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


# ---------------------------------------------------------------------------
# Brand Palette & Layout Geometry (matches Scenario Course_NTICCReno_October 2026.pptx)
# ---------------------------------------------------------------------------
COLOR_CRIMSON = RGBColor(0xAB, 0x19, 0x2D)      # #AB192D Top banner & accent bars
COLOR_DARK = RGBColor(0x22, 0x22, 0x22)         # #222222 Primary body text
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)        # #000000 Top-left corner block
COLOR_SIDEBAR = RGBColor(0x7F, 0x7F, 0x7F)      # #7F7F7F Left vertical accent bar
COLOR_CARD_BG = RGBColor(0xE7, 0xE6, 0xE6)      # #E7E6E6 SMS callout & card fill
COLOR_ALT_ROW = RGBColor(0xF3, 0xF3, 0xF3)      # #F3F3F3 Table alternating row fill
COLOR_BORDER = RGBColor(0xD0, 0xD0, 0xD0)       # #D0D0D0 Subtle card/table border
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)        # #FFFFFF White text & slide canvas
COLOR_LINK = RGBColor(0x05, 0x63, 0xC1)         # #0563C1 Hyperlink blue

FONT_FAMILY = "Open Sans"
SLIDE_WIDTH = Inches(10.0)
SLIDE_HEIGHT = Inches(5.625)

CONTENT_TOP_START = 1.00      # Below 0.92" crimson banner
MAX_CONTENT_BOTTOM = 4.96     # Safely above bottom NCDP footer logo
CONTENT_LEFT = 0.52
FULL_CONTENT_WIDTH = 8.96
SPLIT_CONTENT_WIDTH = 5.65
IMAGE_RIGHT_LEFT = 6.32
IMAGE_RIGHT_TOP = 1.12
IMAGE_RIGHT_MAX_W = 3.22
IMAGE_RIGHT_MAX_H = 3.65

# Calibrated Open Sans average character width multiplier (prevents line-wrap underestimation)
OPEN_SANS_CHAR_FACTOR = 0.54


# ---------------------------------------------------------------------------
# ZIP Extraction & MkDocs Discovery
# ---------------------------------------------------------------------------
def extract_zip_and_locate_project(zip_path: str, work_dir: str) -> Tuple[Path, Path, Optional[Path]]:
    """Extract ZIP archive, ignore __MACOSX, and locate project root, docs/, and mkdocs.yml."""
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(work_dir)

    root = Path(work_dir)
    mkdocs_candidates = [
        p for p in root.rglob("mkdocs.yml")
        if "__MACOSX" not in p.parts and ".git" not in p.parts
    ]

    if mkdocs_candidates:
        mkdocs_path = sorted(mkdocs_candidates, key=lambda p: len(p.parts))[0]
        project_root = mkdocs_path.parent
    else:
        mkdocs_path = None
        top_dirs = [
            d for d in root.iterdir()
            if d.is_dir() and d.name not in ("__MACOSX", ".git")
        ]
        project_root = top_dirs[0] if len(top_dirs) == 1 else root

    docs_dir = project_root / "docs" if (project_root / "docs").is_dir() else project_root
    return project_root, docs_dir, mkdocs_path


def load_mkdocs_config(mkdocs_path: Optional[Path], docs_dir: Path) -> Tuple[str, List[Tuple[str, str]]]:
    """
    Read site_name and ordered (nav_label, md_rel_path) entries from mkdocs.yml.
    Preserves exact page ordering so each Markdown file remains its own slide unit.
    """
    site_name = "Two-Way Emergency Communication Workshop"
    nav_entries: List[Tuple[str, str]] = []

    if mkdocs_path and mkdocs_path.exists():
        raw = mkdocs_path.read_text(encoding="utf-8", errors="replace")
        if yaml is not None:
            cleaned_yaml = re.sub(r"![a-zA-Z0-9_.-]+", "", raw)
            cfg = yaml.safe_load(cleaned_yaml) or {}
            site_name = cfg.get("site_name", site_name)

            def flatten_nav(items, parent_label: str = ""):
                for item in items or []:
                    if isinstance(item, str) and item.endswith(".md"):
                        label = parent_label or Path(item).stem.replace("_", " ").title()
                        nav_entries.append((label, item))
                    elif isinstance(item, dict):
                        for key, val in item.items():
                            key_str = str(key).strip()
                            if isinstance(val, str) and val.endswith(".md"):
                                nav_entries.append((key_str, val))
                            elif isinstance(val, list):
                                flatten_nav(val, key_str)

            flatten_nav(cfg.get("nav", []))
        else:
            m_site = re.search(r"^site_name:\s*(.+)$", raw, re.MULTILINE)
            if m_site:
                site_name = m_site.group(1).strip().strip("\"'")
            for m_md in re.finditer(r"(?:([^:\n]+):\s*)?['\"]?([A-Za-z0-9_./-]+\.md)['\"]?", raw):
                lbl = (m_md.group(1) or "").strip(" -'\"")
                md_f = m_md.group(2).strip()
                if not lbl:
                    lbl = Path(md_f).stem.replace("_", " ").title()
                nav_entries.append((lbl, md_f))

    if not nav_entries:
        all_md = sorted(p.name for p in docs_dir.glob("*.md"))
        if "index.md" in all_md:
            all_md.remove("index.md")
            all_md = ["index.md"] + all_md
        nav_entries = [(Path(f).stem.replace("_", " ").title(), f) for f in all_md]

    return site_name, nav_entries


def resolve_doc_image(src: str, docs_dir: Path, md_rel_path: str = "") -> Optional[Path]:
    """Resolve an image src from Markdown/HTML relative to the markdown file or docs_dir."""
    if not src:
        return None
    clean_src = urllib.parse.unquote(src.strip())
    if clean_src.startswith(("http://", "https://", "data:")):
        return None

    rel_norm = clean_src.lstrip("./")
    candidates = []
    if md_rel_path:
        md_parent = (docs_dir / md_rel_path).parent
        candidates.append(md_parent / clean_src)
        candidates.append(md_parent / rel_norm)
    candidates.append(docs_dir / clean_src)
    candidates.append(docs_dir / rel_norm)

    for c in candidates:
        if c.exists() and c.is_file():
            return c

    # Fallback: match filename anywhere under docs_dir (ignoring __MACOSX)
    target_name = Path(rel_norm).name.lower()
    for p in docs_dir.rglob("*"):
        if "__MACOSX" not in p.parts and p.is_file() and p.name.lower() == target_name:
            return p
    return None


# ---------------------------------------------------------------------------
# Template & Slide Chrome Helpers (Zero Empty/Master Placeholders)
# ---------------------------------------------------------------------------
def _strip_prompt_placeholders_from_container(container):
    """Remove any 'Click to edit...' or 'Click to add...' placeholder shapes from a master/layout."""
    for shp in list(container.shapes):
        has_prompt = False
        if shp.has_text_frame:
            txt = shp.text_frame.text.strip().lower()
            if "click to edit" in txt or "click to add" in txt:
                has_prompt = True
        if has_prompt:
            sp = shp._element
            sp.getparent().remove(sp)


def initialize_presentation(template_path: Optional[str]) -> Tuple[Presentation, bool]:
    """Load template PPTX (clearing sample slides & master prompts while keeping theme/logos) or create blank 16:9 deck."""
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
        for i in range(len(prs.slides) - 1, -1, -1):
            r_id = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(r_id)
            del prs.slides._sldIdLst[i]
        for master in prs.slide_masters:
            _strip_prompt_placeholders_from_container(master)
            for layout in master.slide_layouts:
                _strip_prompt_placeholders_from_container(layout)
        return prs, True

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs, False


def get_layout(prs: Presentation, preferred_names: List[str]):
    """Find the best matching slide layout by name."""
    for name in preferred_names:
        for layout in prs.slide_layouts:
            if layout.name.lower() == name.lower():
                return layout
    return prs.slide_layouts[-1]


def prune_unused_placeholders(slide, keep_title: bool = False):
    """
    Remove all unused placeholder shapes (and any 'Click to edit/add...' shapes)
    from the slide XML so PowerPoint/Google Slides/LibreOffice never displays
    master placeholder prompts.
    """
    title_shp = None
    if keep_title:
        try:
            title_shp = slide.shapes.title
        except Exception:
            title_shp = None

    for shp in list(slide.shapes):
        if keep_title and title_shp is not None and shp == title_shp:
            continue

        sp = shp._element
        is_ph = shp.is_placeholder or (sp.find(".//" + qn("p:ph")) is not None)
        has_prompt = False
        if shp.has_text_frame:
            txt = shp.text_frame.text.strip().lower()
            if "click to edit" in txt or "click to add" in txt:
                has_prompt = True

        if is_ph or has_prompt:
            sp.getparent().remove(sp)


def apply_standalone_chrome(slide, slide_num: int, is_cover: bool = False):
    """Draw the crimson top header bar and left vertical accent bar when no template is passed."""
    left_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0), Inches(0.22), SLIDE_HEIGHT
    )
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = COLOR_SIDEBAR
    left_bar.line.fill.background()

    corner = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0), Inches(0.22), Inches(0.92)
    )
    corner.fill.solid()
    corner.fill.fore_color.rgb = COLOR_BLACK
    corner.line.fill.background()

    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.22), Inches(0.0), Inches(9.78), Inches(0.92)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = COLOR_CRIMSON
    header.line.fill.background()

    if not is_cover:
        tab = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(8.84), Inches(0.89), Inches(0.73), Inches(0.36)
        )
        tab.fill.solid()
        tab.fill.fore_color.rgb = COLOR_CRIMSON
        tab.line.fill.background()
        tf = tab.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(slide_num)
        r.font.name = FONT_FAMILY
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE


def create_clean_slide(prs: Presentation, using_template: bool, slide_num: int, is_cover: bool = False):
    """Create a slide with all default 'Click to edit/add...' placeholders stripped."""
    if is_cover:
        layout = get_layout(prs, ["Blank", "Title Only", "TITLE_ONLY", "Title Slide"])
        slide = prs.slides.add_slide(layout)
        prune_unused_placeholders(slide, keep_title=False)
        if not using_template:
            apply_standalone_chrome(slide, slide_num, is_cover=True)
        return slide

    layout = get_layout(prs, ["Title and Content", "Title Only", "TITLE_ONLY", "Blank"])
    slide = prs.slides.add_slide(layout)
    has_title_ph = using_template and (slide.shapes.title is not None)
    prune_unused_placeholders(slide, keep_title=has_title_ph)
    if not using_template:
        apply_standalone_chrome(slide, slide_num, is_cover=False)
    return slide


def set_slide_title(slide, title_text: str, using_template: bool, font_size_pt: float = 20.5):
    """Populate the template title placeholder or create a clean header title box."""
    clean_title = strip_md_inline(title_text)
    if len(clean_title) > 78 and font_size_pt > 18.0:
        font_size_pt = 18.0

    if using_template and slide.shapes.title is not None:
        tf = slide.shapes.title.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = clean_title
        run.font.name = FONT_FAMILY
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = COLOR_WHITE
        return

    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.16), Inches(8.35), Inches(0.62))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = clean_title
    run.font.name = FONT_FAMILY
    run.font.size = Pt(font_size_pt)
    run.font.bold = False
    run.font.color.rgb = COLOR_WHITE


# ---------------------------------------------------------------------------
# Inline Markdown Formatting & Calibrated Height Estimation
# ---------------------------------------------------------------------------
INLINE_TOKEN_RE = re.compile(
    r"(\*\*[^*]+\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\)|\*[^*]+\*)"
)


def strip_md_inline(text: str) -> str:
    """Remove markdown symbols for plain-text contexts and width measurements."""
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = text.replace("**", "").replace("`", "").replace("*", "")
    return re.sub(r"\s+", " ", text).strip()


def md_inline_to_html(text: str) -> str:
    """Convert **bold**, *italic*, `code`, and [label](url) to safe HTML."""
    tokens = INLINE_TOKEN_RE.split(text)
    out = []
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            out.append(f"<strong>{html.escape(tok[2:-2])}</strong>")
        elif tok.startswith("`") and tok.endswith("`"):
            out.append(f'<code class="inline-code">{html.escape(tok[1:-1])}</code>')
        elif tok.startswith("[") and "](" in tok and tok.endswith(")"):
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", tok)
            if m:
                lbl, url = html.escape(m.group(1)), html.escape(m.group(2), quote=True)
                out.append(f'<a href="{url}" target="_blank" rel="noopener">{lbl}</a>')
        elif tok.startswith("*") and tok.endswith("*"):
            out.append(f"<em>{html.escape(tok[1:-1])}</em>")
        else:
            out.append(html.escape(tok.replace("`", "")))
    return "".join(out)


def add_formatted_runs(
    paragraph,
    text: str,
    base_size_pt: float = 12.5,
    base_color: RGBColor = COLOR_DARK,
    base_bold: bool = False,
    base_italic: bool = False,
):
    """Render **bold**, *italic*, `code`, and [text](url) hyperlinks into pptx runs."""
    tokens = INLINE_TOKEN_RE.split(text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            r = paragraph.add_run()
            r.text = tok[2:-2]
            r.font.name = FONT_FAMILY
            r.font.size = Pt(base_size_pt)
            r.font.bold = True
            r.font.italic = base_italic
            r.font.color.rgb = base_color
        elif tok.startswith("`") and tok.endswith("`"):
            r = paragraph.add_run()
            r.text = tok[1:-1]
            r.font.name = FONT_FAMILY
            r.font.size = Pt(base_size_pt)
            r.font.bold = True
            r.font.color.rgb = COLOR_CRIMSON
        elif tok.startswith("[") and "](" in tok and tok.endswith(")"):
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", tok)
            if m:
                label, url = m.group(1), m.group(2)
                r = paragraph.add_run()
                r.text = label
                r.font.name = FONT_FAMILY
                r.font.size = Pt(base_size_pt)
                r.font.bold = True
                r.font.color.rgb = COLOR_LINK
                r.hyperlink.address = url
        elif tok.startswith("*") and tok.endswith("*"):
            r = paragraph.add_run()
            r.text = tok[1:-1]
            r.font.name = FONT_FAMILY
            r.font.size = Pt(base_size_pt)
            r.font.bold = base_bold
            r.font.italic = True
            r.font.color.rgb = base_color
        else:
            r = paragraph.add_run()
            r.text = tok.replace("`", "")
            r.font.name = FONT_FAMILY
            r.font.size = Pt(base_size_pt)
            r.font.bold = base_bold
            r.font.italic = base_italic
            r.font.color.rgb = base_color


def set_bullet_xml(paragraph, level: int = 0):
    """Enable native PowerPoint bullet formatting on a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.26 + level * 0.24))))
    pPr.set("indent", str(int(Inches(-0.16))))
    buChar = pPr.find(qn("a:buChar"))
    if buChar is None:
        buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "•" if level == 0 else "▪")


def estimate_text_height_in(
    text: str, width_in: float, font_pt: float, is_bullet: bool = False, compact: bool = False
) -> float:
    """Estimate the rendered height (in inches) of a wrapped paragraph or bullet using Open Sans metrics."""
    plain = strip_md_inline(text)
    if not plain:
        return 0.16
    effective_w = max(1.2, width_in - (0.30 if is_bullet else 0.10))
    chars_per_line = max(10, int((effective_w * 72.0) / (font_pt * OPEN_SANS_CHAR_FACTOR)))
    effective_len = math.ceil(len(plain) * 1.06)
    lines = max(1, math.ceil(effective_len / chars_per_line))
    line_mult = 1.18 if compact else 1.24
    line_height_in = (font_pt * line_mult) / 72.0
    extra_pad = (0.03 if compact else 0.05) if is_bullet else (0.03 if compact else 0.04)
    return lines * line_height_in + extra_pad


def compute_table_col_widths(rows: List[List[str]], total_width_in: float) -> List[float]:
    """Allocate table column widths proportionally to content length so dense columns don't over-wrap."""
    if not rows:
        return []
    n_cols = max(len(r) for r in rows)
    if n_cols == 1:
        return [total_width_in]

    col_weights = []
    for c in range(n_cols):
        lengths = []
        for r_idx, row in enumerate(rows):
            txt = strip_md_inline(row[c]) if c < len(row) else ""
            lengths.append(len(txt) * (0.8 if r_idx == 0 else 1.0))
        max_len = max(lengths) if lengths else 8
        avg_len = (sum(lengths) / len(lengths)) if lengths else 8
        weight = max(6.0, 0.65 * max_len + 0.35 * avg_len)
        col_weights.append(weight)

    min_frac = 0.14 if n_cols <= 4 else 0.10
    max_frac = 0.60
    total_w = sum(col_weights)
    fracs = [max(min_frac, min(max_frac, w / total_w)) for w in col_weights]
    frac_sum = sum(fracs)
    return [total_width_in * (f / frac_sum) for f in fracs]


def estimate_table_height_in(
    rows: List[List[str]], width_in: float, body_font_pt: float = 12.5, compact: bool = False
) -> float:
    """Estimate total table height accounting for proportional column widths and wrapped cells."""
    if not rows:
        return 0.40
    col_widths = compute_table_col_widths(rows, width_in)
    n_cols = len(col_widths)
    total_h = 0.0
    header_pt = max(9.0, body_font_pt - 0.5)
    cell_pt = max(8.5, body_font_pt - 1.0)
    line_mult = 1.16 if compact else 1.22
    cell_vpad = 0.06 if compact else 0.08
    min_row_h = 0.25 if compact else 0.30

    for r_idx, row in enumerate(rows):
        font_pt = header_pt if r_idx == 0 else cell_pt
        max_lines = 1
        for c_idx in range(n_cols):
            cell_txt = row[c_idx] if c_idx < len(row) else ""
            plain = strip_md_inline(cell_txt)
            col_w = col_widths[c_idx]
            chars_per_line = max(6, int(((col_w - 0.16) * 72.0) / (font_pt * OPEN_SANS_CHAR_FACTOR)))
            effective_len = math.ceil(len(plain) * 1.06)
            lines = max(1, math.ceil(effective_len / chars_per_line))
            max_lines = max(max_lines, lines)
        row_h = max(min_row_h, max_lines * ((font_pt * line_mult) / 72.0) + cell_vpad)
        total_h += row_h
    return total_h


def estimate_blocks_total_height(
    blocks: List[Dict], width_in: float, body_font_pt: float, compact: bool = False
) -> float:
    """Compute the total vertical height required by a list of blocks."""
    sub_font_pt = max(8.5, body_font_pt - 1.0)
    total = 0.0
    for idx, b in enumerate(blocks):
        if b["type"] == "table":
            h = estimate_table_height_in(b["rows"], width_in, body_font_pt, compact=compact)
            gap = 0.06 if compact else 0.09
        elif b["type"] == "callout":
            h = estimate_text_height_in(
                b["text"], width_in - 0.28, body_font_pt + 0.5, is_bullet=False, compact=compact
            ) + 0.12
            gap = 0.05 if compact else 0.08
        elif b["type"] == "subheading":
            h = estimate_text_height_in(
                b["text"], width_in, body_font_pt + 1.5, is_bullet=False, compact=compact
            ) + (0.04 if idx > 0 else 0.0)
            gap = 0.03 if compact else 0.05
        elif b["type"] == "paragraph":
            h = estimate_text_height_in(b["text"], width_in, body_font_pt, is_bullet=False, compact=compact)
            gap = 0.04 if compact else 0.06
        elif b["type"] == "bullets":
            h = sum(
                estimate_text_height_in(
                    btxt, width_in, body_font_pt if lvl == 0 else sub_font_pt, is_bullet=True, compact=compact
                )
                for lvl, btxt in b["items"]
            )
            gap = 0.04 if compact else 0.06
        else:
            continue
        total += h + (gap if idx < len(blocks) - 1 else 0.0)
    return total


# ---------------------------------------------------------------------------
# HTML <img>, Page Break & Markdown Parser (1 Markdown Page = 1 Slide Unit)
# ---------------------------------------------------------------------------
IMG_TAG_RE = re.compile(r"<img\b([^>]*)/?>", re.IGNORECASE | re.DOTALL)

# Match HTML/CSS page breaks (e.g. <div style="page-break-after: always;"></div>, break-before: page)
PAGE_BREAK_HTML_RE = re.compile(
    r"""<(?:div|p|span|section|br|hr)\b[^>]*(?:page-break-(?:after|before)\s*:\s*always|break-(?:after|before)\s*:\s*page)[^>]*(?:/>|>\s*</(?:div|p|span|section)>|>)""",
    re.IGNORECASE | re.DOTALL,
)
# Match class-based page/slide break elements (e.g. <div class="page-break"></div>, <hr class="pagebreak">)
PAGE_BREAK_CLASS_RE = re.compile(
    r"""<(?:div|p|span|hr|br)\b[^>]*\bclass\s*=\s*["'][^"']*\b(?:page-?break|new-?page|slide-?break)\b[^"']*["'][^>]*(?:/>|>\s*</(?:div|p|span)>|>)""",
    re.IGNORECASE | re.DOTALL,
)
# Match HTML comment page/slide breaks (e.g. <!-- pagebreak -->, <!-- page-break -->, <!-- slide -->)
PAGE_BREAK_COMMENT_RE = re.compile(
    r"""<!--\s*(?:page-?break|new-?page|slide(?:-?break)?)\s*-->""",
    re.IGNORECASE,
)
# Match LaTeX / Pandoc / Quarto page break commands (\newpage, \pagebreak)
PAGE_BREAK_LATEX_RE = re.compile(
    r"""^\s*\\(?:newpage|pagebreak)\s*$""",
    re.MULTILINE,
)
# Match standalone HTML <hr> tags or Markdown thematic breaks (---, ***, ___, - - -)
PAGE_BREAK_HR_TAG_RE = re.compile(
    r"""<hr\b[^>]*/?>""",
    re.IGNORECASE,
)
PAGE_BREAK_MD_HR_RE = re.compile(
    r"""^\s*(?:-\s*-\s*-[-\s]*|\*\s*\*\s*\*[*\s]*|_\s*_\s*_[_\s]*)$""",
    re.MULTILINE,
)


def normalize_page_breaks_to_markers(md_text: str) -> str:
    """
    Strip top-of-file YAML frontmatter if present, then convert all Markdown,
    HTML/CSS (`page-break-after: always`, `break-before: page`, `.page-break`),
    comment (`<!-- pagebreak -->`), LaTeX (`\\newpage`, `\\pagebreak`), and
    form-feed (`\\f`) page breaks into standalone `[[PAGE_BREAK]]` lines before
    HTML tag cleanup runs.
    """
    # Strip YAML frontmatter only when it opens at the very first line of the file
    text = re.sub(r"\A---\s*\n.*?\n---\s*(?:\n|$)", "", md_text, count=1, flags=re.DOTALL)
    text = text.replace("\x0c", "\n[[PAGE_BREAK]]\n")
    text = PAGE_BREAK_HTML_RE.sub("\n[[PAGE_BREAK]]\n", text)
    text = PAGE_BREAK_CLASS_RE.sub("\n[[PAGE_BREAK]]\n", text)
    text = PAGE_BREAK_COMMENT_RE.sub("\n[[PAGE_BREAK]]\n", text)
    text = PAGE_BREAK_LATEX_RE.sub("\n[[PAGE_BREAK]]\n", text)
    text = PAGE_BREAK_HR_TAG_RE.sub("\n[[PAGE_BREAK]]\n", text)
    text = PAGE_BREAK_MD_HR_RE.sub("\n[[PAGE_BREAK]]\n", text)
    return text


def parse_html_img_tag(tag_attrs: str) -> Optional[Dict]:
    """Extract src, alt, and optional width from an HTML <img> tag's attribute string."""
    src_m = re.search(r"""\bsrc\s*=\s*(?:["']([^"']+)["']|([^\s>]+))""", tag_attrs, re.IGNORECASE)
    if not src_m:
        return None
    src = (src_m.group(1) or src_m.group(2) or "").strip()
    if not src:
        return None

    alt_m = re.search(r"""\balt\s*=\s*(?:["']([^"']*)["']|([^\s>]+))""", tag_attrs, re.IGNORECASE)
    alt = (alt_m.group(1) if alt_m and alt_m.group(1) is not None else (alt_m.group(2) if alt_m else "")).strip()

    width_px = None
    style_w = re.search(r"""width\s*:\s*(\d+)\s*px""", tag_attrs, re.IGNORECASE)
    attr_w = re.search(r"""\bwidth\s*=\s*["']?(\d+)""", tag_attrs, re.IGNORECASE)
    if style_w:
        width_px = int(style_w.group(1))
    elif attr_w:
        width_px = int(attr_w.group(1))

    return {"type": "image", "alt": alt, "src": src, "width_px": width_px}


def normalize_html_images_to_markers(md_text: str) -> str:
    """
    Convert any <img ...> HTML tags into standalone `[[IMG_TOKEN:...]]` lines
    before stripping layout HTML tags, so tags are never lost or shown raw.
    """
    def _repl(match: re.Match) -> str:
        parsed = parse_html_img_tag(match.group(1))
        if not parsed:
            return ""
        src_q = urllib.parse.quote(parsed["src"], safe="/:._-")
        alt_q = urllib.parse.quote(parsed["alt"], safe="")
        w_str = str(parsed["width_px"]) if parsed["width_px"] else ""
        return f"\n[[IMG_TOKEN|{src_q}|{alt_q}|{w_str}]]\n"

    return IMG_TAG_RE.sub(_repl, md_text)


def _clean_inline_html(text: str) -> str:
    """Convert inline HTML formatting tags to Markdown and strip stray container tags."""
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<(?:strong|b)\b[^>]*>(.*?)</(?:strong|b)>", r"**\1**", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<(?:em|i)\b[^>]*>(.*?)</(?:em|i)>", r"*\1*", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text).strip()


def parse_markdown_sections(md_text: str, fallback_title: str = "Workshop Slide") -> Dict:
    """
    Parse a single MkDocs Markdown file into structured slide units:
    - Keeps the entire Markdown page together as ONE slide section by default so
      multiple sections/H2s on the same page stay together (and separate .md files
      never merge).
    - Splits into multiple slides whenever the Markdown writer inserts an explicit
      page break (`---`, `***`, `___`, `<div style="page-break-after: always;"></div>`,
      `<!-- pagebreak -->`, `\\newpage`, `\\pagebreak`, etc.).
    """
    iframe_match = re.search(r'<iframe[^>]+src=["\']([^"\']+)["\']', md_text, re.IGNORECASE)
    iframe_url = iframe_match.group(1).replace("&amp;", "&") if iframe_match else None

    # Normalize explicit page breaks & HTML <img> tags BEFORE stripping HTML layout tags
    cleaned = normalize_page_breaks_to_markers(md_text)
    cleaned = normalize_html_images_to_markers(cleaned)
    cleaned = re.sub(r"<style>.*?</style>", "", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<h1\b[^>]*>(.*?)</h1>", lambda m: f"\n# {_clean_inline_html(m.group(1))}\n", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<h2\b[^>]*>(.*?)</h2>", lambda m: f"\n## {_clean_inline_html(m.group(1))}\n", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<h3\b[^>]*>(.*?)</h3>", lambda m: f"\n### {_clean_inline_html(m.group(1))}\n", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<div[^>]*>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</div>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"<p[^>]*>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</p>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"<iframe[^>]*>.*?</iframe>", "", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"^\s*Loading live results(?:…|\.\.\.)\s*$", "", cleaned, flags=re.MULTILINE)

    lines = cleaned.splitlines()
    page_title = fallback_title
    sections = []
    current = {"title": None, "blocks": []}

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # Explicit page break in the Markdown file -> enforce a new slide
        if stripped == "[[PAGE_BREAK]]":
            if current["blocks"]:
                if current["title"] is None:
                    current["title"] = page_title
                sections.append(current)
                current = {"title": None, "blocks": []}
            i += 1
            continue

        # Top-level H1 title
        if stripped.startswith("# ") and not stripped.startswith("## "):
            h1_clean = strip_md_inline(_clean_inline_html(stripped[2:]))
            if h1_clean:
                page_title = h1_clean
                if current["title"] is None:
                    current["title"] = page_title
            i += 1
            continue

        # H2 / H3 headings:
        # - If this is the very first line on a new slide right after a page break,
        #   use it as that slide's title; otherwise render it as an in-slide subheading.
        if stripped.startswith("## ") or stripped.startswith("### "):
            raw_h = stripped.lstrip("#").strip()
            h_clean = _clean_inline_html(raw_h)
            if h_clean:
                if re.search(r"by text or WhatsApp to", h_clean, re.IGNORECASE):
                    current["blocks"].append({"type": "callout", "text": h_clean})
                elif current["title"] is None and not current["blocks"] and len(sections) > 0:
                    current["title"] = strip_md_inline(h_clean)
                else:
                    current["blocks"].append({"type": "subheading", "text": strip_md_inline(h_clean)})
            i += 1
            continue

        # Normalized HTML <img> token
        tok_m = re.match(r"^\[\[IMG_TOKEN\|([^|]*)\|([^|]*)\|([^]]*)\]\]$", stripped)
        if tok_m:
            src_val = urllib.parse.unquote(tok_m.group(1))
            alt_val = urllib.parse.unquote(tok_m.group(2))
            w_val = int(tok_m.group(3)) if tok_m.group(3).isdigit() else None
            current["blocks"].append({"type": "image", "alt": alt_val, "src": src_val, "width_px": w_val})
            i += 1
            continue

        # Standard Markdown image ![alt](src) or ![alt](src "title")
        img_m = re.match(r"""^!\[([^\]]*)\]\(([^)\s]+)(?:\s+["'][^"']*["'])?\)$""", stripped)
        if img_m:
            current["blocks"].append({
                "type": "image",
                "alt": img_m.group(1),
                "src": img_m.group(2).strip(),
                "width_px": None,
            })
            i += 1
            continue

        # Markdown pipe table
        if stripped.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:|-]+\|$", lines[i + 1].strip()):
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row_str = lines[i].strip()
                if not re.match(r"^\|[\s:|-]+\|$", row_str):
                    cells = [_clean_inline_html(c) for c in row_str.strip("|").split("|")]
                    table_rows.append(cells)
                i += 1
            if table_rows:
                current["blocks"].append({"type": "table", "rows": table_rows})
            continue

        # Bullet list (skip empty list items so blank bullets never render)
        if re.match(r"^(\s*)[-*+]\s+", line):
            bullets = []
            while i < len(lines) and re.match(r"^(\s*)[-*+]\s*(.*)$", lines[i]):
                m = re.match(r"^(\s*)[-*+]\s*(.*)$", lines[i])
                indent_spaces = len(m.group(1))
                lvl = 1 if indent_spaces >= 2 else 0
                item_txt = _clean_inline_html(m.group(2))
                if item_txt:
                    bullets.append((lvl, item_txt))
                i += 1
            if bullets:
                current["blocks"].append({"type": "bullets", "items": bullets})
            continue

        # Standard paragraph
        para_lines = [_clean_inline_html(stripped)]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt == "[[PAGE_BREAK]]"
                or nxt.startswith("#")
                or nxt.startswith("|")
                or nxt.startswith("[[IMG_TOKEN|")
                or re.match(r"^[-*+]\s+", nxt)
                or re.match(r"^!\[", nxt)
            ):
                break
            cleaned_nxt = _clean_inline_html(nxt)
            if cleaned_nxt:
                para_lines.append(cleaned_nxt)
            i += 1
        joined_para = " ".join(p for p in para_lines if p).strip()
        if joined_para:
            current["blocks"].append({"type": "paragraph", "text": joined_para})

    if current["blocks"] or not sections:
        if current["title"] is None:
            current["title"] = page_title
        sections.append(current)

    return {
        "page_title": page_title,
        "iframe_url": iframe_url,
        "sections": sections,
    }


# ---------------------------------------------------------------------------
# Slide Builders (PPTX + Structured Slide Models for HTML Output)
# ---------------------------------------------------------------------------
def add_picture_fit(
    slide,
    img_path: Path,
    left: float,
    top: float,
    max_w: float,
    max_h: float,
    width_px_hint: Optional[int] = None,
    respect_width_hint: bool = True,
):
    """Insert an image scaled proportionally within (max_w, max_h) and centered in the box."""
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    aspect = w_px / float(h_px)

    effective_max_w = max_w
    if respect_width_hint and width_px_hint is not None and width_px_hint > 0:
        effective_max_w = min(max_w, max(1.8, width_px_hint / 96.0))

    box_aspect = effective_max_w / float(max_h)
    if aspect >= box_aspect:
        final_w = effective_max_w
        final_h = effective_max_w / aspect
    else:
        final_h = max_h
        final_w = max_h * aspect

    offset_x = left + (max_w - final_w) / 2.0
    offset_y = top + (max_h - final_h) / 2.0
    return slide.shapes.add_picture(str(img_path), Inches(offset_x), Inches(offset_y), Inches(final_w), Inches(final_h))


def build_cover_slide(
    prs: Presentation,
    using_template: bool,
    site_name: str,
    subtitle: str,
    hero_img: Optional[Path],
    html_slides: List[Dict],
    hero_width_px: Optional[int] = None,
):
    """Build Slide 1 cover slide for both PPTX and HTML without any Master placeholder prompts."""
    slide = create_clean_slide(prs, using_template, 1, is_cover=True)

    if hero_img and hero_img.exists():
        add_picture_fit(
            slide,
            hero_img,
            left=0.70,
            top=1.25,
            max_w=3.85,
            max_h=3.60,
            width_px_hint=hero_width_px,
            respect_width_hint=True,
        )
        text_left, text_w = 4.70, 4.90
    else:
        text_left, text_w = 1.00, 8.00

    title_box = slide.shapes.add_textbox(Inches(text_left), Inches(1.80), Inches(text_w), Inches(2.50))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p_title = tf.paragraphs[0]
    p_title.alignment = PP_ALIGN.CENTER
    r_title = p_title.add_run()
    r_title.text = site_name
    r_title.font.name = FONT_FAMILY
    r_title.font.size = Pt(24)
    r_title.font.bold = False
    r_title.font.color.rgb = COLOR_DARK

    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.alignment = PP_ALIGN.CENTER
        p_sub.space_before = Pt(14)
        add_formatted_runs(p_sub, subtitle, base_size_pt=14, base_color=COLOR_CRIMSON, base_bold=True)

    html_slides.append({
        "type": "cover",
        "num": 1,
        "title": site_name,
        "subtitle": subtitle,
        "hero_img": hero_img,
    })


def build_results_dashboard_slide(
    prs: Presentation,
    using_template: bool,
    slide_num: int,
    slide_title: str,
    screenshot_path: Optional[Path],
    iframe_url: str,
    html_slides: List[Dict],
):
    """
    Build a Streamlit results slide:
      - PPTX: embeds captured PNG screenshot + live URL link
      - HTML: embeds LIVE interactive Streamlit <iframe>
    """
    slide = create_clean_slide(prs, using_template, slide_num, is_cover=False)
    set_slide_title(slide, slide_title, using_template)

    if screenshot_path and screenshot_path.exists():
        frame = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.10), Inches(8.70), Inches(3.52)
        )
        frame.fill.solid()
        frame.fill.fore_color.rgb = COLOR_WHITE
        frame.line.color.rgb = COLOR_BORDER
        frame.line.width = Pt(1.0)
        add_picture_fit(slide, screenshot_path, left=0.70, top=1.14, max_w=8.60, max_h=3.44, respect_width_hint=False)
    else:
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.65), Inches(8.30), Inches(2.40)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_ALT_ROW
        card.line.color.rgb = COLOR_CRIMSON
        tf = card.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = "Live Exercise Results Dashboard"
        r.font.name = FONT_FAMILY
        r.font.size = Pt(20)
        r.font.bold = True
        r.font.color.rgb = COLOR_DARK

    url_box = slide.shapes.add_textbox(Inches(3.40), Inches(4.68), Inches(6.10), Inches(0.32))
    utf = url_box.text_frame
    up = utf.paragraphs[0]
    up.alignment = PP_ALIGN.RIGHT
    ur = up.add_run()
    ur.text = f"Open Live Results: {iframe_url}"
    ur.font.name = FONT_FAMILY
    ur.font.size = Pt(10.0)
    ur.font.color.rgb = COLOR_LINK
    ur.hyperlink.address = iframe_url

    html_slides.append({
        "type": "streamlit",
        "num": slide_num,
        "title": strip_md_inline(slide_title),
        "iframe_url": iframe_url,
    })


def render_styled_table(
    slide,
    rows: List[List[str]],
    left: float,
    top: float,
    width: float,
    height: float,
    body_font_pt: float = 12.5,
    compact: bool = False,
):
    """Render a Markdown table into a formatted PowerPoint table with proportional column widths."""
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table

    col_widths = compute_table_col_widths(rows, width)
    for c_idx, cw in enumerate(col_widths):
        table.columns[c_idx].width = Inches(cw)

    header_pt = max(9.0, body_font_pt - 0.5)
    cell_pt = max(8.5, body_font_pt - 1.0)
    v_margin = Inches(0.025 if compact else 0.04)

    for r_idx, row in enumerate(rows):
        is_header = r_idx == 0
        bg_color = COLOR_CRIMSON if is_header else (COLOR_ALT_ROW if r_idx % 2 == 1 else COLOR_WHITE)
        txt_color = COLOR_WHITE if is_header else COLOR_DARK
        font_sz = header_pt if is_header else cell_pt

        for c_idx in range(n_cols):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.07)
            cell.margin_right = Inches(0.07)
            cell.margin_top = v_margin
            cell.margin_bottom = v_margin

            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            val = row[c_idx] if c_idx < len(row) else ""
            add_formatted_runs(p, val, base_size_pt=font_sz, base_color=txt_color, base_bold=is_header)


def choose_best_fit_for_section(
    blocks: List[Dict],
    width_in: float,
    base_font_pt: float,
    min_font_pt: float,
    available_h: Optional[float] = None,
) -> Tuple[float, bool]:
    """
    Choose the largest font size between `base_font_pt` and `min_font_pt` (and compact spacing
    if needed) that allows the entire slide section to fit on ONE slide without overlapping or
    splitting. Only splits onto `(Cont.)` if `min_font_pt` + compact mode still overflows.
    """
    if available_h is None:
        available_h = MAX_CONTENT_BOTTOM - CONTENT_TOP_START

    if estimate_blocks_total_height(blocks, width_in, base_font_pt, compact=False) <= available_h:
        return base_font_pt, False

    if estimate_blocks_total_height(blocks, width_in, base_font_pt, compact=True) <= available_h:
        return base_font_pt, True

    curr = base_font_pt - 0.5
    while curr >= min_font_pt - 1e-6:
        if estimate_blocks_total_height(blocks, width_in, curr, compact=True) <= available_h:
            return round(curr, 2), True
        curr -= 0.5

    return min_font_pt, True


def expand_splittable_blocks(
    blocks: List[Dict], width_in: float, body_font_pt: float, compact: bool = False
) -> List[Dict]:
    """
    Break oversized bullet lists or tables into smaller chunks so a single block
    never overflows a slide vertically, while keeping normal blocks intact.
    """
    max_single_slide_h = MAX_CONTENT_BOTTOM - CONTENT_TOP_START - 0.04
    sub_font_pt = max(8.5, body_font_pt - 1.0)
    expanded: List[Dict] = []

    for b in blocks:
        if b["type"] == "bullets":
            chunk: List[Tuple[int, str]] = []
            chunk_h = 0.0
            for lvl, btext in b["items"]:
                font_pt = body_font_pt if lvl == 0 else sub_font_pt
                item_h = estimate_text_height_in(btext, width_in, font_pt, is_bullet=True, compact=compact)
                if chunk and (chunk_h + item_h > max_single_slide_h):
                    expanded.append({"type": "bullets", "items": chunk})
                    chunk = [(lvl, btext)]
                    chunk_h = item_h
                else:
                    chunk.append((lvl, btext))
                    chunk_h += item_h
            if chunk:
                expanded.append({"type": "bullets", "items": chunk})
        elif b["type"] == "table" and len(b["rows"]) > 2:
            header = b["rows"][0]
            data_rows = b["rows"][1:]
            curr_rows = [header]
            for r in data_rows:
                cand_h = estimate_table_height_in(curr_rows + [r], width_in, body_font_pt, compact=compact)
                if len(curr_rows) > 1 and cand_h > max_single_slide_h:
                    expanded.append({"type": "table", "rows": curr_rows})
                    curr_rows = [header, r]
                else:
                    curr_rows.append(r)
            if len(curr_rows) > 1:
                expanded.append({"type": "table", "rows": curr_rows})
        else:
            expanded.append(b)

    return expanded


def classify_image_layout(blocks: List[Dict], non_images: List[Dict], body_font_pt: float) -> str:
    """
    Determine image placement based on text volume on the slide:
      - If there is only a line or two of text (no tables/subheadings, <= 1 bullet, <= 150 chars,
        and <= 0.65" tall across full width), stack the image ABOVE ("top") or BELOW ("bottom")
        the text according to the Markdown block order.
      - Otherwise (a lot of text/bullets/table), place the image on the RIGHT ("right") with
        text on the left.
    """
    if not non_images:
        return "top"

    has_table = any(b["type"] == "table" for b in non_images)
    has_subheading = any(b["type"] == "subheading" for b in non_images)
    total_bullets = sum(len(b["items"]) for b in non_images if b["type"] == "bullets")

    plain_chunks = []
    for b in non_images:
        if b["type"] in ("paragraph", "callout", "subheading"):
            plain_chunks.append(strip_md_inline(b["text"]))
        elif b["type"] == "bullets":
            plain_chunks.extend(strip_md_inline(txt) for _, txt in b["items"])
    total_chars = sum(len(c) for c in plain_chunks)
    est_full_h = estimate_blocks_total_height(non_images, FULL_CONTENT_WIDTH, body_font_pt, compact=False)

    is_short_caption = (
        not has_table
        and not has_subheading
        and total_bullets <= 1
        and total_chars <= 150
        and est_full_h <= 0.65
    )
    if not is_short_caption:
        return "right"

    first_img_idx = next((i for i, b in enumerate(blocks) if b["type"] == "image"), 0)
    first_txt_idx = next((i for i, b in enumerate(blocks) if b["type"] != "image"), 1)
    return "top" if first_img_idx < first_txt_idx else "bottom"


def build_content_slides(
    prs: Presentation,
    using_template: bool,
    start_slide_num: int,
    slide_title: str,
    blocks: List[Dict],
    docs_dir: Path,
    md_rel_path: str,
    html_slides: List[Dict],
    body_font_pt: float = 12.5,
    min_font_pt: float = 9.5,
) -> int:
    """
    Build one or more content slides from `blocks`, automatically auto-fitting
    font size down to `min_font_pt` first so each slide section stays on a single
    slide with zero text overlap, and placing images either on the right (when
    there is substantial text) or above/below (when there is only a line or two).
    """
    images = [b for b in blocks if b["type"] == "image"]
    non_images = [b for b in blocks if b["type"] != "image"]

    has_image = False
    img_path = None
    img_width_px = None
    if images:
        candidate = resolve_doc_image(images[0]["src"], docs_dir, md_rel_path)
        if candidate is not None:
            has_image = True
            img_path = candidate
            img_width_px = images[0].get("width_px")

    image_layout = classify_image_layout(blocks, non_images, body_font_pt) if has_image else "none"

    if image_layout == "right":
        content_width = SPLIT_CONTENT_WIDTH
        avail_h = MAX_CONTENT_BOTTOM - CONTENT_TOP_START
        chosen_font_pt, compact_mode = choose_best_fit_for_section(
            non_images, content_width, body_font_pt, min_font_pt, available_h=avail_h
        )
    elif image_layout in ("top", "bottom"):
        content_width = FULL_CONTENT_WIDTH
        chosen_font_pt = max(body_font_pt, 14.0)
        compact_mode = False
    else:
        content_width = FULL_CONTENT_WIDTH
        avail_h = MAX_CONTENT_BOTTOM - CONTENT_TOP_START
        chosen_font_pt, compact_mode = choose_best_fit_for_section(
            non_images, content_width, body_font_pt, min_font_pt, available_h=avail_h
        )

    sub_font_pt = max(8.5, chosen_font_pt - 1.0)
    prepared_blocks = expand_splittable_blocks(non_images, content_width, chosen_font_pt, compact=compact_mode)

    slide_num = start_slide_num
    part_idx = 1

    # Compute exact stacked geometry when image is placed above or below 1-2 lines of text
    stacked_text_top = CONTENT_TOP_START
    stacked_img_top = 1.06
    stacked_img_max_h = 3.15
    if image_layout in ("top", "bottom") and non_images:
        txt_h = max(0.45, estimate_blocks_total_height(prepared_blocks, content_width, chosen_font_pt, compact=False) + 0.08)
        if image_layout == "top":
            stacked_img_top = 1.05
            stacked_img_max_h = max(2.40, MAX_CONTENT_BOTTOM - stacked_img_top - txt_h - 0.14)
            stacked_text_top = stacked_img_top + stacked_img_max_h + 0.10
        else:  # "bottom"
            stacked_text_top = 1.04
            stacked_img_top = stacked_text_top + txt_h + 0.10
            stacked_img_max_h = max(2.40, MAX_CONTENT_BOTTOM - stacked_img_top)

    def open_slide(part_number: int, include_image: bool):
        title_for_slide = slide_title if part_number == 1 else f"{strip_md_inline(slide_title)} (Cont.)"
        s = create_clean_slide(prs, using_template, slide_num, is_cover=False)
        set_slide_title(s, title_for_slide, using_template)
        cur_layout = image_layout if include_image else "none"

        if include_image and img_path is not None:
            if cur_layout == "right":
                add_picture_fit(
                    s,
                    img_path,
                    left=IMAGE_RIGHT_LEFT,
                    top=IMAGE_RIGHT_TOP,
                    max_w=IMAGE_RIGHT_MAX_W,
                    max_h=IMAGE_RIGHT_MAX_H,
                    width_px_hint=img_width_px,
                    respect_width_hint=True,
                )
            elif cur_layout in ("top", "bottom"):
                add_picture_fit(
                    s,
                    img_path,
                    left=CONTENT_LEFT,
                    top=stacked_img_top,
                    max_w=FULL_CONTENT_WIDTH,
                    max_h=stacked_img_max_h,
                    width_px_hint=img_width_px,
                    respect_width_hint=False,
                )

        html_rec = {
            "type": "content",
            "num": slide_num,
            "title": strip_md_inline(title_for_slide),
            "blocks": [],
            "image": img_path if include_image else None,
            "image_layout": cur_layout,
            "font_pt": chosen_font_pt,
        }
        html_slides.append(html_rec)
        return s, html_rec

    slide, current_html = open_slide(part_idx, has_image)
    cursor_top = stacked_text_top if image_layout in ("top", "bottom") else CONTENT_TOP_START
    center_stacked_text = image_layout in ("top", "bottom")

    for b_idx, block in enumerate(prepared_blocks):
        if block["type"] == "table":
            est_h = estimate_table_height_in(block["rows"], content_width, chosen_font_pt, compact=compact_mode)
            gap = 0.06 if compact_mode else 0.09
        elif block["type"] == "callout":
            est_h = estimate_text_height_in(
                block["text"], content_width - 0.28, chosen_font_pt + 0.5, is_bullet=False, compact=compact_mode
            ) + 0.12
            gap = 0.05 if compact_mode else 0.08
        elif block["type"] == "subheading":
            top_pad = 0.05 if (b_idx > 0 and cursor_top > CONTENT_TOP_START + 0.05) else 0.0
            cursor_top += top_pad
            est_h = estimate_text_height_in(
                block["text"], content_width, chosen_font_pt + 1.5, is_bullet=False, compact=compact_mode
            )
            gap = 0.03 if compact_mode else 0.05
        elif block["type"] == "paragraph":
            est_h = estimate_text_height_in(
                block["text"], content_width, chosen_font_pt, is_bullet=False, compact=compact_mode
            )
            gap = 0.04 if compact_mode else 0.06
        elif block["type"] == "bullets":
            est_h = sum(
                estimate_text_height_in(
                    btxt,
                    content_width,
                    chosen_font_pt if lvl == 0 else sub_font_pt,
                    is_bullet=True,
                    compact=compact_mode,
                )
                for lvl, btxt in block["items"]
            )
            gap = 0.04 if compact_mode else 0.06
        else:
            continue

        # Start a continuation slide ONLY if this block genuinely overflows the bottom margin
        if cursor_top > CONTENT_TOP_START + 0.05 and (cursor_top + est_h > MAX_CONTENT_BOTTOM):
            slide_num += 1
            part_idx += 1
            slide, current_html = open_slide(part_idx, include_image=False)
            cursor_top = CONTENT_TOP_START
            center_stacked_text = False

        render_h = min(est_h, MAX_CONTENT_BOTTOM - cursor_top)

        if block["type"] == "table":
            render_styled_table(
                slide, block["rows"], CONTENT_LEFT, cursor_top, content_width, render_h, chosen_font_pt, compact=compact_mode
            )
        elif block["type"] == "callout":
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h)
            )
            card.adjustments[0] = 0.08
            card.fill.solid()
            card.fill.fore_color.rgb = COLOR_CARD_BG
            card.line.fill.background()
            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.14)
            tf.margin_right = Inches(0.14)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            add_formatted_runs(p, block["text"], base_size_pt=chosen_font_pt + 0.5, base_color=COLOR_DARK)
        elif block["type"] == "subheading":
            tb = slide.shapes.add_textbox(Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            add_formatted_runs(
                p, block["text"], base_size_pt=chosen_font_pt + 1.5, base_color=COLOR_CRIMSON, base_bold=True
            )
        elif block["type"] == "paragraph":
            tb = slide.shapes.add_textbox(Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if center_stacked_text else PP_ALIGN.LEFT
            add_formatted_runs(p, block["text"], base_size_pt=chosen_font_pt, base_color=COLOR_DARK)
        elif block["type"] == "bullets":
            tb = slide.shapes.add_textbox(Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            for idx_b, (lvl, btext) in enumerate(block["items"]):
                p = tf.paragraphs[0] if idx_b == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(2.5 if compact_mode else 4.0)
                set_bullet_xml(p, level=lvl)
                add_formatted_runs(
                    p, btext, base_size_pt=chosen_font_pt if lvl == 0 else sub_font_pt, base_color=COLOR_DARK
                )

        current_html["blocks"].append(block)
        cursor_top += render_h + gap

    return slide_num + 1


# ---------------------------------------------------------------------------
# HTML Presentation Builder (Visual Match + Live Streamlit Iframes + Print CSS)
# ---------------------------------------------------------------------------
def image_to_data_uri(img_path: Optional[Path]) -> str:
    """Encode local PNG/JPG into a self-contained base64 data URI for the HTML deck."""
    if not img_path or not img_path.exists():
        return ""
    ext = img_path.suffix.lower().lstrip(".") or "png"
    mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
    b64 = base64.b64encode(img_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def render_html_deck(html_slides: List[Dict], site_name: str, output_html_path: Path, body_font_pt: float = 12.5):
    """Generate a single-file 16:9 HTML slide deck with matching layouts, print page breaks, and live Streamlit iframes."""
    total = len(html_slides)
    rendered_cards: List[str] = []
    default_scale = body_font_pt / 12.5
    body_vw = round(1.22 * default_scale, 3)
    body_max_rem = round(1.04 * default_scale, 3)

    for idx, s in enumerate(html_slides):
        active_cls = " active" if idx == 0 else ""
        s_num = s["num"]
        s_type = s["type"]

        if s_type == "cover":
            hero_uri = image_to_data_uri(s.get("hero_img"))
            hero_html = (
                f'<div class="cover-hero"><img src="{hero_uri}" alt="Cover graphic"></div>'
                if hero_uri else ""
            )
            sub_html = (
                f'<div class="cover-subtitle">{md_inline_to_html(s["subtitle"])}</div>'
                if s.get("subtitle") else ""
            )
            body_html = f"""
            <div class="cover-layout {'has-hero' if hero_uri else 'no-hero'}">
              {hero_html}
              <div class="cover-text">
                <h1>{html.escape(s['title'])}</h1>
                {sub_html}
              </div>
            </div>
            """
            chrome_extra = ""
        else:
            chrome_extra = f"""
            <div class="slide-title-bar">{html.escape(s['title'])}</div>
            <div class="slide-num-tab">{s_num}</div>
            """

            if s_type == "streamlit":
                url_esc = html.escape(s["iframe_url"], quote=True)
                body_html = f"""
                <div class="slide-body streamlit-body">
                  <div class="live-frame">
                    <div class="live-loading">
                      <div class="live-spinner"></div>
                      <div>Loading live Streamlit results&hellip;</div>
                    </div>
                    <iframe
                      src="{url_esc}"
                      loading="lazy"
                      allow="clipboard-read; clipboard-write"
                      onload="var o=this.previousElementSibling; setTimeout(function(){{ o.classList.add('done'); }}, 1500);">
                    </iframe>
                  </div>
                  <div class="streamlit-footer">
                    <a href="{url_esc}" target="_blank" rel="noopener">Open Live Results in New Tab: {html.escape(s['iframe_url'])}</a>
                  </div>
                </div>
                """
            else:
                slide_font_pt = s.get("font_pt", body_font_pt)
                slide_scale = round(slide_font_pt / body_font_pt, 3) if body_font_pt else 1.0
                scale_style = f' style="--slide-scale: {slide_scale};"' if abs(slide_scale - 1.0) > 0.01 else ""

                blocks_html: List[str] = []
                for b in s.get("blocks", []):
                    if b["type"] == "subheading":
                        blocks_html.append(f'<h3 class="body-subheading">{md_inline_to_html(b["text"])}</h3>')
                    elif b["type"] == "paragraph":
                        blocks_html.append(f'<p class="body-para">{md_inline_to_html(b["text"])}</p>')
                    elif b["type"] == "callout":
                        blocks_html.append(f'<div class="inline-callout">{md_inline_to_html(b["text"])}</div>')
                    elif b["type"] == "bullets":
                        lis = []
                        for lvl, btxt in b["items"]:
                            cls = ' class="sub-bullet"' if lvl > 0 else ""
                            lis.append(f"<li{cls}>{md_inline_to_html(btxt)}</li>")
                        blocks_html.append(f'<ul class="body-list">{"".join(lis)}</ul>')
                    elif b["type"] == "table":
                        rows = b["rows"]
                        col_widths = compute_table_col_widths(rows, FULL_CONTENT_WIDTH)
                        col_tags = "".join(
                            f'<col style="width: {round((cw / FULL_CONTENT_WIDTH) * 100, 1)}%;">'
                            for cw in col_widths
                        )
                        ths = "".join(f"<th>{md_inline_to_html(c)}</th>" for c in rows[0])
                        trs = []
                        for r in rows[1:]:
                            tds = "".join(f"<td>{md_inline_to_html(c)}</td>" for c in r)
                            trs.append(f"<tr>{tds}</tr>")
                        blocks_html.append(
                            f'<table class="styled-table"><colgroup>{col_tags}</colgroup>'
                            f'<thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table>'
                        )

                img_uri = image_to_data_uri(s.get("image"))
                img_layout = s.get("image_layout", "right")
                if img_uri and img_layout == "top":
                    body_html = f"""
                    <div class="slide-body stacked-top"{scale_style}>
                      <div class="stacked-image"><img src="{img_uri}" alt="Slide visual"></div>
                      <div class="stacked-caption">{''.join(blocks_html)}</div>
                    </div>
                    """
                elif img_uri and img_layout == "bottom":
                    body_html = f"""
                    <div class="slide-body stacked-bottom"{scale_style}>
                      <div class="stacked-caption">{''.join(blocks_html)}</div>
                      <div class="stacked-image"><img src="{img_uri}" alt="Slide visual"></div>
                    </div>
                    """
                elif img_uri:
                    body_html = f"""
                    <div class="slide-body split-layout"{scale_style}>
                      <div class="content-col">{''.join(blocks_html)}</div>
                      <div class="image-col"><img src="{img_uri}" alt="Slide visual"></div>
                    </div>
                    """
                else:
                    body_html = f"""
                    <div class="slide-body"{scale_style}>
                      <div class="content-col full">{''.join(blocks_html)}</div>
                    </div>
                    """

        rendered_cards.append(f"""
        <section class="slide{active_cls}" data-index="{idx}">
          <div class="chrome-sidebar"></div>
          <div class="chrome-corner"></div>
          <div class="chrome-banner"></div>
          {chrome_extra}
          {body_html}
          <div class="ncdp-footer">National Center for Disaster Preparedness &bull; Columbia Climate School</div>
        </section>
        """)

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(site_name)} — Interactive Presentation</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
  :root {{
    --crimson: #AB192D;
    --dark: #222222;
    --sidebar: #7F7F7F;
    --card-bg: #E7E6E6;
    --alt-row: #F3F3F3;
    --border: #D0D0D0;
    --link: #0563C1;
    --slide-scale: 1;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    background: #1b1d21;
    color: var(--dark);
    font-family: 'Open Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    display: flex;
    flex-direction: column;
    min-height: 100vh;
  }}
  .toolbar {{
    background: #111317;
    color: #f1f3f5;
    padding: 0.55rem 1.25rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    border-bottom: 1px solid #2c3038;
    position: sticky;
    top: 0;
    z-index: 50;
  }}
  .toolbar-title {{
    font-size: 0.92rem;
    font-weight: 600;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .toolbar-controls {{
    display: flex;
    align-items: center;
    gap: 0.5rem;
  }}
  button {{
    background: #2b303b;
    color: #fff;
    border: 1px solid #404756;
    border-radius: 6px;
    padding: 0.35rem 0.75rem;
    font-family: inherit;
    font-size: 0.85rem;
    cursor: pointer;
  }}
  button:hover {{ background: var(--crimson); border-color: var(--crimson); }}
  .counter {{
    font-size: 0.85rem;
    min-width: 4.5rem;
    text-align: center;
    color: #cbd2d9;
  }}
  main.deck-stage {{
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 1.25rem;
  }}
  body.scroll-mode main.deck-stage {{
    flex-direction: column;
    gap: 1.75rem;
  }}
  .slide {{
    position: relative;
    width: min(94vw, 1200px);
    aspect-ratio: 16 / 9;
    background: #ffffff;
    box-shadow: 0 18px 45px rgba(0, 0, 0, 0.45);
    overflow: hidden;
    display: none;
  }}
  .slide.active,
  body.scroll-mode .slide {{
    display: block;
  }}
  .chrome-sidebar {{
    position: absolute; left: 0; top: 0; bottom: 0; width: 2.2%;
    background: var(--sidebar);
  }}
  .chrome-corner {{
    position: absolute; left: 0; top: 0; width: 2.2%; height: 16.35%;
    background: #000000;
  }}
  .chrome-banner {{
    position: absolute; left: 2.2%; top: 0; right: 0; height: 16.35%;
    background: var(--crimson);
  }}
  .slide-title-bar {{
    position: absolute; left: 5.5%; top: 2.2%; right: 11%; height: 12%;
    display: flex; align-items: center;
    color: #ffffff;
    font-size: clamp(1.05rem, 2.05vw, 1.65rem);
    font-weight: 400;
    line-height: 1.2;
  }}
  .slide-num-tab {{
    position: absolute; right: 4.3%; top: 15.8%; width: 6.8%; height: 6.2%;
    background: var(--crimson);
    color: #ffffff;
    font-weight: 700;
    font-size: clamp(0.75rem, 1.05vw, 0.95rem);
    display: flex; align-items: center; justify-content: center;
  }}
  .ncdp-footer {{
    position: absolute; left: 5.8%; bottom: 2.1%;
    font-size: clamp(0.62rem, 0.82vw, 0.76rem);
    color: #666666;
  }}
  /* Cover Slide */
  .cover-layout {{
    position: absolute; inset: 18% 5% 8% 6%;
    display: flex; align-items: center; justify-content: center; gap: 3%;
  }}
  .cover-hero {{ flex: 0 0 38%; height: 85%; display: flex; align-items: center; justify-content: center; }}
  .cover-hero img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
  .cover-text {{ flex: 1; text-align: center; padding: 0 1rem; }}
  .cover-text h1 {{
    margin: 0; font-weight: 400; color: var(--dark);
    font-size: clamp(1.25rem, 2.4vw, 2.05rem); line-height: 1.3;
  }}
  .cover-subtitle {{
    margin-top: 1rem; color: var(--crimson); font-weight: 700;
    font-size: clamp(0.85rem, 1.3vw, 1.12rem);
  }}
  /* Content Slides */
  .slide-body {{
    position: absolute;
    left: 5.2%; top: 18.2%; right: 4.8%; bottom: 8.5%;
    display: flex; flex-direction: column;
    overflow: hidden;
  }}
  .slide-body.split-layout {{
    flex-direction: row; gap: 2.2%; align-items: flex-start;
  }}
  .slide-body.stacked-top,
  .slide-body.stacked-bottom {{
    flex-direction: column; align-items: center; justify-content: center; gap: 0.65rem;
  }}
  .stacked-image {{
    flex: 1 1 auto; min-height: 0; width: 100%;
    display: flex; align-items: center; justify-content: center;
  }}
  .stacked-image img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
  .stacked-caption {{ flex: 0 0 auto; width: 100%; text-align: center; }}
  .stacked-caption .body-para {{ margin-bottom: 0; }}
  .content-col {{ flex: 1; min-width: 0; }}
  .image-col {{
    flex: 0 0 34%; height: 92%;
    display: flex; align-items: center; justify-content: center;
  }}
  .image-col img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
  .body-subheading {{
    margin: calc(0.42rem * var(--slide-scale)) 0 calc(0.24rem * var(--slide-scale)) 0;
    color: var(--crimson);
    font-weight: 700;
    font-size: calc(clamp(0.82rem, {round(body_vw * 1.12, 3)}vw, {round(body_max_rem * 1.12, 3)}rem) * var(--slide-scale));
    line-height: 1.3;
  }}
  .body-subheading:first-child {{ margin-top: 0; }}
  .body-para {{
    margin: 0 0 calc(0.40rem * var(--slide-scale)) 0;
    font-size: calc(clamp(0.74rem, {body_vw}vw, {body_max_rem}rem) * var(--slide-scale));
    line-height: 1.34;
  }}
  .inline-callout {{
    background: var(--card-bg);
    border-radius: 8px;
    padding: calc(0.42rem * var(--slide-scale)) 0.85rem;
    margin: 0 0 calc(0.42rem * var(--slide-scale)) 0;
    text-align: center;
    font-size: calc(clamp(0.78rem, {round(body_vw * 1.05, 3)}vw, {round(body_max_rem * 1.05, 3)}rem) * var(--slide-scale));
  }}
  .body-list {{
    margin: 0 0 calc(0.40rem * var(--slide-scale)) 1.30rem; padding: 0;
    font-size: calc(clamp(0.74rem, {body_vw}vw, {body_max_rem}rem) * var(--slide-scale));
    line-height: 1.33;
  }}
  .body-list li {{ margin-bottom: calc(0.26rem * var(--slide-scale)); }}
  .body-list li.sub-bullet {{
    margin-left: 1.20rem;
    font-size: calc(clamp(0.69rem, {round(body_vw * 0.92, 3)}vw, {round(body_max_rem * 0.92, 3)}rem) * var(--slide-scale));
  }}
  .inline-code {{ color: var(--crimson); font-weight: 700; font-family: inherit; }}
  a {{ color: var(--link); text-decoration: none; font-weight: 600; }}
  a:hover {{ text-decoration: underline; }}
  /* Tables */
  .styled-table {{
    width: 100%; border-collapse: collapse; margin-bottom: 0.42rem;
    font-size: calc(clamp(0.67rem, {round(body_vw * 0.86, 3)}vw, {round(body_max_rem * 0.88, 3)}rem) * var(--slide-scale));
  }}
  .styled-table th {{
    background: var(--crimson); color: #ffffff; text-align: left;
    padding: calc(0.30rem * var(--slide-scale)) 0.46rem; border: 1px solid var(--crimson);
  }}
  .styled-table td {{
    padding: calc(0.28rem * var(--slide-scale)) 0.46rem; border: 1px solid var(--border);
    vertical-align: middle; line-height: 1.26;
  }}
  .styled-table tbody tr:nth-child(odd) {{ background: var(--alt-row); }}
  .styled-table tbody tr:nth-child(even) {{ background: #ffffff; }}
  /* Live Streamlit Iframe Slide */
  .streamlit-body {{ top: 19.2%; bottom: 6.5%; }}
  .live-frame {{
    position: relative; flex: 1; width: 100%;
    border: 1px solid var(--border); background: #fff;
  }}
  .live-frame iframe {{ width: 100%; height: 100%; border: none; }}
  .live-loading {{
    position: absolute; inset: 0; z-index: 2;
    display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 0.75rem;
    background: #ffffff; color: #555555; font-size: 0.9rem;
    transition: opacity 0.45s ease;
  }}
  .live-loading.done {{ opacity: 0; pointer-events: none; }}
  .live-spinner {{
    width: 36px; height: 36px;
    border: 4px solid #ddd; border-top-color: var(--crimson); border-radius: 50%;
    animation: spin 0.9s linear infinite;
  }}
  @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
  .streamlit-footer {{
    text-align: right; padding-top: 0.35rem;
    font-size: clamp(0.68rem, 0.92vw, 0.82rem);
  }}
  /* Print / PDF Export: enforce page break after every slide */
  @media print {{
    body {{ background: #ffffff; }}
    .toolbar {{ display: none; }}
    main.deck-stage {{ display: block; padding: 0; }}
    .slide {{
      display: block !important;
      width: 100%;
      box-shadow: none;
      page-break-after: always;
      break-after: page;
    }}
    .slide:last-child {{
      page-break-after: auto;
      break-after: auto;
    }}
  }}
</style>
</head>
<body>
  <header class="toolbar">
    <div class="toolbar-title">{html.escape(site_name)}</div>
    <div class="toolbar-controls">
      <button type="button" id="btn-prev">&larr; Prev</button>
      <span class="counter" id="slide-counter">1 / {total}</span>
      <button type="button" id="btn-next">Next &rarr;</button>
      <button type="button" id="btn-mode">Scroll View</button>
    </div>
  </header>
  <main class="deck-stage">
    {''.join(rendered_cards)}
  </main>
  <script>
    (function() {{
      const slides = Array.from(document.querySelectorAll('.slide'));
      const counter = document.getElementById('slide-counter');
      const btnMode = document.getElementById('btn-mode');
      let current = 0;

      function show(idx) {{
        if (!slides.length) return;
        current = (idx + slides.length) % slides.length;
        slides.forEach((s, i) => s.classList.toggle('active', i === current));
        counter.textContent = (current + 1) + ' / ' + slides.length;
      }}

      document.getElementById('btn-prev').addEventListener('click', () => show(current - 1));
      document.getElementById('btn-next').addEventListener('click', () => show(current + 1));
      btnMode.addEventListener('click', () => {{
        const isScroll = document.body.classList.toggle('scroll-mode');
        btnMode.textContent = isScroll ? 'Slide View' : 'Scroll View';
      }});

      document.addEventListener('keydown', (e) => {{
        if (document.body.classList.contains('scroll-mode')) return;
        if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {{
          e.preventDefault(); show(current + 1);
        }} else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {{
          e.preventDefault(); show(current - 1);
        }}
      }});
    }})();
  </script>
</body>
</html>
"""
    output_html_path.parent.mkdir(parents=True, exist_ok=True)
    output_html_path.write_text(doc, encoding="utf-8")


# ---------------------------------------------------------------------------
# Optional Live Streamlit Screenshot Capture for PPTX
# ---------------------------------------------------------------------------
def capture_streamlit_screenshot(url: str, dest_png: Path) -> bool:
    """Optionally capture a live screenshot of an embedded Streamlit dashboard via Playwright."""
    try:
        from playwright.sync_api import sync_playwright
        import time

        dest_png.parent.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1280, "height": 800}, device_scale_factor=2)
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.locator('[data-testid="stAppViewContainer"], .stApp').first.wait_for(timeout=45000)
            time.sleep(5)
            page.screenshot(path=str(dest_png), full_page=False)
            browser.close()
        return True
    except Exception as exc:
        print(f"  [warn] Live capture skipped for {url}: {exc}")
        return False


def find_screenshot_for_md(md_filename: str, docs_dir: Path) -> Optional[Path]:
    """Match e.g. 'interactiveexercises_HVTalert_2.md' -> 'HVTalert_results.png'."""
    stem = Path(md_filename).stem
    m = re.search(r"_(HVT[a-zA-Z0-9]+)_2$", stem)
    keyword = m.group(1) if m else stem

    candidates = [
        docs_dir / "download_streamlit_images" / "streamlit_screenshots" / f"{keyword}_results.png",
        docs_dir / f"{keyword}_results.png",
    ]
    for c in candidates:
        if c.exists():
            return c
    for png in docs_dir.rglob("*.png"):
        if keyword.lower() in png.name.lower():
            return png
    return None


# ---------------------------------------------------------------------------
# Cover Slide Image Selection (First Image in index.md)
# ---------------------------------------------------------------------------
def extract_cover_hero_image(
    parsed_index: Dict,
    docs_dir: Path,
    rel_md: str,
    cover_image_override: Optional[str] = None,
) -> Tuple[Optional[Path], Optional[int]]:
    """
    Determine the Title/Cover slide image:
      1. If --cover-image is explicitly passed, use that.
      2. Otherwise, pick the FIRST image found in index.md (Markdown or HTML <img>)
         and remove that image block from index.md so it doesn't repeat on Slide 2.
      3. Fallback to cover.png / hero.png / evacuated.png if index.md has no image.
    """
    if cover_image_override:
        p = Path(cover_image_override)
        if p.exists():
            return p, None
        resolved = resolve_doc_image(cover_image_override, docs_dir, rel_md)
        if resolved:
            return resolved, None

    for sec in parsed_index["sections"]:
        for idx_b, b in enumerate(sec["blocks"]):
            if b["type"] == "image":
                resolved = resolve_doc_image(b["src"], docs_dir, rel_md)
                if resolved is not None:
                    sec["blocks"].pop(idx_b)
                    return resolved, b.get("width_px")

    for fallback_name in ("cover.png", "hero.png", "evacuated.png"):
        candidate = docs_dir / fallback_name
        if candidate.exists():
            return candidate, None
    return None, None


# ---------------------------------------------------------------------------
# Main Conversion Pipeline
# ---------------------------------------------------------------------------
def convert_zip_to_pptx_and_html(
    zip_path: str,
    output_pptx: str,
    output_html: Optional[str] = None,
    template_pptx: Optional[str] = None,
    capture_live: bool = False,
    body_font_pt: float = 12.5,
    min_font_pt: float = 9.5,
    cover_image: Optional[str] = None,
):
    tmp_dir = tempfile.mkdtemp(prefix="gh_md_to_pptx_")
    try:
        _, docs_dir, mkdocs_path = extract_zip_and_locate_project(zip_path, tmp_dir)
        site_name, nav_entries = load_mkdocs_config(mkdocs_path, docs_dir)
        prs, using_template = initialize_presentation(template_pptx)

        parsed_pages: List[Tuple[str, Dict]] = []
        for nav_label, rel_md in nav_entries:
            md_path = docs_dir / rel_md
            if not md_path.exists():
                continue
            parsed = parse_markdown_sections(
                md_path.read_text(encoding="utf-8", errors="replace"),
                fallback_title=nav_label,
            )
            parsed_pages.append((rel_md, parsed))

        slide_counter = 1
        html_slides: List[Dict] = []

        for idx, (rel_md, parsed) in enumerate(parsed_pages):
            # 1. Cover slide + Slide 2 (plus any post-page-break sections in index.md)
            if idx == 0 and Path(rel_md).name.lower() == "index.md":
                hero_img, hero_w_px = extract_cover_hero_image(parsed, docs_dir, rel_md, cover_image)

                first_sec = parsed["sections"][0]
                subtitle = ""
                remaining_blocks = []
                for b in first_sec["blocks"]:
                    if not subtitle and b["type"] == "paragraph" and b["text"].startswith("**"):
                        subtitle = b["text"]
                    else:
                        remaining_blocks.append(b)

                build_cover_slide(
                    prs, using_template, site_name, subtitle, hero_img, html_slides, hero_width_px=hero_w_px
                )
                slide_counter += 1

                if remaining_blocks:
                    slide_counter = build_content_slides(
                        prs,
                        using_template,
                        slide_counter,
                        first_sec["title"] or parsed["page_title"],
                        remaining_blocks,
                        docs_dir,
                        rel_md,
                        html_slides,
                        body_font_pt=body_font_pt,
                        min_font_pt=min_font_pt,
                    )

                for extra_sec in parsed["sections"][1:]:
                    if extra_sec["blocks"]:
                        slide_counter = build_content_slides(
                            prs,
                            using_template,
                            slide_counter,
                            extra_sec["title"] or parsed["page_title"],
                            extra_sec["blocks"],
                            docs_dir,
                            rel_md,
                            html_slides,
                            body_font_pt=body_font_pt,
                            min_font_pt=min_font_pt,
                        )
                continue

            # 2. Results page (*_2.md with Streamlit <iframe>)
            if parsed["iframe_url"]:
                shot = find_screenshot_for_md(rel_md, docs_dir)
                if capture_live:
                    target_shot = shot or (
                        docs_dir / "download_streamlit_images" / "streamlit_screenshots" / f"{Path(rel_md).stem}.png"
                    )
                    if capture_streamlit_screenshot(parsed["iframe_url"], target_shot):
                        shot = target_shot

                build_results_dashboard_slide(
                    prs, using_template, slide_counter, parsed["page_title"], shot, parsed["iframe_url"], html_slides
                )
                slide_counter += 1
                continue

            # 3. Standard content pages (1 Markdown page = 1 slide unit, split on any page break or vertical overflow)
            for sec in parsed["sections"]:
                if sec["blocks"]:
                    slide_counter = build_content_slides(
                        prs,
                        using_template,
                        slide_counter,
                        sec["title"] or parsed["page_title"],
                        sec["blocks"],
                        docs_dir,
                        rel_md,
                        html_slides,
                        body_font_pt=body_font_pt,
                        min_font_pt=min_font_pt,
                    )

        out_pptx_path = Path(output_pptx)
        out_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out_pptx_path))

        out_html_path = Path(output_html) if output_html else out_pptx_path.with_suffix(".html")
        render_html_deck(html_slides, site_name, out_html_path, body_font_pt=body_font_pt)

        print(f"Successfully generated PPTX: {out_pptx_path} ({len(prs.slides)} slides)")
        print(f"Successfully generated HTML: {out_html_path} ({len(html_slides)} slides with live Streamlit embeds)")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="Convert a GitHub.io / MkDocs Markdown ZIP into a PowerPoint deck and matching live HTML presentation."
    )
    parser.add_argument("zip_file", help="Path to the GitHub.io / MkDocs .zip file (e.g., hoopavalley.zip)")
    parser.add_argument("-o", "--output", default="HoopaValley_Workshop.pptx", help="Output .pptx path")
    parser.add_argument(
        "--html",
        default=None,
        help="Output .html presentation path (defaults to the same filename as -o with .html extension)",
    )
    parser.add_argument(
        "-t",
        "--template",
        default=None,
        help="Optional reference .pptx template to inherit Slide Master, theme, and logos from",
    )
    parser.add_argument(
        "--font-size",
        type=float,
        default=12.5,
        help="Preferred body text font size in points (default: 12.5)",
    )
    parser.add_argument(
        "--min-font-size",
        type=float,
        default=9.5,
        help="Minimum auto-fit body font size in points before splitting a slide (default: 9.5)",
    )
    parser.add_argument(
        "--cover-image",
        default=None,
        help="Optional image filename/path for Slide 1 cover (defaults to the first image in index.md)",
    )
    parser.add_argument(
        "--capture-live",
        action="store_true",
        help="Capture fresh screenshots of embedded Streamlit iframes for the PPTX via Playwright",
    )
    args = parser.parse_args()
    convert_zip_to_pptx_and_html(
        args.zip_file,
        args.output,
        args.html,
        args.template,
        args.capture_live,
        args.font_size,
        args.min_font_size,
        args.cover_image,
    )


if __name__ == "__main__":
    main()
