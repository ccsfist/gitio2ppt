#!/usr/bin/env python3
"""
Convert a GitHub.io / MkDocs Markdown ZIP archive into:
  1. A PowerPoint deck (.pptx) matching Scenario Course_NTICCReno_October 2026.pptx
     with zero empty placeholders, automatic slide splitting to prevent text
     overlap, and static dashboard screenshots.
  2. A matching self-contained HTML presentation (.html) with the same visual
     slide chrome, embedded images, and LIVE interactive Streamlit iframes on
     the results slides.

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

CONTENT_TOP_START = 1.15      # Below 0.92" crimson banner
MAX_CONTENT_BOTTOM = 4.92     # Above 5.11" NCDP footer / bottom margin
CONTENT_LEFT = 0.58
FULL_CONTENT_WIDTH = 8.85
SPLIT_CONTENT_WIDTH = 5.35


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


def load_mkdocs_config(mkdocs_path: Optional[Path], docs_dir: Path) -> Tuple[str, List[str]]:
    """Read site_name and ordered Markdown file list from mkdocs.yml (works with or without PyYAML)."""
    site_name = "Two-Way Emergency Communication Workshop"
    nav_files: List[str] = []

    if mkdocs_path and mkdocs_path.exists():
        raw = mkdocs_path.read_text(encoding="utf-8", errors="replace")
        if yaml is not None:
            cleaned_yaml = re.sub(r"![a-zA-Z0-9_.-]+", "", raw)
            cfg = yaml.safe_load(cleaned_yaml) or {}
            site_name = cfg.get("site_name", site_name)

            def flatten_nav(items):
                for item in items or []:
                    if isinstance(item, str) and item.endswith(".md"):
                        nav_files.append(item)
                    elif isinstance(item, dict):
                        for val in item.values():
                            if isinstance(val, str) and val.endswith(".md"):
                                nav_files.append(val)
                            elif isinstance(val, list):
                                flatten_nav(val)

            flatten_nav(cfg.get("nav", []))
        else:
            m_site = re.search(r"^site_name:\s*(.+)$", raw, re.MULTILINE)
            if m_site:
                site_name = m_site.group(1).strip().strip("\"'")
            for m_md in re.finditer(r"['\"]([A-Za-z0-9_./-]+\.md)['\"]", raw):
                nav_files.append(m_md.group(1))

    if not nav_files:
        all_md = sorted(p.name for p in docs_dir.glob("*.md"))
        if "index.md" in all_md:
            all_md.remove("index.md")
            nav_files = ["index.md"] + all_md
        else:
            nav_files = all_md

    return site_name, nav_files


# ---------------------------------------------------------------------------
# Template & Slide Chrome Helpers (Zero Empty Placeholders)
# ---------------------------------------------------------------------------
def initialize_presentation(template_path: Optional[str]) -> Tuple[Presentation, bool]:
    """Load template PPTX (clearing sample slides while keeping Master/Logos) or create blank 16:9 deck."""
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
        for i in range(len(prs.slides) - 1, -1, -1):
            r_id = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(r_id)
            del prs.slides._sldIdLst[i]
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
    Remove all placeholder shapes from the slide XML so PowerPoint/Google Slides
    never displays 'Click to add title', 'Click to add subtitle', or 'Click to add text'.
    """
    title_shp = None
    if keep_title:
        try:
            title_shp = slide.shapes.title
        except Exception:
            title_shp = None

    for shp in list(slide.shapes):
        if shp.is_placeholder:
            if keep_title and title_shp is not None and shp == title_shp:
                continue
            sp = shp._element
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
    """Create a slide with all default 'Click to add...' placeholders stripped."""
    if is_cover:
        layout = get_layout(prs, ["TITLE", "Title Slide", "Blank"])
        slide = prs.slides.add_slide(layout)
        prune_unused_placeholders(slide, keep_title=False)
        if not using_template:
            apply_standalone_chrome(slide, slide_num, is_cover=True)
        return slide

    layout = get_layout(prs, ["Title and Content", "TITLE_ONLY", "Blank"])
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
# Inline Markdown Formatting & Height Estimation Helpers
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
            out.append(html.escape(tok))
    return "".join(out)


def add_formatted_runs(
    paragraph,
    text: str,
    base_size_pt: float = 14.5,
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
            r.text = tok
            r.font.name = FONT_FAMILY
            r.font.size = Pt(base_size_pt)
            r.font.bold = base_bold
            r.font.italic = base_italic
            r.font.color.rgb = base_color


def set_bullet_xml(paragraph, level: int = 0):
    """Enable native PowerPoint bullet formatting on a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.28 + level * 0.25))))
    pPr.set("indent", str(int(Inches(-0.18))))
    buChar = pPr.find(qn("a:buChar"))
    if buChar is None:
        buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "•" if level == 0 else "▪")


def estimate_text_height_in(text: str, width_in: float, font_pt: float, is_bullet: bool = False) -> float:
    """Estimate the rendered height (in inches) of a wrapped paragraph or bullet."""
    plain = strip_md_inline(text)
    if not plain:
        return 0.22
    effective_w = max(1.2, width_in - (0.35 if is_bullet else 0.14))
    # Conservative Open Sans average glyph width (~0.53 * font_pt)
    chars_per_line = max(12, int((effective_w * 72.0) / (font_pt * 0.54)))
    lines = max(1, math.ceil(len(plain) / chars_per_line))
    line_height_in = (font_pt * 1.32) / 72.0
    extra_pad = 0.09 if is_bullet else 0.06
    return lines * line_height_in + extra_pad


def estimate_table_height_in(rows: List[List[str]], width_in: float) -> float:
    """Estimate total table height accounting for multi-line wrapped cells."""
    if not rows:
        return 0.5
    n_cols = max(len(r) for r in rows)
    col_w = max(1.0, width_in / max(1, n_cols))
    total_h = 0.0
    for r_idx, row in enumerate(rows):
        font_pt = 12.5 if r_idx == 0 else 12.0
        max_lines = 1
        for cell_txt in row:
            plain = strip_md_inline(cell_txt)
            chars_per_line = max(8, int(((col_w - 0.20) * 72.0) / (font_pt * 0.55)))
            lines = max(1, math.ceil(len(plain) / chars_per_line))
            max_lines = max(max_lines, lines)
        row_h = max(0.36, max_lines * ((font_pt * 1.28) / 72.0) + 0.14)
        total_h += row_h
    return total_h


# ---------------------------------------------------------------------------
# Markdown Parser
# ---------------------------------------------------------------------------
def parse_markdown_sections(md_text: str) -> Dict:
    """
    Parse a MkDocs Markdown file into structured slide units:
    - page_title (# ...)
    - iframe_url (if Streamlit embed exists)
    - sms_callout (if 'To start the activity, send ...' block exists)
    - sections (list of dicts with title and blocks)
    """
    iframe_match = re.search(r'<iframe[^>]+src=["\']([^"\']+)["\']', md_text, re.IGNORECASE)
    iframe_url = iframe_match.group(1).replace("&amp;", "&") if iframe_match else None

    sms_callout = None
    sms_match = re.search(
        r"To start the activity,\s*send[^\n]*\n+([#*\sA-Za-z0-9_+-]+by text or WhatsApp to[^\n]+)\n+To end the activity at any time send\s+([^\n]+)",
        md_text,
        re.IGNORECASE | re.DOTALL,
    )
    if sms_match:
        line1 = strip_md_inline(sms_match.group(1))
        line2 = strip_md_inline(sms_match.group(2))
        kw_match = re.match(r"^([A-Za-z0-9_-]+)\s+by text or WhatsApp to\s+(.+)$", line1, re.IGNORECASE)
        if kw_match:
            sms_callout = {
                "keyword": kw_match.group(1).strip(),
                "phone": kw_match.group(2).strip(),
                "end_cmd": line2.strip(),
            }
        md_text = md_text[:sms_match.start()] + "\n" + md_text[sms_match.end():]

    # Strip <style>, <div id="slide-config">, and <iframe> tags
    cleaned = re.sub(r"<style>.*?</style>", "", md_text, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<div[^>]*>.*?</div>", "", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"<iframe[^>]*>.*?</iframe>", "", cleaned, flags=re.DOTALL | re.IGNORECASE)

    lines = cleaned.splitlines()
    page_title = "Workshop Slide"
    sections = []
    current = {"title": None, "blocks": []}

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped.startswith("# ") and not stripped.startswith("## "):
            page_title = strip_md_inline(stripped[2:])
            if current["title"] is None:
                current["title"] = page_title
            i += 1
            continue

        if stripped.startswith("## "):
            if current["blocks"]:
                sections.append(current)
            current = {"title": strip_md_inline(stripped[3:]), "blocks": []}
            i += 1
            continue

        img_m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", stripped)
        if img_m:
            current["blocks"].append({"type": "image", "alt": img_m.group(1), "src": img_m.group(2).strip()})
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:|-]+\|$", lines[i + 1].strip()):
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row_str = lines[i].strip()
                if not re.match(r"^\|[\s:|-]+\|$", row_str):
                    cells = [c.strip() for c in row_str.strip("|").split("|")]
                    table_rows.append(cells)
                i += 1
            if table_rows:
                current["blocks"].append({"type": "table", "rows": table_rows})
            continue

        if re.match(r"^(\s*)[-*+]\s+", line):
            bullets = []
            while i < len(lines) and re.match(r"^(\s*)[-*+]\s+", lines[i]):
                m = re.match(r"^(\s*)[-*+]\s+(.*)$", lines[i])
                indent_spaces = len(m.group(1))
                lvl = 1 if indent_spaces >= 2 else 0
                bullets.append((lvl, m.group(2).strip()))
                i += 1
            current["blocks"].append({"type": "bullets", "items": bullets})
            continue

        para_lines = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt.startswith("#")
                or nxt.startswith("|")
                or re.match(r"^[-*+]\s+", nxt)
                or re.match(r"^!\[", nxt)
            ):
                break
            para_lines.append(nxt)
            i += 1
        current["blocks"].append({"type": "paragraph", "text": " ".join(para_lines)})

    if current["blocks"] or not sections:
        if current["title"] is None:
            current["title"] = page_title
        sections.append(current)

    return {
        "page_title": page_title,
        "iframe_url": iframe_url,
        "sms_callout": sms_callout,
        "sections": sections,
    }


# ---------------------------------------------------------------------------
# Slide Builders (PPTX + Structured Slide Models for HTML Output)
# ---------------------------------------------------------------------------
def add_picture_fit(slide, img_path: Path, left: float, top: float, max_w: float, max_h: float):
    """Insert an image scaled proportionally within (max_w, max_h) and centered in the box."""
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    aspect = w_px / float(h_px)
    box_aspect = max_w / float(max_h)

    if aspect >= box_aspect:
        final_w = max_w
        final_h = max_w / aspect
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
):
    """Build Slide 1 cover slide for both PPTX and HTML."""
    slide = create_clean_slide(prs, using_template, 1, is_cover=True)

    if hero_img and hero_img.exists():
        add_picture_fit(slide, hero_img, left=0.75, top=1.35, max_w=3.75, max_h=3.55)
        text_left, text_w = 4.70, 4.90
    else:
        text_left, text_w = 1.00, 8.00

    title_box = slide.shapes.add_textbox(Inches(text_left), Inches(1.85), Inches(text_w), Inches(2.50))
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


def build_sms_activity_slide(
    prs: Presentation,
    using_template: bool,
    slide_num: int,
    slide_title: str,
    sms_info: Dict[str, str],
    results_url: Optional[str],
    html_slides: List[Dict],
):
    """Build the light-gray rounded SMS prompt card slide."""
    slide = create_clean_slide(prs, using_template, slide_num, is_cover=False)
    set_slide_title(slide, slide_title, using_template)

    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.45), Inches(1.55), Inches(9.10), Inches(2.85)
    )
    card.adjustments[0] = 0.04
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_BG
    card.line.fill.background()

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.4)
    tf.margin_right = Inches(0.4)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(10)
    add_formatted_runs(
        p1,
        f"To start the activity, send **{sms_info['keyword']}** by text or WhatsApp to",
        base_size_pt=24,
        base_color=COLOR_DARK,
    )

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_after = Pt(14)
    r_phone = p2.add_run()
    r_phone.text = sms_info["phone"]
    r_phone.font.name = FONT_FAMILY
    r_phone.font.size = Pt(30)
    r_phone.font.bold = True
    r_phone.font.color.rgb = COLOR_DARK

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    add_formatted_runs(
        p3,
        f"To end the activity at any time send **{sms_info['end_cmd']}**",
        base_size_pt=22,
        base_color=COLOR_DARK,
    )

    if results_url:
        link_box = slide.shapes.add_textbox(Inches(0.55), Inches(4.62), Inches(3.60), Inches(0.35))
        ltf = link_box.text_frame
        lp = ltf.paragraphs[0]
        lr = lp.add_run()
        lr.text = "Latest Results (Live Dashboard)"
        lr.font.name = FONT_FAMILY
        lr.font.size = Pt(11.5)
        lr.font.color.rgb = COLOR_LINK
        lr.hyperlink.address = results_url

    html_slides.append({
        "type": "sms",
        "num": slide_num,
        "title": strip_md_inline(slide_title),
        "sms_info": sms_info,
        "results_url": results_url,
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
      - HTML: embeds LIVE interactive Streamlit <iframe>!
    """
    slide = create_clean_slide(prs, using_template, slide_num, is_cover=False)
    set_slide_title(slide, slide_title, using_template)

    if screenshot_path and screenshot_path.exists():
        frame = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.12), Inches(8.70), Inches(3.62)
        )
        frame.fill.solid()
        frame.fill.fore_color.rgb = COLOR_WHITE
        frame.line.color.rgb = COLOR_BORDER
        frame.line.width = Pt(1.0)
        add_picture_fit(slide, screenshot_path, left=0.70, top=1.16, max_w=8.60, max_h=3.54)
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

    url_box = slide.shapes.add_textbox(Inches(3.60), Inches(4.80), Inches(5.90), Inches(0.35))
    utf = url_box.text_frame
    up = utf.paragraphs[0]
    up.alignment = PP_ALIGN.RIGHT
    ur = up.add_run()
    ur.text = f"Open Live Results: {iframe_url}"
    ur.font.name = FONT_FAMILY
    ur.font.size = Pt(10.5)
    ur.font.color.rgb = COLOR_LINK
    ur.hyperlink.address = iframe_url

    html_slides.append({
        "type": "streamlit",
        "num": slide_num,
        "title": strip_md_inline(slide_title),
        "iframe_url": iframe_url,
    })


def render_styled_table(slide, rows: List[List[str]], left: float, top: float, width: float, height: float):
    """Render a Markdown table into a formatted PowerPoint table matching the crimson/gray palette."""
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table

    for r_idx, row in enumerate(rows):
        is_header = r_idx == 0
        bg_color = COLOR_CRIMSON if is_header else (COLOR_ALT_ROW if r_idx % 2 == 1 else COLOR_WHITE)
        txt_color = COLOR_WHITE if is_header else COLOR_DARK
        font_sz = 12.5 if is_header else 12.0

        for c_idx in range(n_cols):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.10)
            cell.margin_right = Inches(0.10)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)

            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            val = row[c_idx] if c_idx < len(row) else ""
            add_formatted_runs(p, val, base_size_pt=font_sz, base_color=txt_color, base_bold=is_header)


def expand_splittable_blocks(blocks: List[Dict], width_in: float) -> List[Dict]:
    """
    Break oversized bullet lists or tables into smaller chunks so a single block
    never overflows a slide vertically, while keeping normal blocks intact.
    """
    max_single_slide_h = MAX_CONTENT_BOTTOM - CONTENT_TOP_START - 0.10
    expanded: List[Dict] = []

    for b in blocks:
        if b["type"] == "bullets":
            chunk: List[Tuple[int, str]] = []
            chunk_h = 0.0
            for lvl, btext in b["items"]:
                font_pt = 14.0 if lvl == 0 else 13.0
                item_h = estimate_text_height_in(btext, width_in, font_pt, is_bullet=True)
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
                cand_h = estimate_table_height_in(curr_rows + [r], width_in)
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


def build_content_slides(
    prs: Presentation,
    using_template: bool,
    start_slide_num: int,
    slide_title: str,
    blocks: List[Dict],
    docs_dir: Path,
    html_slides: List[Dict],
) -> int:
    """
    Build one or more content slides from `blocks`, automatically splitting onto
    continuation slides ('Title (Cont.)') whenever vertical height would exceed
    MAX_CONTENT_BOTTOM so text boxes and tables never overlap.
    Returns the next available slide number.
    """
    images = [b for b in blocks if b["type"] == "image"]
    non_images = [b for b in blocks if b["type"] != "image"]

    has_image = False
    img_path = None
    if images:
        candidate = docs_dir / images[0]["src"]
        if candidate.exists():
            has_image = True
            img_path = candidate

    content_width = SPLIT_CONTENT_WIDTH if has_image else FULL_CONTENT_WIDTH
    prepared_blocks = expand_splittable_blocks(non_images, content_width)

    slide_num = start_slide_num
    part_idx = 1

    def open_slide(part_number: int, include_image: bool):
        title_for_slide = slide_title if part_number == 1 else f"{strip_md_inline(slide_title)} (Cont.)"
        s = create_clean_slide(prs, using_template, slide_num, is_cover=False)
        set_slide_title(s, title_for_slide, using_template)
        if include_image and img_path is not None:
            add_picture_fit(s, img_path, left=6.10, top=1.25, max_w=3.45, max_h=3.55)
        html_rec = {
            "type": "content",
            "num": slide_num,
            "title": strip_md_inline(title_for_slide),
            "blocks": [],
            "image": img_path if include_image else None,
        }
        html_slides.append(html_rec)
        return s, html_rec

    slide, current_html = open_slide(part_idx, has_image)
    cursor_top = CONTENT_TOP_START

    for block in prepared_blocks:
        if block["type"] == "table":
            est_h = estimate_table_height_in(block["rows"], content_width)
            gap = 0.16
        elif block["type"] == "paragraph":
            est_h = estimate_text_height_in(block["text"], content_width, 14.0, is_bullet=False)
            gap = 0.12
        elif block["type"] == "bullets":
            est_h = sum(
                estimate_text_height_in(btxt, content_width, 14.0 if lvl == 0 else 13.0, is_bullet=True)
                for lvl, btxt in block["items"]
            )
            gap = 0.12
        else:
            continue

        # Start a new continuation slide if this block would overflow the bottom margin
        if cursor_top > CONTENT_TOP_START + 0.05 and (cursor_top + est_h > MAX_CONTENT_BOTTOM):
            slide_num += 1
            part_idx += 1
            slide, current_html = open_slide(part_idx, include_image=False)
            cursor_top = CONTENT_TOP_START

        render_h = min(est_h, MAX_CONTENT_BOTTOM - cursor_top)

        if block["type"] == "table":
            render_styled_table(slide, block["rows"], CONTENT_LEFT, cursor_top, content_width, render_h)
        elif block["type"] == "paragraph":
            tb = slide.shapes.add_textbox(Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            add_formatted_runs(p, block["text"], base_size_pt=14.0, base_color=COLOR_DARK)
        elif block["type"] == "bullets":
            tb = slide.shapes.add_textbox(Inches(CONTENT_LEFT), Inches(cursor_top), Inches(content_width), Inches(render_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            for idx_b, (lvl, btext) in enumerate(block["items"]):
                p = tf.paragraphs[0] if idx_b == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(5)
                set_bullet_xml(p, level=lvl)
                add_formatted_runs(p, btext, base_size_pt=14.0 if lvl == 0 else 13.0, base_color=COLOR_DARK)

        current_html["blocks"].append(block)
        cursor_top += render_h + gap

    return slide_num + 1


# ---------------------------------------------------------------------------
# HTML Presentation Builder (Visual Match + Live Streamlit Iframes)
# ---------------------------------------------------------------------------
def image_to_data_uri(img_path: Optional[Path]) -> str:
    """Encode local PNG/JPG into a self-contained base64 data URI for the HTML deck."""
    if not img_path or not img_path.exists():
        return ""
    ext = img_path.suffix.lower().lstrip(".") or "png"
    mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
    b64 = base64.b64encode(img_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def render_html_deck(html_slides: List[Dict], site_name: str, output_html_path: Path):
    """Generate a single-file 16:9 HTML slide deck with live Streamlit iframes."""
    total = len(html_slides)
    rendered_cards: List[str] = []

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

            if s_type == "sms":
                sms = s["sms_info"]
                res_link = (
                    f'<a class="results-pill" href="{html.escape(s["results_url"], quote=True)}" target="_blank" rel="noopener">Latest Results (Live Dashboard) &rarr;</a>'
                    if s.get("results_url") else ""
                )
                body_html = f"""
                <div class="slide-body sms-body">
                  <div class="sms-card">
                    <p class="sms-line1">To start the activity, send <strong>{html.escape(sms['keyword'])}</strong> by text or WhatsApp to</p>
                    <p class="sms-phone">{html.escape(sms['phone'])}</p>
                    <p class="sms-line3">To end the activity at any time send <strong>{html.escape(sms['end_cmd'])}</strong></p>
                  </div>
                  {res_link}
                </div>
                """
            elif s_type == "streamlit":
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
                blocks_html: List[str] = []
                for b in s.get("blocks", []):
                    if b["type"] == "paragraph":
                        blocks_html.append(f'<p class="body-para">{md_inline_to_html(b["text"])}</p>')
                    elif b["type"] == "bullets":
                        lis = []
                        for lvl, btxt in b["items"]:
                            cls = ' class="sub-bullet"' if lvl > 0 else ""
                            lis.append(f"<li{cls}>{md_inline_to_html(btxt)}</li>")
                        blocks_html.append(f'<ul class="body-list">{"".join(lis)}</ul>')
                    elif b["type"] == "table":
                        rows = b["rows"]
                        ths = "".join(f"<th>{md_inline_to_html(c)}</th>" for c in rows[0])
                        trs = []
                        for r in rows[1:]:
                            tds = "".join(f"<td>{md_inline_to_html(c)}</td>" for c in r)
                            trs.append(f"<tr>{tds}</tr>")
                        blocks_html.append(
                            f'<table class="styled-table"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table>'
                        )

                img_uri = image_to_data_uri(s.get("image"))
                if img_uri:
                    body_html = f"""
                    <div class="slide-body split-layout">
                      <div class="content-col">{''.join(blocks_html)}</div>
                      <div class="image-col"><img src="{img_uri}" alt="Slide visual"></div>
                    </div>
                    """
                else:
                    body_html = f"""
                    <div class="slide-body">
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
  /* Slide Chrome matching Scenario Course_NTICCReno_October 2026.pptx */
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
    left: 5.8%; top: 20.5%; right: 5.2%; bottom: 8.5%;
    display: flex; flex-direction: column;
    overflow: hidden;
  }}
  .slide-body.split-layout {{
    flex-direction: row; gap: 2.5%; align-items: flex-start;
  }}
  .content-col {{ flex: 1; min-width: 0; }}
  .image-col {{
    flex: 0 0 36%; height: 95%;
    display: flex; align-items: center; justify-content: center;
  }}
  .image-col img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
  .body-para {{
    margin: 0 0 0.65rem 0;
    font-size: clamp(0.82rem, 1.32vw, 1.12rem);
    line-height: 1.42;
  }}
  .body-list {{
    margin: 0 0 0.65rem 1.35rem; padding: 0;
    font-size: clamp(0.82rem, 1.32vw, 1.12rem);
    line-height: 1.4;
  }}
  .body-list li {{ margin-bottom: 0.42rem; }}
  .body-list li.sub-bullet {{
    margin-left: 1.25rem;
    font-size: clamp(0.78rem, 1.2vw, 1.02rem);
  }}
  .inline-code {{ color: var(--crimson); font-weight: 700; font-family: inherit; }}
  a {{ color: var(--link); text-decoration: none; font-weight: 600; }}
  a:hover {{ text-decoration: underline; }}
  /* Tables */
  .styled-table {{
    width: 100%; border-collapse: collapse; margin-bottom: 0.65rem;
    font-size: clamp(0.72rem, 1.08vw, 0.94rem);
  }}
  .styled-table th {{
    background: var(--crimson); color: #ffffff; text-align: left;
    padding: 0.42rem 0.55rem; border: 1px solid var(--crimson);
  }}
  .styled-table td {{
    padding: 0.38rem 0.55rem; border: 1px solid var(--border);
    vertical-align: middle; line-height: 1.32;
  }}
  .styled-table tbody tr:nth-child(odd) {{ background: var(--alt-row); }}
  .styled-table tbody tr:nth-child(even) {{ background: #ffffff; }}
  /* SMS Activity Card */
  .sms-body {{ justify-content: center; align-items: center; }}
  .sms-card {{
    width: 94%; background: var(--card-bg); border-radius: 14px;
    padding: 2.1rem 2rem; text-align: center;
  }}
  .sms-line1 {{ margin: 0 0 0.6rem 0; font-size: clamp(1.05rem, 2.0vw, 1.7rem); }}
  .sms-phone {{ margin: 0 0 0.8rem 0; font-size: clamp(1.35rem, 2.6vw, 2.2rem); font-weight: 700; }}
  .sms-line3 {{ margin: 0; font-size: clamp(0.95rem, 1.8vw, 1.5rem); }}
  .results-pill {{
    align-self: flex-start; margin-left: 3%; margin-top: 0.65rem;
    font-size: clamp(0.75rem, 1.05vw, 0.92rem);
  }}
  /* Live Streamlit Iframe Slide */
  .streamlit-body {{ top: 19.5%; bottom: 6.5%; }}
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
# Main Conversion Pipeline
# ---------------------------------------------------------------------------
def convert_zip_to_pptx_and_html(
    zip_path: str,
    output_pptx: str,
    output_html: Optional[str] = None,
    template_pptx: Optional[str] = None,
    capture_live: bool = False,
):
    tmp_dir = tempfile.mkdtemp(prefix="gh_md_to_pptx_")
    try:
        _, docs_dir, mkdocs_path = extract_zip_and_locate_project(zip_path, tmp_dir)
        site_name, nav_files = load_mkdocs_config(mkdocs_path, docs_dir)
        prs, using_template = initialize_presentation(template_pptx)

        parsed_pages: List[Tuple[str, Dict]] = []
        iframe_by_prefix: Dict[str, str] = {}
        for rel_md in nav_files:
            md_path = docs_dir / rel_md
            if not md_path.exists():
                continue
            parsed = parse_markdown_sections(md_path.read_text(encoding="utf-8", errors="replace"))
            parsed_pages.append((rel_md, parsed))
            if parsed["iframe_url"]:
                prefix = re.sub(r"_2\.md$", "", rel_md)
                iframe_by_prefix[prefix] = parsed["iframe_url"]

        slide_counter = 1
        html_slides: List[Dict] = []

        for idx, (rel_md, parsed) in enumerate(parsed_pages):
            # 1. Cover slide from index.md + mkdocs.yml site_name
            if idx == 0 and Path(rel_md).name.lower() == "index.md":
                first_sec = parsed["sections"][0]
                subtitle = ""
                remaining_blocks = []
                for b in first_sec["blocks"]:
                    if not subtitle and b["type"] == "paragraph" and b["text"].startswith("**"):
                        subtitle = b["text"]
                    else:
                        remaining_blocks.append(b)

                hero_img = None
                for candidate_name in ("evacuated.png", "hero.png", "cover.png"):
                    if (docs_dir / candidate_name).exists():
                        hero_img = docs_dir / candidate_name
                        break

                build_cover_slide(prs, using_template, site_name, subtitle, hero_img, html_slides)
                slide_counter += 1

                if remaining_blocks:
                    slide_counter = build_content_slides(
                        prs, using_template, slide_counter, parsed["page_title"], remaining_blocks, docs_dir, html_slides
                    )

                for extra_sec in parsed["sections"][1:]:
                    slide_counter = build_content_slides(
                        prs, using_template, slide_counter, extra_sec["title"], extra_sec["blocks"], docs_dir, html_slides
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

            # 3. Standard content sections + optional SMS Activity Callout slide
            for sec in parsed["sections"]:
                cleaned_blocks = [
                    b for b in sec["blocks"]
                    if not (
                        b["type"] == "paragraph"
                        and re.match(r"^To start the activity,\s*send\b", b["text"], re.IGNORECASE)
                    )
                ]
                if cleaned_blocks:
                    slide_counter = build_content_slides(
                        prs,
                        using_template,
                        slide_counter,
                        sec["title"] or parsed["page_title"],
                        cleaned_blocks,
                        docs_dir,
                        html_slides,
                    )

            if parsed["sms_callout"]:
                prefix = re.sub(r"_1\.md$", "", rel_md)
                results_url = iframe_by_prefix.get(prefix)
                build_sms_activity_slide(
                    prs,
                    using_template,
                    slide_counter,
                    parsed["page_title"],
                    parsed["sms_callout"],
                    results_url,
                    html_slides,
                )
                slide_counter += 1

        out_pptx_path = Path(output_pptx)
        out_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out_pptx_path))

        out_html_path = Path(output_html) if output_html else out_pptx_path.with_suffix(".html")
        render_html_deck(html_slides, site_name, out_html_path)

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
        "--capture-live",
        action="store_true",
        help="Capture fresh screenshots of embedded Streamlit iframes for the PPTX via Playwright",
    )
    args = parser.parse_args()
    convert_zip_to_pptx_and_html(args.zip_file, args.output, args.html, args.template, args.capture_live)


if __name__ == "__main__":
    main()
