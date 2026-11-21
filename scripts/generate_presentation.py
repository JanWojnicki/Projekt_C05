#!/usr/bin/env python3
"""Generate a complete, visually polished 16:9 PowerPoint presentation for Project C05."""
from pathlib import Path
import re
from PIL import Image
from pygments import lex
from pygments.lexers import CppLexer, BashLexer
from pygments.token import Token
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


# =============================================================================
# DESIGN SYSTEM — HPC / ACADEMIC ENGINEERING PALETTE (16:9 WIDESCREEN)
# =============================================================================
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT_MAIN = "Segoe UI"
FONT_SEMI = "Segoe UI Semibold"
FONT_CODE = "Consolas"

# Dark theme (Title, Live Demo, Closing)
C_DARK_BG     = RGBColor(15, 23, 42)     # #0F172A Slate 900
C_DARK_CARD   = RGBColor(30, 41, 59)     # #1E293B Slate 800
C_DARK_BORDER = RGBColor(51, 65, 85)     # #334155 Slate 700
C_DARK_TEXT   = RGBColor(241, 245, 249)  # #F1F5F9 Slate 100
C_DARK_MUTED  = RGBColor(148, 163, 184)  # #94A3B8 Slate 400

# Light theme (Technical & Results slides — blends seamlessly with white plots)
C_LIGHT_BG    = RGBColor(248, 250, 252)  # #F8FAFC Slate 50
C_WHITE       = RGBColor(255, 255, 255)
C_BORDER      = RGBColor(226, 232, 240)  # #E2E8F0 Slate 200
C_BORDER_STR  = RGBColor(203, 213, 225)  # #CBD5E1 Slate 300
C_NAVY        = RGBColor(15, 23, 42)     # #0F172A
C_TEXT        = RGBColor(30, 41, 59)     # #1E293B
C_MUTED       = RGBColor(100, 116, 139)  # #64748B

# Semantic backend & accent colors
C_SKY         = RGBColor(56, 189, 248)   # #38BDF8 accent on dark
C_BLUE        = RGBColor(2, 132, 199)    # #0284C7 OpenMP
C_BLUE_BG     = RGBColor(224, 242, 254)  # #E0F2FE
C_VIOLET      = RGBColor(124, 58, 237)   # #7C3AED MPI
C_VIOLET_BG   = RGBColor(237, 233, 254)  # #EDE9FE
C_GREEN       = RGBColor(5, 150, 105)    # #059669 CUDA
C_GREEN_LT    = RGBColor(52, 211, 153)   # #34D399 CUDA on dark
C_GREEN_BG    = RGBColor(209, 250, 229)  # #D1FAE5
C_AMBER       = RGBColor(217, 119, 6)    # #D97706 Highlight / Incremental
C_AMBER_LT    = RGBColor(251, 191, 36)   # #FBBF24 Amber on dark
C_AMBER_BG    = RGBColor(254, 243, 199)  # #FEF3C7
C_RED         = RGBColor(220, 38, 38)    # #DC2626 Contention / Remove
C_RED_BG      = RGBColor(254, 226, 226)  # #FEE2E2
C_SLATE       = RGBColor(71, 85, 105)    # #475569 Reference
C_SLATE_BG    = RGBColor(241, 245, 249)  # #F1F5F9

# Syntax highlighting colors (One Dark / VS Code inspired on #0F172A / #1E293B)
SYNTAX_COLORS = {
    Token.Keyword:              RGBColor(198, 120, 221),  # purple
    Token.Keyword.Type:         RGBColor(86, 182, 194),   # cyan
    Token.Name.Builtin:         RGBColor(86, 182, 194),
    Token.Name.Function:        RGBColor(97, 175, 239),   # blue
    Token.Name.Class:           RGBColor(229, 192, 123),  # gold
    Token.Comment:              RGBColor(127, 140, 163),  # muted gray-blue
    Token.Comment.Preproc:      RGBColor(244, 114, 182),  # pink (#pragma / #include)
    Token.Comment.PreprocFile:  RGBColor(152, 195, 121),  # green
    Token.Literal.String:       RGBColor(152, 195, 121),  # green
    Token.Literal.Number:       RGBColor(209, 154, 102),  # orange
    Token.Operator:             RGBColor(148, 163, 184),
    Token.Punctuation:          RGBColor(203, 213, 225),
}
SYNTAX_DEFAULT = RGBColor(226, 232, 240)
CUDA_KEYWORDS = {
    "__global__", "__device__", "__shared__", "__syncthreads",
    "atomicAdd", "atomicSub", "blockIdx", "blockDim", "threadIdx", "gridDim",
    "cudaOccupancyMaxActiveBlocksPerMultiprocessor", "cudaDeviceSynchronize",
    "cudaEventRecord", "cudaEventElapsedTime", "cudaMemcpy",
    "MPI_Init", "MPI_Bcast", "MPI_Isend", "MPI_Waitall", "MPI_Recv",
    "MPI_Barrier", "MPI_Reduce", "MPI_Gatherv", "MPI_Wtime",
    "omp_get_num_threads", "omp_get_thread_num", "omp_set_num_threads",
}


# =============================================================================
# LOW-LEVEL DRAWING HELPERS
# =============================================================================
def set_slide_bg(slide, color: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_cell_border(cell, color_hex="E2E8F0", width_pt=0.5, sides=("top", "bottom", "left", "right")):
    tcPr = cell._tc.get_or_add_tcPr()
    w_emu = str(int(width_pt * 12700))
    side_map = {"top": "lnT", "bottom": "lnB", "left": "lnL", "right": "lnR"}
    for side in ("top", "bottom", "left", "right"):
        tag = f"a:{side_map[side]}"
        for existing in tcPr.findall(qn(tag)):
            tcPr.remove(existing)
        ln = OxmlElement(tag)
        ln.set("w", w_emu)
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        if side in sides:
            solidFill = OxmlElement("a:solidFill")
            srgbClr = OxmlElement("a:srgbClr")
            srgbClr.set("val", color_hex)
            solidFill.append(srgbClr)
            ln.append(solidFill)
        else:
            ln.append(OxmlElement("a:noFill"))
        tcPr.append(ln)


def set_cell_margins(cell, left=0.08, right=0.08, top=0.04, bottom=0.04):
    cell.margin_left = Inches(left)
    cell.margin_right = Inches(right)
    cell.margin_top = Inches(top)
    cell.margin_bottom = Inches(bottom)


def add_rect(slide, left, top, width, height, fill_color, border_color=None,
             border_pt=1.0, rounded=False, radius=0.06):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    if rounded and shape.adjustments:
        shape.adjustments[0] = radius
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color is not None:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_card(slide, left, top, width, height, title=None, accent_color=None,
             bg_color=C_WHITE, border_color=C_BORDER, title_color=C_NAVY,
             title_size=12.5, radius=0.04):
    card = add_rect(slide, left, top, width, height, bg_color, border_color=border_color,
                    border_pt=1.0, rounded=True, radius=radius)
    if accent_color is not None:
        # Top thin accent bar
        bar = add_rect(slide, left + 0.02, top, width - 0.04, 0.06, accent_color)
        bar.line.fill.background()
    if title:
        tb = slide.shapes.add_textbox(Inches(left + 0.16), Inches(top + (0.11 if accent_color else 0.08)),
                                      Inches(width - 0.32), Inches(0.32))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        r.font.name = FONT_SEMI
        r.font.size = Pt(title_size)
        r.font.bold = True
        r.font.color.rgb = title_color
    return card


def add_textbox(slide, left, top, width, height, text="", size=12, color=C_TEXT,
                bold=False, italic=False, font=FONT_MAIN, align=PP_ALIGN.LEFT,
                valign=MSO_ANCHOR.TOP, line_spacing=1.18):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if text:
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = text
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_bullets(slide, left, top, width, height, items, size=11.5, color=C_TEXT,
                bullet_color=C_BLUE, space_after=5.5, line_spacing=1.16, bullet_char="▸ "):
    """items: list of str or tuple(bold_prefix, regular_text)."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        rb = p.add_run()
        rb.text = bullet_char
        rb.font.name = FONT_SEMI
        rb.font.size = Pt(size)
        rb.font.bold = True
        rb.font.color.rgb = bullet_color
        if isinstance(item, tuple):
            prefix, body = item
            rp = p.add_run()
            rp.text = prefix
            rp.font.name = FONT_SEMI
            rp.font.size = Pt(size)
            rp.font.bold = True
            rp.font.color.rgb = C_NAVY if color == C_TEXT else C_DARK_TEXT
            rt = p.add_run()
            rt.text = body
            rt.font.name = FONT_MAIN
            rt.font.size = Pt(size)
            rt.font.color.rgb = color
        else:
            rt = p.add_run()
            rt.text = item
            rt.font.name = FONT_MAIN
            rt.font.size = Pt(size)
            rt.font.color.rgb = color
    return tb


def add_badge(slide, left, top, width, height, text, bg_color, text_color, size=9.5):
    shape = add_rect(slide, left, top, width, height, bg_color, rounded=True, radius=0.45)
    tf = shape.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT_SEMI
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = text_color
    return shape


def add_header(slide, section_tag, title, subtitle=None, tag_color=C_BLUE,
               tag_bg=C_BLUE_BG, dark=False):
    # Top slim accent line
    add_rect(slide, 0, 0, 13.333, 0.055, tag_color)
    # Section pill
    tag_w = max(1.45, len(section_tag) * 0.095 + 0.30)
    add_badge(slide, 0.55, 0.22, tag_w, 0.26, section_tag.upper(), tag_bg, tag_color, size=8.5)
    # Title
    add_textbox(slide, 0.55, 0.50, 12.23, 0.44, title, size=22.5,
                color=C_DARK_TEXT if dark else C_NAVY, bold=True, font=FONT_SEMI)
    if subtitle:
        add_textbox(slide, 0.55, 0.94, 12.23, 0.28, subtitle, size=11.5,
                    color=C_DARK_MUTED if dark else C_MUTED, italic=False)
    # Subtle divider line
    div_y = 1.24 if subtitle else 1.00
    add_rect(slide, 0.55, div_y, 12.23, 0.015, C_DARK_BORDER if dark else C_BORDER)


def add_footer(slide, page_num, total_pages, dark=False):
    line_color = C_DARK_BORDER if dark else C_BORDER
    txt_color = C_DARK_MUTED if dark else C_MUTED
    add_rect(slide, 0.55, 7.06, 12.23, 0.012, line_color)
    add_textbox(slide, 0.55, 7.12, 9.8, 0.25,
                "Projekt C05: Entropia i segmentacja plików binarnych (PRiR)  •  Jan Wojnicki (147020), Adrian Sajdak (146992)",
                size=8.8, color=txt_color)
    add_textbox(slide, 11.4, 7.12, 1.38, 0.25, f"{page_num} / {total_pages}",
                size=9.2, color=txt_color, bold=True, align=PP_ALIGN.RIGHT)


def add_notes(slide, text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = text.strip()


def _token_color(ttype, value):
    if value.strip() in CUDA_KEYWORDS:
        return RGBColor(52, 211, 153)  # bright emerald for CUDA/MPI/OpenMP API
    t = ttype
    while t is not None:
        if t in SYNTAX_COLORS:
            return SYNTAX_COLORS[t]
        t = t.parent
    return SYNTAX_DEFAULT


def add_code_block(slide, left, top, width, height, code, filename=None,
                   lang="cpp", font_size=9.3, start_line=1, accent_color=C_SKY,
                   line_spacing=1.05):
    # Outer dark container
    add_rect(slide, left, top, width, height, C_DARK_BG, border_color=C_DARK_BORDER,
             border_pt=1.0, rounded=True, radius=0.03)
    header_h = 0.0
    if filename:
        header_h = 0.28
        add_rect(slide, left, top, width, header_h, C_DARK_CARD, border_color=C_DARK_BORDER,
                 border_pt=0.75)
        for idx, dot_c in enumerate((RGBColor(239, 68, 68), RGBColor(245, 158, 11), RGBColor(16, 185, 129))):
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left + 0.12 + idx * 0.16),
                                         Inches(top + 0.09), Inches(0.095), Inches(0.095))
            dot.fill.solid()
            dot.fill.fore_color.rgb = dot_c
            dot.line.fill.background()
        add_textbox(slide, left + 0.64, top + 0.04, width - 0.8, 0.21, filename,
                    size=9.0, color=accent_color, bold=True, font=FONT_CODE)

    lines = code.strip("\n").splitlines()
    # Auto-clamp font size if lines would otherwise overflow the box height
    avail_h = height - header_h - 0.16
    est_line_h = (font_size * 1.24 * line_spacing) / 72.0
    if lines and len(lines) * est_line_h > avail_h:
        font_size = max(6.8, (avail_h / len(lines)) * 72.0 / (1.24 * line_spacing))

    tb = slide.shapes.add_textbox(Inches(left + 0.12), Inches(top + header_h + 0.06),
                                  Inches(width - 0.22), Inches(height - header_h - 0.10))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    lexer = CppLexer() if lang == "cpp" else (BashLexer() if lang == "bash" else None)
    for line_idx, raw_line in enumerate(lines):
        p = tf.paragraphs[0] if line_idx == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        p.space_after = Pt(0)
        if start_line is not None:
            rn = p.add_run()
            rn.text = f"{start_line + line_idx:3d} │ "
            rn.font.name = FONT_CODE
            rn.font.size = Pt(font_size)
            rn.font.color.rgb = RGBColor(100, 116, 139)
        if lang == "ps1":
            stripped = raw_line.lstrip()
            if stripped.startswith("#"):
                r = p.add_run()
                r.text = raw_line
                r.font.name = FONT_CODE
                r.font.size = Pt(font_size)
                r.font.color.rgb = RGBColor(127, 140, 163)
            else:
                parts = raw_line.split(" ")
                for pidx, part in enumerate(parts):
                    if pidx > 0:
                        sp = p.add_run()
                        sp.text = " "
                        sp.font.name = FONT_CODE
                        sp.font.size = Pt(font_size)
                    r = p.add_run()
                    r.text = part
                    r.font.name = FONT_CODE
                    r.font.size = Pt(font_size)
                    if part.startswith(".\\"):
                        r.font.color.rgb = C_GREEN_LT
                        r.font.bold = True
                    elif part.startswith("-"):
                        r.font.color.rgb = C_AMBER_LT
                    else:
                        r.font.color.rgb = C_SKY
            continue
        tokens = list(lex(raw_line, lexer))
        for ttype, value in tokens:
            val = value.rstrip("\r\n")
            if not val:
                continue
            r = p.add_run()
            r.text = val
            r.font.name = FONT_CODE
            r.font.size = Pt(font_size)
            r.font.color.rgb = _token_color(ttype, val)
            if ttype in Token.Keyword or val.strip() in CUDA_KEYWORDS:
                r.font.bold = True


def add_picture_fit(slide, img_path, left, top, max_w, max_h, card=True, caption=None):
    """Place an image inside (left, top, max_w, max_h) preserving exact aspect ratio."""
    cap_h = 0.28 if caption else 0.0
    pad = 0.08 if card else 0.0
    if card:
        add_rect(slide, left, top, max_w, max_h, C_WHITE, border_color=C_BORDER,
                 border_pt=1.0, rounded=True, radius=0.03)
    avail_w = max_w - 2 * pad
    avail_h = max_h - 2 * pad - cap_h
    with Image.open(img_path) as im:
        iw, ih = im.size
    ratio = iw / ih
    if avail_w / avail_h > ratio:
        draw_h = avail_h
        draw_w = draw_h * ratio
    else:
        draw_w = avail_w
        draw_h = draw_w / ratio
    x = left + (max_w - draw_w) / 2
    y = top + pad + (avail_h - draw_h) / 2
    slide.shapes.add_picture(str(img_path), Inches(x), Inches(y), width=Inches(draw_w), height=Inches(draw_h))
    if caption:
        add_textbox(slide, left + 0.12, top + max_h - cap_h + 0.02, max_w - 0.24, cap_h - 0.04,
                    caption, size=8.8, color=C_MUTED, italic=True, align=PP_ALIGN.CENTER)


def add_table(slide, left, top, width, height, headers, rows, col_widths=None,
              right_cols=(), header_bg=C_NAVY, font_size=9.5, header_size=9.5,
              highlight_rows=()):
    table_shape = slide.shapes.add_table(len(rows) + 1, len(headers),
                                         Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = table_shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    if col_widths:
        for idx, cw in enumerate(col_widths):
            tbl.columns[idx].width = Inches(cw)

    # Header row
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        set_cell_margins(cell, left=0.07, right=0.07, top=0.04, bottom=0.04)
        set_cell_border(cell, color_hex="334155", width_pt=0.6, sides=("bottom",))
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT if j in right_cols else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = h
        r.font.name = FONT_SEMI
        r.font.size = Pt(header_size)
        r.font.bold = True
        r.font.color.rgb = C_WHITE

    # Body rows
    for i, row in enumerate(rows):
        is_hl = i in highlight_rows
        bg = C_GREEN_BG if is_hl else (C_SLATE_BG if i % 2 == 1 else C_WHITE)
        for j, val in enumerate(row):
            cell = tbl.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_cell_margins(cell, left=0.07, right=0.07, top=0.03, bottom=0.03)
            set_cell_border(cell, color_hex="E2E8F0", width_pt=0.5, sides=("bottom",))
            tf = cell.text_frame
            raw_lines = str(val).split("\n")
            for lidx, raw_line in enumerate(raw_lines):
                p = tf.paragraphs[0] if lidx == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.RIGHT if j in right_cols else PP_ALIGN.LEFT
                for seg in re.split(r"(\*\*.*?\*\*)", raw_line):
                    if not seg:
                        continue
                    seg_bold = is_hl or (seg.startswith("**") and seg.endswith("**"))
                    clean_seg = seg[2:-2] if (seg.startswith("**") and seg.endswith("**")) else seg
                    r = p.add_run()
                    r.text = clean_seg
                    r.font.name = FONT_SEMI if seg_bold else FONT_MAIN
                    r.font.size = Pt(font_size)
                    r.font.bold = seg_bold
                    r.font.color.rgb = C_NAVY if seg_bold else C_TEXT
    return tbl


def add_arrow_right(slide, left, top, width, height, color):
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left), Inches(top), Inches(width), Inches(height))
    arr.fill.solid()
    arr.fill.fore_color.rgb = color
    arr.line.fill.background()
    arr.shadow.inherit = False
    return arr


# =============================================================================
# SLIDE BUILDERS (17 SLIDES)
# =============================================================================
def build_presentation(root: Path, output_path: Path):
    plots = root / "results" / "plots"
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    TOTAL = 17

    # -------------------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE (DARK HPC HERO)
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_DARK_BG)
    # Decorative geometric strips
    add_rect(s, 0, 0, 13.333, 0.09, C_SKY)
    add_rect(s, 0, 0.09, 0.14, 7.41, C_BLUE)
    # Subtle right-side decorative bars suggesting entropy histogram
    bar_heights = [0.6, 1.1, 1.9, 1.5, 2.7, 3.4, 3.1, 2.2, 0.5, 0.4, 2.8, 3.6, 3.5, 1.3]
    for idx, bh in enumerate(bar_heights):
        bx = 9.45 + idx * 0.25
        bc = C_GREEN if bh > 2.6 else (C_BLUE if bh > 1.2 else C_DARK_CARD)
        bar = add_rect(s, bx, 4.75 - bh, 0.16, bh, bc, rounded=True, radius=0.15)
        bar.fill.solid()

    add_badge(s, 0.72, 0.65, 4.85, 0.34,
              "PROGRAMOWANIE RÓWNOLEGŁE I ROZPROSZONE  •  PROJEKT C05",
              C_DARK_CARD, C_SKY, size=9.5)

    add_textbox(s, 0.72, 1.25, 8.5, 1.45,
                "Entropia i segmentacja\nplików binarnych",
                size=42, color=C_DARK_TEXT, bold=True, font=FONT_SEMI, line_spacing=1.08)

    add_textbox(s, 0.72, 2.88, 8.6, 0.95,
                "Analiza porównawcza implementacji C++20 Reference, OpenMP, MPI oraz NVIDIA CUDA\n"
                "z przyrostową aktualizacją przesuwnego histogramu (zakres zespołu dwuosobowego)",
                size=14.8, color=C_DARK_MUTED, line_spacing=1.32)

    # Author card
    add_rect(s, 0.72, 4.05, 8.2, 0.88, C_DARK_CARD, border_color=C_DARK_BORDER, rounded=True, radius=0.08)
    add_textbox(s, 0.95, 4.16, 3.8, 0.28, "ZESPÓŁ AUTORSKI (2 OSOBY)",
                size=9, color=C_SKY, bold=True, font=FONT_SEMI)
    add_textbox(s, 0.95, 4.44, 7.8, 0.36,
                "Jan Wojnicki (nr albumu: 147020)     •     Adrian Sajdak (nr albumu: 146992)",
                size=14, color=C_DARK_TEXT, bold=True, font=FONT_SEMI)

    # 4 KPI highlight cards along the bottom
    kpis = [
        ("4 × 2", "Technologii i trybów", "Ref, OpenMP, MPI, CUDA\n× Naive + Incremental (100% kod własny)", C_SKY),
        ("129,8×", "Przyspieszenie rdzenia GPU", "GTX 1660 (10,02 ms vs 1300 ms CPU)\ndo 140,2 GB/s przepustowości logicznej", C_GREEN_LT),
        ("92,4%", "Efektywność CPU (6 rdzeni)", "Prywatne histogramy L1 w OpenMP\ni podział z zakładką halo w MPI", C_BLUE),
        ("4 / 4 PASS", "Pełna zgodność CTest", "Błąd CPU = 0,0 (bitowa zgodność)\nBłąd GPU ≤ 1,69×10⁻¹⁴  •  0 FP granic", C_AMBER_LT),
    ]
    for idx, (val, label, desc, col) in enumerate(kpis):
        kx = 0.72 + idx * 3.03
        add_rect(s, kx, 5.28, 2.86, 1.65, C_DARK_CARD, border_color=C_DARK_BORDER, rounded=True, radius=0.06)
        add_rect(s, kx + 0.04, 5.28, 2.78, 0.06, col)
        add_textbox(s, kx + 0.16, 5.42, 2.54, 0.45, val, size=24, color=col, bold=True, font=FONT_SEMI)
        add_textbox(s, kx + 0.16, 5.90, 2.54, 0.28, label, size=11, color=C_DARK_TEXT, bold=True, font=FONT_SEMI)
        add_textbox(s, kx + 0.16, 6.20, 2.54, 0.65, desc, size=9.2, color=C_DARK_MUTED, line_spacing=1.18)

    add_notes(s, (
        "SLAJD 1 (Otwarcie — Jan & Adrian):\n"
        "• Przedstawiamy projekt C05: Entropia i segmentacja plików binarnych w wersji rozszerzonej dla zespołu 2-osobowego.\n"
        "• Zrealizowaliśmy 4 pełne implementacje w C++20 i CUDA C++ (Reference, OpenMP, MPI, NVIDIA CUDA), a każdą z nich w dwóch wariantach algorytmicznych: naiwnym (O(K*W)) i przyrostowym (O(W+K*S)).\n"
        "• Cały rdzeń obliczeniowy to 100% kod własny; na koniec uruchomimy na żywo plik demo.ps1, który skompiluje projekt i uruchomi wszystkie 4 technologie oraz detekcję granic."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 2: PROBLEM & MATHEMATICAL FORMULATION
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "01  •  Problem i matematyka",
               "Cel projektu i sformułowanie matematyczne",
               "Prześwietlanie struktury statystycznej pliku binarnego metodą przesuwnego okna entropii Shannona")

    # Left column: Motivation + Entropy scale + Domain note
    add_card(s, 0.55, 1.42, 5.75, 3.65, "Skala entropii bajtowej i zastosowania", accent_color=C_BLUE)
    scale_rows = [
        ("0,0 bit/B", "Padding / zera (0x00)", "Stały bajt — zero różnorodności (k = 1)", C_SLATE, C_SLATE_BG),
        ("2,0 – 4,5 bit/B", "Dane strukturalne / ASCII", "Ograniczony alfabet (np. tekst 23-znakowy: log₂23 ≈ 4,52)", C_BLUE, C_BLUE_BG),
        ("5,5 – 6,8 bit/B", "Kod maszynowy (.text)", "Instrukcje CPU o nierównomiernym rozkładzie bajtów", C_VIOLET, C_VIOLET_BG),
        ("7,9 – 8,0 bit/B", "Kompresja / Szyfrowanie", "Rozkład bliski jednostajnemu na 256 symbolach (k = 256)", C_GREEN, C_GREEN_BG),
    ]
    for idx, (rng, kind, desc, col, bg) in enumerate(scale_rows):
        ry = 1.94 + idx * 0.74
        add_badge(s, 0.73, ry + 0.06, 1.45, 0.36, rng, bg, col, size=9.5)
        add_textbox(s, 2.30, ry, 3.80, 0.25, kind, size=11, color=C_NAVY, bold=True, font=FONT_SEMI)
        add_textbox(s, 2.30, ry + 0.25, 3.80, 0.38, desc, size=9.5, color=C_MUTED, line_spacing=1.12)

    # Domain warning card
    add_card(s, 0.55, 5.22, 5.75, 1.66, "Ważna uwaga domenowa (interpretacja)",
             accent_color=C_AMBER, bg_color=C_AMBER_BG, border_color=C_AMBER)
    add_textbox(s, 0.73, 5.68, 5.38, 1.12,
                "Wysoka entropia (~8 bit/bajt) NIE stanowi samodzielnego dowodu złośliwości pliku ani użycia szyfrowania! "
                "Identyczny rozkład jednostajny mają legalne archiwa skompresowane (ZIP, GZ, XZ), multimedia (JPEG, MP4) "
                "oraz strumienie pseudolosowe. Entropia to szybki radar wskazujący granice obszarów o odmiennej strukturze.",
                size=10.5, color=C_NAVY, line_spacing=1.22)

    # Right column: 3 Mathematical Formulation Cards
    # Card 1: Window layout
    add_card(s, 6.55, 1.42, 6.23, 1.66, "1. Układ pełnych okien i obsługa końcówki pliku", accent_color=C_NAVY)
    add_rect(s, 6.73, 1.88, 5.87, 0.42, C_SLATE_BG, border_color=C_BORDER, rounded=True, radius=0.1)
    add_textbox(s, 6.78, 1.94, 5.77, 0.30,
                "K = 0 (gdy N < W)   |   K = 1 + ⌊(N − W) / S⌋ (gdy N ≥ W)",
                size=10.3, color=C_NAVY, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_textbox(s, 6.73, 2.38, 5.87, 0.64,
                "Okno i ∈ {0..K−1} obejmuje przedział [i·S, i·S + W). Przetwarzamy wyłącznie pełne okna "
                "(krótszy ogon zaniżałby estymator entropii). Pominięty ogon (omitted_trailing_bytes) "
                "oraz przerwy dla S > W (uncovered_gap_bytes) są zawsze jawnie raportowane na stderr.",
                size=9.8, color=C_TEXT, line_spacing=1.15)

    # Card 2: Shannon entropy
    add_card(s, 6.55, 3.24, 6.23, 1.66, "2. Histogram 256-koszykowy i entropia Shannona", accent_color=C_GREEN)
    add_rect(s, 6.73, 3.70, 5.87, 0.42, C_SLATE_BG, border_color=C_BORDER, rounded=True, radius=0.1)
    add_textbox(s, 6.78, 3.76, 5.77, 0.30,
                "Hᵢ = −∑ (cᵢ[v] / W)·log₂(cᵢ[v] / W)   dla v ∈ {0..255} : cᵢ[v] > 0",
                size=9.8, color=C_GREEN, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_textbox(s, 6.73, 4.20, 5.87, 0.64,
                "Wartość Hᵢ ∈ [0, 8] bit/bajt obliczana jest w 64-bitowej arytmetyce IEEE-754 double. "
                "Koszyki zerowe (cᵢ[v] = 0) są jawnie pomijane zgodnie z lim (p → 0⁺) p·log₂(p) = 0, "
                "eliminując ryzyko wartości nieoznaczonych NaN (0 · −∞).",
                size=9.8, color=C_TEXT, line_spacing=1.15)

    # Card 3: Boundary detection & merging
    add_card(s, 6.55, 5.06, 6.23, 1.82, "3. Detekcja granic i scalanie zdarzeń (--events)", accent_color=C_VIOLET)
    add_rect(s, 6.73, 5.52, 5.87, 0.42, C_SLATE_BG, border_color=C_BORDER, rounded=True, radius=0.1)
    add_textbox(s, 6.78, 5.58, 5.77, 0.30,
                "|Hᵢ − Hᵢ₋₁| ≥ T (i ≥ 1)  ⇒  x̂ = ⌊((i_first−1)·S + i_last·S + W) / 2⌋",
                size=9.5, color=C_VIOLET, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_textbox(s, 6.73, 6.02, 5.87, 0.78,
                "Przy oknach nakładających się (S < W) jedna fizyczna granica wyzwala ~W/S kolejnych flag. "
                "Procedura merge_boundaries() łączy maksymalne spójne ciągi flag [i_first, i_last] w pojedyncze "
                "zdarzenie zmiany BoundaryEvent i szacuje środek przejścia x̂ (estimated_offset).",
                size=9.8, color=C_TEXT, line_spacing=1.15)

    add_footer(s, 2, TOTAL)
    add_notes(s, (
        "SLAJD 2 (Matematyka i cel — Jan):\n"
        "• Podkreślić, dlaczego przetwarzamy tylko pełne okna: krótsze okno końcowe miałoby mniejszą próbę W i silnie ujemny bias Millera-Madowa, co wywołałoby fałszywą granicę na końcu pliku.\n"
        "• Zwrócić uwagę na nierówność słabą >= T (zweryfikowaną w testach na przykładzie dokładnie reprezentowalnych H=0 i H=1 przy T=1.0).\n"
        "• Mechanizm --events (merge_boundaries) rozwiązuje praktyczny problem wielu sąsiednich alarmów na jednej granicy przy S < W."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 3: ARCHITECTURE, CLI & OWN CODE CONTRIBUTION
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "02  •  Architektura systemu",
               "Architektura projektu, jednolity interfejs CLI i kod własny",
               "Wspólny rdzeń C++20 (libcommon.a) + 4 niezależne backendy + zautomatyzowany potok testów i benchmarków")

    # Top Architectural Diagram
    add_card(s, 0.55, 1.40, 12.23, 2.55, "Struktura modułowa (CMake 3.24+ / Ninja / C++20 / CUDA 20)", accent_color=C_NAVY)

    # Box 1: common/
    add_rect(s, 0.78, 1.92, 3.15, 1.82, C_SLATE_BG, border_color=C_NAVY, border_pt=1.5, rounded=True, radius=0.05)
    add_textbox(s, 0.92, 1.98, 2.88, 0.28, "common/  (libcommon.a)", size=11.5, color=C_NAVY, bold=True, font=FONT_CODE)
    add_bullets(s, 0.92, 2.28, 2.88, 1.40, [
        ("types.hpp: ", "enum Mode, Options, Boundary, BoundaryEvent, WindowRange"),
        ("io.cpp: ", "CLI, read_binary, block_range, detect/merge_boundaries, write_outputs"),
        ("compute.cpp: ", "shannon_entropy, compute_naive/incremental"),
    ], size=8.8, space_after=3.5, bullet_color=C_NAVY)

    add_arrow_right(s, 4.02, 2.68, 0.38, 0.28, C_NAVY)

    # Box 2: 4 executables in src/
    backends = [
        ("entropy_ref", "src/reference.cpp", "1 wątek CPU (wzorzec)", C_SLATE, C_SLATE_BG),
        ("entropy_omp", "src/openmp.cpp", "OpenMP (prywatny L1 hist)", C_BLUE, C_BLUE_BG),
        ("entropy_mpi", "src/mpi.cpp", "MPI (podział + halo W−S)", C_VIOLET, C_VIOLET_BG),
        ("entropy_cuda", "src/cuda_entropy.cu", "CUDA (shared mem + tree)", C_GREEN, C_GREEN_BG),
    ]
    for idx, (exe, src, desc, col, bg) in enumerate(backends):
        bx = 4.48 + (idx % 2) * 2.22
        by = 1.92 + (idx // 2) * 0.95
        add_rect(s, bx, by, 2.10, 0.86, bg, border_color=col, border_pt=1.2, rounded=True, radius=0.06)
        add_textbox(s, bx + 0.10, by + 0.06, 1.90, 0.24, exe, size=10.5, color=col, bold=True, font=FONT_CODE)
        add_textbox(s, bx + 0.10, by + 0.30, 1.90, 0.22, src, size=8.5, color=C_NAVY, font=FONT_CODE)
        add_textbox(s, bx + 0.10, by + 0.53, 1.90, 0.26, desc, size=8.3, color=C_MUTED)

    add_arrow_right(s, 8.95, 2.68, 0.38, 0.28, C_GREEN)

    # Box 3: Verification & Experiments
    add_rect(s, 9.42, 1.92, 3.16, 1.82, C_SLATE_BG, border_color=C_GREEN, border_pt=1.2, rounded=True, radius=0.05)
    add_textbox(s, 9.56, 1.98, 2.88, 0.28, "tests/  &  scripts/", size=11.5, color=C_GREEN, bold=True, font=FONT_CODE)
    add_bullets(s, 9.56, 2.28, 2.88, 1.40, [
        ("CTest (4 zestawy): ", "check_cuda_model, verify_correctness, verify_mpi, verify_cuda"),
        ("Dane & Pomiary: ", "generate_datasets.py, benchmark_runner.py"),
        ("Wizualizacja: ", "plot_results.py (PNG/PDF), plot_entropy.py, demo.ps1"),
    ], size=8.8, space_after=3.5, bullet_color=C_GREEN)

    # Bottom Left: Unified CLI code block
    cli_code = (
        "# Identyczny zestaw flag dla wszystkich 4 backendów:\n"
        "./build/entropy_ref  --input PLIK --window W --step S --threshold T \\\n"
        "                     --output WYNIK.csv [--events ZMIANY.csv] --mode naive|incremental\n"
        "./build/entropy_omp  ... --threads 12 --mode incremental\n"
        "mpiexec -n 6 ./build/entropy_mpi ... --mode incremental\n"
        "./build/entropy_cuda ... --threads 256 --mode incremental\n"
        "\n"
        "# Kompilacja Release (0 ostrzeżeń, bez --use_fast_math dla dokładności FP64):\n"
        "cmake -S . -B build -GNinja -DCMAKE_BUILD_TYPE=Release \\\n"
        "      -DCMAKE_CUDA_ARCHITECTURES=75 -DC05_REQUIRE_ALL=ON && cmake --build build -j"
    )
    add_code_block(s, 0.55, 4.12, 7.65, 2.78, cli_code,
                   filename="CLI & Build (CMakeLists.txt + common/io.cpp)", lang="bash",
                   font_size=8.8, start_line=None, accent_color=C_SKY)

    # Bottom Right: Key engineering highlights & own code contribution
    add_card(s, 8.40, 4.12, 4.38, 2.78, "Kluczowe zasady inżynierskie i kod własny", accent_color=C_AMBER)
    add_bullets(s, 8.56, 4.58, 4.06, 2.22, [
        ("100% kod własny: ", "histogramy, tryb przyrostowy, entropia, redukcja drzewiasta GPU, podział halo MPI, detekcja i scalanie granic."),
        ("Biblioteki tylko pomocnicze: ", "OpenMP, OpenMPI, CUDA Runtime API, STL oraz Python do wykresów."),
        ("Single Source of Truth: ", "jedno miejsce progu (detect_boundaries), wspólne block_range() i write_outputs()."),
        ("Rygor CLI i bezpieczeństwa: ", "blokada nadpisania wejścia (same_path), walidacja NaN/uint32, pełna precyzja 17 cyfr w CSV."),
    ], size=9.6, space_after=5.0, bullet_color=C_AMBER)

    add_footer(s, 3, TOTAL)
    add_notes(s, (
        "SLAJD 3 (Architektura i CLI — Jan):\n"
        "• Podkreślić czystość architektury po refaktoringu: common/ stanowi bibliotekę statyczną libcommon.a linkowaną przez wszystkie 4 binaria.\n"
        "• Zwrócić uwagę, że w CMake celowo NIE używamy --use_fast_math w nvcc, żeby zachować pełną precyzję double IEEE-754.\n"
        "• Wszystkie programy przyjmują identyczne argumenty CLI, a błędy argumentów (np. --output wskazujący na --input, ujemny próg, NaN, window > uint32_max) są odrzucane."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 4: 2-PERSON SCOPE — NAIVE VS INCREMENTAL + CODE
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "03  •  Rozszerzenie 2-osobowe",
               "Wariant naiwny vs przyrostowa aktualizacja histogramu",
               "Redukcja złożoności z O(K·W) do O(W + K·2S) zaimplementowana spójnie we wszystkich 4 technologiach",
               tag_color=C_AMBER, tag_bg=C_AMBER_BG)

    # Top Visual Diagram of Sliding Window
    add_card(s, 0.55, 1.40, 12.23, 1.98, "Idea przesuwnego histogramu (przykład dla S = W / 4 — 75% nakładania się okien)",
             accent_color=C_AMBER)

    # Window i-1
    add_textbox(s, 0.78, 1.92, 1.45, 0.36, "Okno i − 1:", size=10.5, color=C_NAVY, bold=True, font=FONT_SEMI, valign=MSO_ANCHOR.MIDDLE)
    r1 = add_rect(s, 2.25, 1.92, 1.65, 0.42, C_RED_BG, border_color=C_RED, border_pt=1.5, rounded=True, radius=0.08)
    add_textbox(s, 2.25, 1.92, 1.65, 0.42, "−S opuszczające", size=9.8, color=C_RED, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 3.96, 1.92, 4.95, 0.42, C_BLUE_BG, border_color=C_BLUE, border_pt=1.5, rounded=True, radius=0.04)
    add_textbox(s, 3.96, 1.92, 4.95, 0.42, "Część wspólna (W − S = 75% okna)  —  NIE PRZELICZAMY OD NOWA!",
                size=10.0, color=C_BLUE, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Window i
    add_textbox(s, 0.78, 2.48, 1.45, 0.36, "Okno i:", size=10.5, color=C_NAVY, bold=True, font=FONT_SEMI, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 3.96, 2.48, 4.95, 0.42, C_BLUE_BG, border_color=C_BLUE, border_pt=1.5, rounded=True, radius=0.04)
    add_textbox(s, 3.96, 2.48, 4.95, 0.42, "Część wspólna (zachowana w histogramie)",
                size=10.0, color=C_BLUE, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 8.97, 2.48, 1.65, 0.42, C_GREEN_BG, border_color=C_GREEN, border_pt=1.5, rounded=True, radius=0.08)
    add_textbox(s, 8.97, 2.48, 1.65, 0.42, "+S wchodzące", size=9.8, color=C_GREEN, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Formula box on right side of diagram
    add_rect(s, 10.80, 1.92, 1.80, 0.98, C_DARK_BG, rounded=True, radius=0.08)
    add_textbox(s, 10.86, 1.98, 1.68, 0.30, "Redukcja operacji:", size=9.0, color=C_DARK_MUTED, align=PP_ALIGN.CENTER)
    add_textbox(s, 10.86, 2.25, 1.68, 0.34, "W / (2·S)", size=15.0, color=C_AMBER_LT, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER)
    add_textbox(s, 10.86, 2.60, 1.68, 0.26, "S=W/4 ⇒ 2× mniej", size=8.8, color=C_GREEN_LT, bold=True, align=PP_ALIGN.CENTER)

    add_textbox(s, 0.78, 3.00, 11.8, 0.28,
                "Aktualizacja:   H_hist[i] = H_hist[i − 1]  −  hist(S bajtów z lewej)  +  hist(S bajtów z prawej).   Dla S ≥ W automatyczny powrót do compute_naive().",
                size=9.8, color=C_MUTED, bold=True)

    # Bottom Left: Real code from common/compute.cpp
    code_compute = """double shannon_entropy(const Histogram& histogram, std::size_t window) {
    double entropy = 0, denominator = static_cast<double>(window);
    for (const auto count : histogram) {
        if (count != 0) {
            const double p = static_cast<double>(count) / denominator;
            entropy -= p * std::log2(p);
        }
    }
    return entropy;
}
void compute_incremental(std::span<const std::uint8_t> bytes, std::size_t window,
                         std::size_t step, std::span<double> output) {
    if (step >= window) return compute_naive(bytes, window, step, output);
    if (output.empty()) return;
    Histogram histogram{};
    for (std::size_t j = 0; j < window; ++j) ++histogram[bytes[j]];
    output[0] = shannon_entropy(histogram, window);
    for (std::size_t i = 1; i < output.size(); ++i) {
        const auto old_start = (i - 1) * step;
        for (std::size_t j = 0; j < step; ++j) {
            --histogram[bytes[old_start + j]];
            ++histogram[bytes[old_start + window + j]];
        }
        output[i] = shannon_entropy(histogram, window);
    }
}"""
    add_code_block(s, 0.55, 3.50, 7.15, 3.44, code_compute,
                   filename="common/compute.cpp  (shannon_entropy + compute_incremental)", lang="cpp",
                   font_size=7.9, start_line=None, accent_color=C_AMBER_LT, line_spacing=1.02)

    # Bottom Right: Comparison Table + Dependency note
    add_table(s, 7.88, 3.52, 4.90, 2.22,
              ["Cecha", "Naive", "Incremental"],
              [
                  ["Aktualizacja/okno", "W inkrementacji\n+ 256 zerowań", "**2·S aktualizacji**\n(dla S < W)"],
                  ["Złożoność bajtów", "O(K · W)", "**O(W + K · S)**"],
                  ["Koszt log₂(p)", "≤ 256 FP64 / okno", "≤ 256 FP64 / okno"],
                  ["Zależności okien", "Brak (niezależne)", "Łańcuch w paśmie\nwątku / ranku / bloku"],
              ],
              col_widths=[1.48, 1.66, 1.76], header_bg=C_NAVY, font_size=9.0, header_size=9.3)

    add_card(s, 7.88, 5.86, 4.90, 1.06, "Wyzwanie zrównoleglenia trybu incremental", accent_color=C_RED)
    add_textbox(s, 8.04, 6.26, 4.58, 0.60,
                "Ponieważ okno i zależy od histogramu okna i−1, zrównoleglenie wymaga podziału na spójne pasma: "
                "każdy wątek (OMP), proces (MPI) lub blok GPU (CUDA) inicjalizuje od zera tylko pierwsze okno swojego pasma!",
                size=9.2, color=C_TEXT, line_spacing=1.15)

    add_footer(s, 4, TOTAL)
    add_notes(s, (
        "SLAJD 4 (Rozszerzenie 2-osobowe — Adrian):\n"
        "• Zwrócić uwagę na elegancję linii 27 w common/compute.cpp: if (step >= window) return compute_naive(...); — dzięki temu wewnątrz głównej pętli nie ma żadnego warunku if, co poprawia przewidywanie skoków CPU.\n"
        "• Wyjaśnić próg opłacalności: dla S < W/2 liczba operacji 2S jest ściśle mniejsza od W. Przy S = W/4 teoretyczny zysk na zliczaniu to dokładnie 2,0x (ograniczony przez stały koszt 256 log2 dla małych W).\n"
        "• Kluczowe dla przedmiotu PRiR: tryb incremental wprowadza zależność danych między kolejnymi oknami — dlatego na kolejnych slajdach pokazujemy, jak przełamać tę zależność przez podział pasmowy (chunking)."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 5: OPENMP IMPLEMENTATION
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "04  •  Zrównoleglenie CPU (OpenMP)",
               "Implementacja OpenMP: pamięć współdzielona bez atomików",
               "Prywatne histogramy na stosie wątku (L1 cache), zero false sharingu i podział pasmowy block_range()",
               tag_color=C_BLUE, tag_bg=C_BLUE_BG)

    # Left: Real code from src/openmp.cpp + block_range from common/io.cpp
    code_omp = """// common/io.cpp — wspólny podział blokowy (OpenMP + MPI), różnica ≤ 1 okno
WindowRange block_range(std::size_t count, std::size_t parts, std::size_t index) {
    const auto base = count / parts, extra = count % parts;
    return {base * index + std::min(index, extra), base + (index < extra ? 1 : 0)};
}

// src/openmp.cpp — dwie ścieżki równoległe (Naive oraz Incremental)
if (options.mode == c05::Mode::Naive) {
    #pragma omp parallel
    {
        #pragma omp single
        used_threads = omp_get_num_threads();
        // window_entropy trzyma Histogram[256] na prywatnym stosie wątku (L1)!
        #pragma omp for schedule(static)
        for (std::ptrdiff_t i = 0; i < static_cast<std::ptrdiff_t>(layout.count); ++i) {
            const auto index = static_cast<std::size_t>(i);
            output[index] = c05::window_entropy(
                data.subspan(index * options.step, options.window));
        }
    }
} else {
    #pragma omp parallel
    {
        const auto threads = static_cast<std::size_t>(omp_get_num_threads());
        #pragma omp single
        used_threads = static_cast<int>(threads);
        const auto range = c05::block_range(layout.count, threads,
                                            static_cast<std::size_t>(omp_get_thread_num()));
        if (range.count)
            c05::compute_incremental(data.subspan(range.first * options.step),
                                     options.window, options.step,
                                     output.subspan(range.first, range.count));
    }
}"""
    add_code_block(s, 0.55, 1.40, 7.55, 5.50, code_omp,
                   filename="src/openmp.cpp (linie 27–53)  &  common/io.cpp (block_range)",
                   lang="cpp", font_size=8.6, start_line=None, accent_color=C_SKY)

    # Right: 4 crisp technical cards
    cards_omp = [
        ("1. Zero atomików i zero False Sharingu", C_BLUE,
         "Każdy wątek alokuje na własnym stosie prywatną tablicę c05::Histogram (256 × 4 B = 1 KB), "
         "która mieści się w całości w pamięci podręcznej L1 rdzenia (32 KB na Ryzen 5 2600). "
         "Tablica wejściowa data jest tylko do odczytu; zapis do output[i] następuje raz na okno do rozłącznych indeksów."),
        ("2. Harmonogram schedule(static) w trybie Naive", C_SLATE,
         "Każde pełne okno ma dokładnie taką samą długość W bajtów → identyczny koszt obliczeniowy. "
         "Przydział statyczny eliminuje narzut kolejki dynamicznej i zapewnia ciągły, przyjazny dla prefetchera CPU dostęp do pamięci."),
        ("3. Podział pasmowy w trybie Incremental", C_AMBER,
         "Funkcja block_range(K, P, tid) dzieli K okien na P ciągłych pasm o długościach różniących się "
         "o co najwyżej 1 okno. Koszt pełnego zliczenia W bajtów ponoszony jest tylko RAZ na wątek (pierwsze okno pasma), "
         "a pozostałe tysiące okien aktualizowane są w czasie O(S)."),
        ("4. Determinizm, synchronizacja i przypięcie wątków", C_GREEN,
         "Jedyna synchronizacja to niejawna bariera na końcu regionu #pragma omp parallel. "
         "W pomiarach wymuszono OMP_PROC_BIND=TRUE, OMP_PLACES=cores, OMP_DYNAMIC=FALSE. "
         "Wyniki są bitowo identyczne z wersją referencyjną (max_error = 0,0)."),
    ]
    for idx, (ctitle, ccol, cbody) in enumerate(cards_omp):
        cy = 1.40 + idx * 1.40
        add_card(s, 8.28, cy, 4.50, 1.28, ctitle, accent_color=ccol, title_size=11.0)
        add_textbox(s, 8.44, cy + 0.42, 4.18, 0.80, cbody, size=9.2, color=C_TEXT, line_spacing=1.15)

    add_footer(s, 5, TOTAL)
    add_notes(s, (
        "SLAJD 5 (OpenMP — Jan & Adrian):\n"
        "• Jan: omówić eliminację false sharingu (dlaczego prywatny Histogram 1 KB na stosie wątku zamiast globalnej tablicy [threads][256] bez paddingu) i wybór schedule(static).\n"
        "• Adrian: omówić block_range() w trybie incremental — dlaczego nie używamy zwykłego #pragma omp for w trybie przyrostowym (potrzebujemy jawnie znać początek i długość pasma, żeby wywołać compute_incremental).\n"
        "• Przypomnieć, że dzięki identycznej kolejności sumowania koszyków 0..255 wynik OpenMP ma błąd 0.0 względem Reference."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 6: MPI IMPLEMENTATION (HALO & COMMUNICATION)
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "05  •  Przetwarzanie rozproszone (MPI)",
               "Implementacja MPI: podział domenowy z zakładką Halo",
               "Własność 1:1 okien bez duplikacji wyników, nieblokująca dystrybucja MPI_Isend i scalanie MPI_Gatherv",
               tag_color=C_VIOLET, tag_bg=C_VIOLET_BG)

    # Top Diagram: MPI Halo Domain Decomposition
    add_card(s, 0.55, 1.38, 12.23, 1.64,
             "Schemat podziału pliku na P procesów z zakładką Halo (szerokość zakładki = W − S bajtów)",
             accent_color=C_VIOLET)

    # Rank 0 chunk + halo
    add_textbox(s, 0.75, 1.88, 0.95, 0.34, "Rank 0:", size=9.8, color=C_VIOLET, bold=True, font=FONT_SEMI, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 1.72, 1.88, 3.00, 0.36, C_VIOLET_BG, border_color=C_VIOLET, border_pt=1.3, rounded=True, radius=0.06)
    add_textbox(s, 1.72, 1.88, 3.00, 0.36, "Okna [0 .. k₀−1]", size=9.3, color=C_VIOLET, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 4.74, 1.88, 1.05, 0.36, C_AMBER_BG, border_color=C_AMBER, border_pt=1.3, rounded=True, radius=0.08)
    add_textbox(s, 4.74, 1.88, 1.05, 0.36, "Halo (W−S)", size=8.6, color=C_AMBER, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Rank 1 chunk + halo
    add_textbox(s, 0.75, 2.30, 0.95, 0.34, "Rank 1:", size=9.8, color=C_BLUE, bold=True, font=FONT_SEMI, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 4.74, 2.30, 3.00, 0.36, C_BLUE_BG, border_color=C_BLUE, border_pt=1.3, rounded=True, radius=0.06)
    add_textbox(s, 4.74, 2.30, 3.00, 0.36, "Okna [k₀ .. k₁−1]  (start: k₀·S)", size=9.3, color=C_BLUE, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_rect(s, 7.76, 2.30, 1.05, 0.36, C_AMBER_BG, border_color=C_AMBER, border_pt=1.3, rounded=True, radius=0.08)
    add_textbox(s, 7.76, 2.30, 1.05, 0.36, "Halo (W−S)", size=8.6, color=C_AMBER, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Rank P-1 chunk (no right halo needed)
    add_rect(s, 7.76, 1.88, 2.55, 0.36, C_GREEN_BG, border_color=C_GREEN, border_pt=1.3, rounded=True, radius=0.06)
    add_textbox(s, 7.76, 1.88, 2.55, 0.36, "Rank P−1:  [k_last .. K−1]", size=9.0, color=C_GREEN, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Summary formula on right
    add_rect(s, 10.48, 1.84, 2.12, 0.92, C_DARK_BG, rounded=True, radius=0.08)
    add_textbox(s, 10.56, 1.90, 1.96, 0.24, "Rozmiar bufora rangi r:", size=8.5, color=C_DARK_MUTED, align=PP_ALIGN.CENTER)
    add_textbox(s, 10.56, 2.15, 1.96, 0.30, "(count − 1)·S + W", size=11.0, color=C_AMBER_LT, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER)
    add_textbox(s, 10.56, 2.47, 1.96, 0.24, "0 gdy count = 0 (P > K)", size=8.3, color=C_GREEN_LT, align=PP_ALIGN.CENTER)

    add_textbox(s, 0.75, 2.72, 9.6, 0.24,
                "Każde okno ma dokładnie jednego właściciela ⇒ brak duplikacji obliczeń na granicach podobszarów i brak komunikacji w fazie rdzenia!",
                size=9.2, color=C_MUTED, bold=True)

    # Bottom Left: Real code from src/mpi.cpp
    code_mpi = """// 1. Rozgłoszenie rozmiaru i zakresy z zakładką (wszystkie rangi)
check(MPI_Bcast(&size, 1, MPI_UINT64_T, 0, MPI_COMM_WORLD));
for (std::size_t r = 0; r < processes; ++r) {
    const auto range = c05::block_range(layout.count, processes, r);
    counts[r] = mpi_count(range.count);  offsets[r] = mpi_count(range.first);
    byte_counts[r]  = mpi_count(range.count ? (range.count - 1) * S + W : 0);
    byte_offsets[r] = range.count ? mpi_count(range.first * S) : 0;
}
// 2. Dystrybucja nieblokująca (halo nakłada się na root → MPI_Scatterv zakazane)
if (rank == 0) {
    if (!local.empty()) std::copy_n(bytes.data(), local.size(), local.data());
    for (int dst = 1; dst < ranks; ++dst) if (byte_counts[dst]) {
        requests.emplace_back();
        check(MPI_Isend(bytes.data() + byte_offsets[dst], byte_counts[dst],
                        MPI_UINT8_T, dst, 10, MPI_COMM_WORLD, &requests.back()));
    }
    if (!requests.empty()) check(MPI_Waitall(requests.size(), requests.data(), MPI_STATUSES_IGNORE));
} else if (byte_counts[r]) {
    check(MPI_Recv(local.data(), byte_counts[r], MPI_UINT8_T, 0, 10, MPI_COMM_WORLD, MPI_STATUSES_IGNORE));
}
// 3. Synchronizowany rdzeń, redukcja min/max czasu i scalenie MPI_Gatherv
check(MPI_Barrier(MPI_COMM_WORLD));
const double t0 = MPI_Wtime();
c05::compute_mode(options.mode, local, W, S, local_entropy);
const double dt = MPI_Wtime() - t0;
std::array<double, 2> mm{-dt, dt}, red{};
check(MPI_Reduce(mm.data(), red.data(), 2, MPI_DOUBLE, MPI_MAX, 0, MPI_COMM_WORLD));
check(MPI_Gatherv(local_entropy.data(), counts[r], MPI_DOUBLE, entropy.data(),
                  counts.data(), offsets.data(), MPI_DOUBLE, 0, MPI_COMM_WORLD));"""
    add_code_block(s, 0.55, 3.14, 7.95, 3.80, code_mpi,
                   filename="src/mpi.cpp  (linie 49–94)",
                   lang="cpp", font_size=7.6, start_line=None,
                   accent_color=RGBColor(196, 181, 253), line_spacing=1.01)

    # Bottom Right: Technical details cards
    add_card(s, 8.68, 3.14, 4.10, 3.80, "Kluczowe detale techniczne MPI", accent_color=C_VIOLET)
    add_bullets(s, 8.84, 3.60, 3.78, 3.24, [
        ("Dlaczego MPI_Isend zamiast MPI_Scatterv? ",
         "Standard MPI zabrania wielokrotnego czytania tej samej pamięci źródłowej na rootcie w MPI_Scatterv. Nieblokujące MPI_Isend + MPI_Waitall bezpiecznie wysyłają nakładające się zakładki bez kopiowania całego pliku."),
        ("Zachowanie globalnych offsetów: ",
         "MPI_Gatherv umieszcza local_entropy pod globalnym przesunięciem offsets[r] = first_r, dzięki czemu offset_i = i·S jest identyczny jak w wersji sekwencyjnej."),
        ("Sprytny pomiar min i max czasu: ",
         "Pojedyncze MPI_Reduce(..., MPI_MAX) na parze {-local_core, +local_core} zwraca jednocześnie najwolniejszą i najszybszą rangę (diagnostyka balansu)."),
        ("Przypadek brzegowy P > K: ",
         "Gdy procesów jest więcej niż okien, nadmiarowe rangi mają count = 0 i bezpiecznie pomijają odbiór i obliczenia (przetestowane dla p = 7 w verify_mpi.py)."),
    ], size=9.1, space_after=5.0, bullet_color=C_VIOLET)

    add_footer(s, 6, TOTAL)
    add_notes(s, (
        "SLAJD 6 (MPI — Adrian):\n"
        "• To bardzo często zadawane pytanie na obronie PRiR: 'Dlaczego nie użyliście MPI_Scatterv?' -> Odpowiedź: w MPI_Scatterv nakładające się buforowo zakresy wysyłane z roota naruszają regułę braku aliasingu bufora wejściowego (MPI standard), natomiast punkt-punkt MPI_Isend + MPI_Waitall rozwiązuje to czysto bez duplikacji 150 MB RAM.\n"
        "• Zwrócić uwagę na trik z parą {-local_core, local_core} w jednym MPI_Reduce(MPI_MAX) — pozwala raportować zarówno T_core (max) jak i T_core_min_rank_s w jednej komunikacji kolektywnej.\n"
        "• Podkreślić jawność metodologiczną: nasze pomiary MPI wykonano na 1 węźle (transport shared-memory OpenMPI, /dev/shm = 2 GB)."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 7: CUDA PART 1 — SHARED MEMORY, NAIVE KERNEL & TREE REDUCTION
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "06  •  Akceleracja GPU (NVIDIA CUDA — 1/2)",
               "CUDA: pamięć współdzielona per blok i redukcja drzewiasta",
               "Brak globalnych operacji atomowych, koalescencyjny odczyt z VRAM i 8-krokowa redukcja FP64 w shared memory",
               tag_color=C_GREEN, tag_bg=C_GREEN_BG)

    # Left: Real code from src/cuda_entropy.cu (block_entropy + histogram_entropy_kernel)
    code_cuda1 = """__device__ double block_entropy(const unsigned int* s_hist,
                                double* s_terms, std::size_t window) {
    const unsigned int thread = threadIdx.x;
    for (unsigned int bin = thread; bin < 256; bin += blockDim.x) {
        const unsigned int c = s_hist[bin];
        if (c == 0) s_terms[bin] = 0.0;       // pomijamy zerowe koszyki
        else {
            const double p = static_cast<double>(c) / static_cast<double>(window);
            s_terms[bin] = -p * log2(p);
        }
    }
    __syncthreads();
    // Redukcja drzewiasta (pętla wewnętrzna obsługuje też bloki 64 i 128!)
    for (unsigned int stride = 128; stride > 0; stride >>= 1) {
        for (unsigned int bin = thread; bin < stride; bin += blockDim.x)
            s_terms[bin] += s_terms[bin + stride];
        __syncthreads();
    }
    const double result = s_terms[0];
    __syncthreads();
    return result;
}

__global__ void histogram_entropy_kernel(const std::uint8_t* data,
    std::size_t window, std::size_t step, std::size_t count, double* d_entropy) {
    __shared__ unsigned int s_hist[256];      // 1 KB per blok w SM
    __shared__ double       s_terms[256];     // 2 KB per blok w SM
    const unsigned int thread = threadIdx.x;
    for (std::size_t index = blockIdx.x; index < count;) {
        for (unsigned int b = thread; b < 256; b += blockDim.x) s_hist[b] = 0;
        __syncthreads();
        const std::size_t offset = index * step;
        for (std::size_t j = thread; j < window; j += blockDim.x)
            atomicAdd(&s_hist[data[offset + j]], 1U); // koalescencja + shared atomic
        __syncthreads();
        const double h = block_entropy(s_hist, s_terms, window);
        if (thread == 0) d_entropy[index] = h;
        if (count - index <= static_cast<std::size_t>(gridDim.x)) break;
        index += gridDim.x;
    }
}"""
    add_code_block(s, 0.55, 1.40, 7.65, 5.52, code_cuda1,
                   filename="src/cuda_entropy.cu  (linie 82–124)", lang="cpp",
                   font_size=8.1, start_line=82, accent_color=C_GREEN_LT)

    # Right Top: 4-step visual pipeline inside a CUDA Block
    add_card(s, 8.38, 1.40, 4.40, 3.22, "4 fazy pracy jednego bloku CUDA (64..1024 wątki)", accent_color=C_GREEN)
    steps_cuda = [
        ("Faza 1", "Równoległe zerowanie s_hist[0..255] → __syncthreads()", C_SLATE),
        ("Faza 2", "Koalescencyjny odczyt bajtów z VRAM (krok blockDim.x) + atomicAdd(&s_hist[b], 1U) → __syncthreads()", C_BLUE),
        ("Faza 3", "Wyznaczenie 256 składników −p·log₂(p) w FP64 double do s_terms[0..255] → __syncthreads()", C_VIOLET),
        ("Faza 4", "8-krokowa redukcja drzewiasta (128→64→…→1); wątek 0 zapisuje d_entropy[index] = s_terms[0]", C_GREEN),
    ]
    for idx, (st, desc, col) in enumerate(steps_cuda):
        sy = 1.88 + idx * 0.66
        add_badge(s, 8.54, sy + 0.06, 0.80, 0.32, st, col, C_WHITE, size=8.8)
        add_textbox(s, 9.44, sy, 3.20, 0.58, desc, size=9.1, color=C_TEXT, line_spacing=1.12)

    # Right Bottom: Boundary kernel + Precision note
    add_card(s, 8.38, 4.76, 4.40, 2.16, "Detekcja granic na GPU i precyzja FP64", accent_color=C_NAVY)
    add_bullets(s, 8.54, 5.20, 4.08, 1.66, [
        ("Drugi kernel (boundary_flags_kernel): ",
         "równoległa pętla grid-stride (256 wątków/blok) ustawia d_flags[i] = (i != 0 && fabs(ΔH) >= T)."),
        ("Weryfikacja na hoście: ",
         "po transferze D2H host sprawdza brak NaN/Inf oraz zgodność 1:1 flag GPU z CPU (std::mismatch)."),
        ("Dlaczego błąd GPU wynosi 1,69×10⁻¹⁴? ",
         "Wyłącznie przez inną kolejność sumowania FP64 (drzewo binarne zamiast 0..255); brak --use_fast_math."),
    ], size=9.0, space_after=4.5, bullet_color=C_GREEN)

    add_footer(s, 7, TOTAL)
    add_notes(s, (
        "SLAJD 7 (CUDA 1/2 — Jan):\n"
        "• Podkreślić 3 kluczowe optymalizacje GPU: (1) pamięć współdzielona __shared__ na blok (3 KB łącznie -> wysoka okupacja SM), (2) koalescencyjny dostęp do pamięci globalnej (sąsiednie wątki czytają sąsiednie bajty data[offset + threadIdx.x]), (3) pętla grid-stride z zabezpieczeniem przed przepełnieniem size_t.\n"
        "• Bardzo ważny szczegół w block_entropy: wewnętrzna pętla for (bin = thread; bin < stride; bin += blockDim.x) oraz końcowe __syncthreads() przed return — dzięki temu funkcja działa bezbłędnie także dla bloków 64-wątkowych i w trybie incremental (gdzie zaraz po powrocie modyfikujemy s_hist).\n"
        "• Wytłumaczyć 1.69e-14: dodawanie IEEE-754 nie jest ściśle łączne; drzewo sumuje (0+128), (1+129) itd., a CPU 0+1+2+..."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 8: CUDA PART 2 — INCREMENTAL KERNEL, OCCUPANCY & TIMING
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "07  •  Akceleracja GPU (NVIDIA CUDA — 2/2)",
               "CUDA Incremental: pasma okien, Occupancy API i pomiar faz",
               "Przesuwanie histogramu w shared memory (atomicSub + atomicAdd) z dynamicznym równoważeniem SM",
               tag_color=C_GREEN, tag_bg=C_GREEN_BG)

    # Left: Real code from src/cuda_entropy.cu (incremental_entropy_kernel + host occupancy)
    code_cuda2 = """// 1. Kernel przyrostowy (uruchamiany wyłącznie gdy Mode::Incremental && S < W)
__global__ void incremental_entropy_kernel(const std::uint8_t* data, std::size_t window,
    std::size_t step, std::size_t count, std::size_t per_block, double* d_entropy) {
    __shared__ unsigned int s_hist[256];  __shared__ double s_terms[256];
    const unsigned int thread = threadIdx.x;
    const std::size_t chunks = (count + per_block - 1) / per_block;
    for (std::size_t chunk = blockIdx.x; chunk < chunks;) {
        const std::size_t first = chunk * per_block;
        const std::size_t last  = count - first < per_block ? count : first + per_block;
        // Pierwsze okno pasma: pełna budowa histogramu (W wywołań atomicAdd)
        for (unsigned int b = thread; b < 256; b += blockDim.x) s_hist[b] = 0;
        __syncthreads();
        for (std::size_t j = thread; j < window; j += blockDim.x)
            atomicAdd(&s_hist[data[first * step + j]], 1U);
        __syncthreads();
        double h = block_entropy(s_hist, s_terms, window);
        if (thread == 0) d_entropy[first] = h;
        // Kolejne okna w paśmie: tylko -S starych i +S nowych bajtów!
        for (std::size_t i = first + 1; i < last; ++i) {
            const std::size_t old_start = (i - 1) * step;
            for (std::size_t j = thread; j < step; j += blockDim.x) {
                atomicSub(&s_hist[data[old_start + j]], 1U);
                atomicAdd(&s_hist[data[old_start + window + j]], 1U);
            }
            __syncthreads();
            h = block_entropy(s_hist, s_terms, window);
            if (thread == 0) d_entropy[i] = h;
        }
        if (chunks - chunk <= static_cast<std::size_t>(gridDim.x)) break;
        chunk += gridDim.x;
    }
}
// 2. Host: dynamiczny dobór liczby okien na blok (Occupancy API, limit 64)
const bool sliding = options.mode == c05::Mode::Incremental && options.step < options.window;
if (sliding) {
    int per_sm = 0;
    CHECK_CUDA(cudaOccupancyMaxActiveBlocksPerMultiprocessor(&per_sm, incremental_entropy_kernel, threads, 0));
    const std::size_t resident = std::max(1, per_sm) * device.multiProcessorCount;
    windows_per_block = std::min<std::size_t>(64, std::max<std::size_t>(1, (layout.count + resident - 1) / resident));
}"""
    add_code_block(s, 0.55, 1.40, 7.75, 5.52, code_cuda2,
                   filename="src/cuda_entropy.cu  (linie 131–162 oraz 193–227)", lang="cpp",
                   font_size=7.4, start_line=None, accent_color=C_GREEN_LT, line_spacing=1.01)

    # Right: 3 deep technical insight cards
    add_card(s, 8.48, 1.40, 4.30, 1.80, "Poprawność przeplotu atomicSub / atomicAdd", accent_color=C_AMBER)
    add_textbox(s, 8.64, 1.84, 3.98, 1.28,
                "Wątki bloku mogą wykonywać atomicSub(-1) i atomicAdd(+1) w dowolnej kolejności — "
                "czasami chwilowo zawijając licznik uint32_t poniżej 0 (np. 0 − 1 = 2³² − 1). "
                "Ponieważ dodawanie modulo 2³² jest przemienne i łączne, a każdy odejmowany bajt istniał w poprzednim oknie, "
                "po barierze __syncthreads() histogram jest zawsze dokładnie poprawny (zweryfikowane w check_cuda_model.py).",
                size=9.0, color=C_TEXT, line_spacing=1.14)

    add_card(s, 8.48, 3.34, 4.30, 1.72, "Balans: amortyzacja pasma vs obłożenie SM", accent_color=C_GREEN)
    add_textbox(s, 8.64, 3.78, 3.98, 1.20,
                "• Za długie pasmo ⇒ zbyt mało bloków dla 22 multiprocesorów SM (niskie obłożenie GPU).\n"
                "• Za krótkie pasmo ⇒ częste budowanie pierwszego okna od zera.\n"
                "• Limit 64 okien/pasmo + Occupancy API nasyca wszystkie SM (dla S ≥ W od razu uruchamiamy histogram_entropy_kernel!).",
                size=9.1, color=C_TEXT, line_spacing=1.18)

    add_card(s, 8.48, 5.20, 4.30, 1.72, "RAII i dwuzegarowy pomiar faz GPU", accent_color=C_BLUE)
    add_textbox(s, 8.64, 5.64, 3.98, 1.20,
                "• Zasoby GPU opakowane w klasy RAII: DeviceBuffer<T> i Event (bez wycieków pamięci VRAM).\n"
                "• Szablon measure() wykonuje cudaDeviceSynchronize() przed i po każdej fazie, mierząc osobno "
                "H2D, kernel entropii, kernel granic oraz D2H zegarem CPU (wall) i sprzętowym cudaEventElapsedTime().",
                size=9.1, color=C_TEXT, line_spacing=1.16)

    add_footer(s, 8, TOTAL)
    add_notes(s, (
        "SLAJD 8 (CUDA 2/2 — Adrian):\n"
        "• Świetny techniczny punkt na pytania komisji: co się stanie, jeśli wątek A zrobi atomicSub na bajcie X przed wątkiem B robiącym atomicAdd na bajcie X? Dzięki arytmetyce modularnej unsigned int mod 2^32 (grupa abelowa) kolejność nie ma znaczenia! Ten mechanizm ma własny test w tests/check_cuda_model.py.\n"
        "• Podkreślić refaktoring flagi sliding = (Mode::Incremental && step < window): gdy okna są rozłączne, w ogóle nie uruchamiamy incremental_entropy_kernel, co usunęło martwą gałąź else i dodatkowy __syncthreads() wewnątrz kernela.\n"
        "• Przypomnieć, że T_core dla CUDA to wyłącznie czas kernela entropii z cudaEventElapsedTime(), aby było to uczciwie porównywalne z CPU T_core."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 9: HARDWARE, DATASETS & BENCHMARK METHODOLOGY
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "08  •  Środowisko i metodologia",
               "Platforma sprzętowa, zbiory danych i metodologia pomiarów",
               "Wszystkie pomiary wykonane na jednej maszynie fizycznej ze ścisłym rozdzieleniem T_core i T_total")

    # Left: Hardware & Software Table
    add_card(s, 0.55, 1.40, 5.35, 3.62, "Konfiguracja stanowiska testowego", accent_color=C_NAVY)
    add_table(s, 0.70, 1.88, 5.05, 2.98,
              ["Komponent", "Specyfikacja"],
              [
                  ["Procesor (CPU)", "AMD Ryzen 5 2600 (Zen+, 1 węzeł, 6 rdzeni fiz. / 12 SMT)"],
                  ["Pamięć RAM", "16 GB DDR4 (11,68 GB w WSL2, /dev/shm = 2 GB dla MPI)"],
                  ["Karta GPU", "NVIDIA GeForce GTX 1660 (Turing sm_75, 22 SM, 1408 CUDA, 6 GB)"],
                  ["System / OS", "Ubuntu 24.04 LTS (WSL2, Linux 5.15.167, glibc 2.39)"],
                  ["Kompilatory", "GCC 13.3.0 (-O3 -march=native), nvcc 12.6.85 (-O3 -arch=sm_75)"],
                  ["MPI / Python", "OpenMPI 4.1.6 (local shared-memory), Python 3.12 + SciPy stack"],
              ],
              col_widths=[1.45, 3.60], header_bg=C_NAVY, font_size=8.6, header_size=9.0)

    # Right: Datasets Table
    add_card(s, 6.08, 1.40, 6.70, 3.62, "Deterministyczne zbiory danych (seed = 2026, sumy SHA-256)", accent_color=C_BLUE)
    add_table(s, 6.23, 1.88, 6.40, 2.98,
              ["Zbiór danych", "Rozmiar", "Symbole k", "H_teor [bit/B]", "Przeznaczenie"],
              [
                  ["constant_bytes.bin", "100 MB", "1 (0x00)", "0,000", "100% kolizji atomic GPU"],
                  ["uniform_random.bin", "100 MB", "256", "8,000", "Minimalna kolizja atomic"],
                  ["mixed_segmented.bin", "150 MB", "23 / 252 / 1", "4,52 / 7,95 / 0", "Główny (2 granice: 50, 100 MB)"],
                  ["multi_segmented.bin", "150 MB", "1/4/23/256/16", "0/2/4,5/8/4", "5 segmentów (4 granice)"],
                  ["ksym_{1..256}.bin", "9 × 50 MB", "1, 2, 4 … 256", "log₂(k) = 0..8", "Krzywa kolizji atomowych GPU"],
              ],
              col_widths=[1.60, 0.85, 1.00, 1.05, 1.90], right_cols=(1,),
              header_bg=C_BLUE, font_size=8.5, header_size=8.8)

    # Bottom: 3 Methodology Pillars
    pillars = [
        ("Rozdzielenie T_core vs T_total", C_GREEN,
         "• T_core: wyłącznie histogram + entropia (CPU steady_clock; MPI bariera + MPI_MAX; GPU cudaEventElapsedTime).\n"
         "• T_total: od odczytu pliku z dysku, przez setup/MPI/PCIe i detekcję, do flush+close CSV."),
        ("Rygor statystyczny (134 konf.)", C_AMBER,
         "• Każda konfiguracja: 1 odrzucany przebieg rozgrzewkowy (warm-up) + 5 niezależnych powtórzeń właściwych.\n"
         "• Losowa kolejność konfiguracji (seed=2026) eliminuje dryft termiczny; raportujemy medianę i rozrzut min–max."),
        ("Dwie metryki przepustowości", C_VIOLET,
         "• Wejściowa (N / T_core): porównuje rzeczywistą szybkość skanowania pliku między naive a incremental.\n"
         "• Logiczna (K·W / T_core): efektywne tempo przetwarzania bajtów okien (nie utożsamiać z pasmem DRAM!)."),
    ]
    for idx, (pt, pc, pb) in enumerate(pillars):
        px = 0.55 + idx * 4.14
        add_card(s, px, 5.18, 3.96, 1.74, pt, accent_color=pc, title_size=11.0)
        add_textbox(s, px + 0.14, 5.60, 3.68, 1.24, pb, size=9.2, color=C_TEXT, line_spacing=1.18)

    add_footer(s, 9, TOTAL)
    add_notes(s, (
        "SLAJD 9 (Środowisko i metodologia — Jan):\n"
        "• Warto zwrócić uwagę na rozróżnienie dwóch przepustowości: wejściowej N/T_core oraz logicznej K*W/T_core — przy nakładaniu S=W/4 przepustowość logiczna jest 4x większa od wejściowej, co jest jawnie udokumentowane w metadata.json i na osiach wykresów.\n"
        "• Każdy zbiór jest w 100% odtwarzalny (generate_datasets.py --seed 2026) i ma zapisany SHA-256 oraz dokładne ground_truth_boundaries.json."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 10: VALIDATION & ENTROPY ESTIMATION ERROR (MILLER-MADOW)
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "09  •  Walidacja i dokładność",
               "Weryfikacja poprawności (CTest) i błąd estymacji entropii",
               "100% zaliczonych testów automatycznych (CPU błąd = 0,0; GPU ≤ 1,69×10⁻¹⁴) i weryfikacja obciążenia Millera–Madowa",
               tag_color=C_GREEN, tag_bg=C_GREEN_BG)

    # Left Top: CTest & Cross-backend table
    add_card(s, 0.55, 1.40, 5.85, 3.22, "Wyniki walidacji CTest (4/4 PASS) i błędu na 150 MB", accent_color=C_GREEN)
    add_table(s, 0.70, 1.88, 5.55, 2.60,
              ["Test / Backend", "Zakres weryfikacji", "Max |ΔH|", "Flagi"],
              [
                  ["host_model_correctness", "500 modeli redukcji + uint32 mod", "< 10⁻¹²", "**PASS**"],
                  ["ref_and_openmp_correctness", "144 przypadki + CLI + --events", "**0,0**", "**PASS**"],
                  ["mpi_correctness", "128 konf., p ∈ {1,2,4,7} (p > K)", "**0,0**", "**PASS**"],
                  ["cuda_correctness (GTX 1660)", "112 konf., bloki 64..1024", "3,55×10⁻¹⁵", "**PASS**"],
                  ["Ref/OMP/MPI (150 MB, W=64K)", "Naive + Incremental (9152 okna)", "**0,0**", "100%"],
                  ["CUDA / CUDA-Inc (150 MB)", "Naive + Incremental (9152 okna)", "1,69×10⁻¹⁴", "100%"],
              ],
              col_widths=[1.92, 2.08, 0.90, 0.65], right_cols=(2, 3),
              header_bg=C_GREEN, font_size=8.4, header_size=8.7, highlight_rows=(3, 5))

    # Left Bottom: Miller-Madow statistical insight
    add_card(s, 0.55, 4.76, 5.85, 2.16, "Statystyczne obciążenie estymatora (poprawka Millera–Madowa)", accent_color=C_BLUE)
    add_rect(s, 0.72, 5.20, 5.50, 0.40, C_SLATE_BG, border_color=C_BORDER, rounded=True, radius=0.1)
    add_textbox(s, 0.78, 5.25, 5.38, 0.30,
                "Bias_MM(W, k) = E[H_emp] − H_teor ≈ − (k − 1) / (2 · W · ln 2)",
                size=10.5, color=C_BLUE, bold=True, font=FONT_CODE, align=PP_ALIGN.CENTER)
    add_bullets(s, 0.72, 5.68, 5.50, 1.18, [
        ("Zgodność z teorią < 1%: ", "np. dla uniform_256 (k=256) i W=1024: bias teor. = −0,1796, empiryczny = −0,1913; dla W=65536: −0,00281 vs −0,00281."),
        ("Praktyczny dobór W: ", "błąd maleje ∝ W⁻¹. Przy W=1024 losowe fluktuacje sięgają 0,27 bit/B (ryzyko FP dla T=0,1); od W ≥ 4096 spadają < 0,045 bit/B."),
    ], size=9.1, space_after=4.0, bullet_color=C_BLUE)

    # Right: High-res plot entropy_error_analysis.png
    add_picture_fit(s, plots / "entropy_error_analysis.png", 6.56, 1.40, 6.22, 5.52,
                    card=True, caption="Empiryczny błąd |H_emp − H_teor| vs krzywa Millera–Madowa (entropy_error_analysis.png)")

    add_footer(s, 10, TOTAL)
    add_notes(s, (
        "SLAJD 10 (Walidacja i błąd entropii — Jan):\n"
        "• Rozdzielamy dwie zupełnie różne klasy błędów: (1) błąd numeryczny między implementacjami równoległymi a referencją (0.0 na CPU, 1.69e-14 na GPU — czyli o 8 rzędów wielkości poniżej tolerancji 1e-6), oraz (2) błąd statystyczny estymacji entropii z okna skończonego W względem entropii teoretycznej generatora.\n"
        "• Wykres log-log po prawej pięknie potwierdza nachylenie -1 zgodne z poprawką Millera-Madowa dla wszystkich badanych alfabetów k = 4, 16, 23, 252, 256."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 11: STRONG SCALING — OPENMP & LOCAL MPI
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "10  •  Skalowanie CPU i MPI",
               "Skalowanie silne OpenMP oraz lokalnego MPI (Ryzen 5 2600)",
               "mixed_segmented.bin (150 MB), W = 65536, S = 16384 (K = 9152 okna)  •  1 węzeł: 6 rdzeni fizycznych / 12 wątków SMT",
               tag_color=C_BLUE, tag_bg=C_BLUE_BG)

    # Top: Wide plot speedup_efficiency_omp_mpi.png (ratio 2.465 fits full width!)
    add_picture_fit(s, plots / "speedup_efficiency_omp_mpi.png", 0.55, 1.38, 12.23, 3.35,
                    card=True, caption="Mediana z 5 powtórzeń z konserwatywną obwiednią rozrzutu min–max; pionowa linia oznacza 6 rdzeni fizycznych")

    # Bottom Left: Compact table of key scaling rows
    add_table(s, 0.55, 4.86, 7.35, 2.08,
              ["Backend", "p", "T_core [ms]", "T_total [ms]", "N/T_core", "S(p)", "E(p)"],
              [
                  ["OMP (naive)", "1 → 6 → 8 → 12", "1335 → **241** → 200 → 222", "303 (p=6)", "0,62 GB/s", "**5,54×** (p=6) / 6,67× (p=8)", "**92,4%** (p=6)"],
                  ["OMP-Incremental", "1 → 6 → 8 → 12", "655 → **126** → **96,9** → 97,4", "**166,6** (p=12)", "**1,55 GB/s**", "5,20× (p=6) / **6,76×** (p=8)", "**86,7%** (p=6)"],
                  ["MPI (naive, lok.)", "1 → 6 → 8", "1325 → **243** → 183", "359 (p=6)", "0,62 GB/s", "**5,46×** (p=6) / **7,23×** (p=8)", "**91,0%** (p=6)"],
                  ["MPI-Inc (lok.)", "1 → 6 → 8", "655 → **131** → **97,7**", "272 (p=6)", "1,53 GB/s", "4,99× (p=6) / 6,70× (p=8)", "**83,1%** (p=6)"],
              ],
              col_widths=[1.22, 1.15, 1.62, 0.92, 0.72, 1.00, 0.72],
              right_cols=(2, 3, 4, 5, 6), header_bg=C_NAVY, font_size=7.9, header_size=8.4,
              highlight_rows=(1,))

    # Bottom Right: Key analytical takeaways
    add_card(s, 8.06, 4.86, 4.72, 2.08, "Kluczowe wnioski ze skalowania CPU / MPI", accent_color=C_VIOLET)
    add_bullets(s, 8.22, 5.28, 4.40, 1.60, [
        ("Niemal liniowe skalowanie do p = 6: ", "efektywność > 91% (OMP 92,4%, MPI 91,0%) potwierdza brak kolizji w L1 cache."),
        ("Efekt SMT (p = 8 i p = 12): ", "8 wątków podnosi przyspieszenie do 6,76–7,23×; przy 12 wątkach następuje nasycenie jednostek wykonawczych (E = 50–56%)."),
        ("Koszt komunikacji MPI (1 węzeł): ", "T_total w MPI jest wyższe o 45–70 ms od OpenMP przez dystrybucję 150 MB z halo (~26 ms) i start orted."),
    ], size=8.9, space_after=4.0, bullet_color=C_VIOLET)

    add_footer(s, 11, TOTAL)
    add_notes(s, (
        "SLAJD 11 (Skalowanie OpenMP i MPI — Adrian):\n"
        "• Zwrócić uwagę na pionową przerywaną linię przy p=6 — do liczby fizycznych rdzeni obie krzywe niemal przylegają do idealnej przekątnej S(p)=p.\n"
        "• Dlaczego OMP-Incremental przy p=12 ma niższe T_total (166,6 ms) niż przy p=6 (200,2 ms), mimo że efektywność E(12) spada do 56%? Bo absolutny czas rdzenia spadł ze 126 ms do 97 ms.\n"
        "• Zaznaczyć uczciwie, że MPI mierzono na pojedynczym węźle (shared-memory); w klastrze wielowęzłowym narzut T_distribute rósłby wraz z opóźnieniem sieci."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 12: ALL 8 VARIANTS & NAIVE VS INCREMENTAL COMPARISON
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "11  •  Porównanie technologii i trybów",
               "Wszystkie 4 technologie × Naive vs Incremental (zakres 2-osobowy)",
               "Przepustowość rdzenia i przyspieszenie wariantu przyrostowego na zbiorze mixed_segmented.bin (150 MB)",
               tag_color=C_AMBER, tag_bg=C_AMBER_BG)

    # Top Left & Top Right plots side by side
    add_picture_fit(s, plots / "throughput_comparison.png", 0.55, 1.38, 7.45, 3.02,
                    card=True, caption="Przepustowość N / T_core (log-log): S = W (rozłączne) vs S = W/4 (75% zakładki)")
    add_picture_fit(s, plots / "naive_vs_incremental_comparison.png", 8.15, 1.38, 4.63, 3.02,
                    card=True, caption="Przyspieszenie T_core(naive) / T_core(inc) dla S = W/4")

    # Bottom Left: Summary Table for W=65536, S=16384
    add_table(s, 0.55, 4.52, 8.20, 2.42,
              ["Implementacja (W=64K, S=16K)", "T_core [ms]", "T_total [ms]", "N/T_core", "KW/T_core", "vs Ref-Naive", "Inc / Naive"],
              [
                  ["Ref-Naive (1 wątek)", "1300,18", "1359,41", "0,12 GB/s", "0,46 GB/s", "1,00×", "—"],
                  ["Ref-Incremental (1 wątek)", "666,94", "784,00", "0,22 GB/s", "0,90 GB/s", "1,95×", "**1,95×**"],
                  ["OMP (12 wątków)", "222,45", "299,91", "0,67 GB/s", "2,70 GB/s", "5,84×", "—"],
                  ["OMP-Incremental (12 wątków)", "97,39", "**166,56**", "1,54 GB/s", "6,16 GB/s", "13,35×", "**2,28×**"],
                  ["MPI (6 procesów)", "242,64", "358,62", "0,62 GB/s", "2,47 GB/s", "5,36×", "—"],
                  ["MPI-Incremental (6 procesów)", "131,39", "272,35", "1,14 GB/s", "4,56 GB/s", "9,90×", "**1,85×**"],
                  ["CUDA (blok 256 / opt. 512)", "20,52 / 18,83", "285,43", "7,31 GB/s", "29,23 GB/s", "63,4× / 69,1×", "—"],
                  ["CUDA-Inc (blok 256 / opt. 64)", "**11,49 / 10,02**", "290,30", "**13,05 GB/s**", "**52,2 / 59,9**", "**113,1× / 129,8×**", "**1,78× / 2,44×**"],
              ],
              col_widths=[2.22, 1.15, 0.88, 0.90, 1.00, 1.12, 0.93],
              right_cols=(1, 2, 3, 4, 5, 6), header_bg=C_NAVY, font_size=7.8, header_size=8.1,
              highlight_rows=(3, 7))

    # Bottom Right: Analytical insights
    add_card(s, 8.90, 4.52, 3.88, 2.42, "Zysk wariantu przyrostowego", accent_color=C_AMBER)
    add_bullets(s, 9.05, 4.94, 3.58, 1.94, [
        ("Blisko granicy 2,0× dla S = W/4: ", "dla W ≥ 16 KB redukcja czasu rdzenia sięga 1,78–2,28× we wszystkich 4 backendach (w CUDA blk=64 aż 2,44×)."),
        ("Dlaczego dla W = 1024 zysk spada do 1,25–1,55×? ", "Stały koszt 256 logarytmów FP64 na okno waży wtedy niemal tyle samo co 2S = 512 operacji całkowitych."),
        ("Okna rozłączne (S = W): ", "oba warianty wykonują tę samą pracę i osiągają identyczną wydajność."),
    ], size=8.8, space_after=4.5, bullet_color=C_AMBER)

    add_footer(s, 12, TOTAL)
    add_notes(s, (
        "SLAJD 12 (Porównanie wszystkich 8 wariantów — Adrian):\n"
        "• Ten slajd podsumowuje serce porównania 2-osobowego: dla S=W/4 wariant incremental wygrywa we WSZYSTKICH 4 technologiach, zbliżając się do teoretycznego 2,0x (a w OpenMP przy W=64K nawet 2,28x dzięki lepszej lokalności w cache!).\n"
        "• Zwrócić uwagę na kluczowy kontrast w tabeli: w T_core bezapelacyjnie dominuje CUDA-Incremental (10–11,5 ms, do 129,8x szybciej od Ref-Naive), natomiast w T_total dla pojedynczego pliku 150 MB wygrywa 12-wątkowy OMP-Incremental (166,6 ms vs 290,3 ms) — wyjaśnienie na następnym slajdzie!"
    ))

    # -------------------------------------------------------------------------
    # SLIDE 13: GPU ANALYSIS 1/2 — TRANSFERS, BLOCK SIZE & DATA SIZE SCALING
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "12  •  Analiza wydajności GPU (1/2)",
               "Koszt transferów PCIe, rozmiar bloku i skalowanie rozmiaru pliku",
               "Kiedy narzut PCIe i inicjalizacji sterownika WDDM przewyższa czas wykonania kernela na GTX 1660?",
               tag_color=C_GREEN, tag_bg=C_GREEN_BG)

    # Top Left: Transfer breakdown
    add_picture_fit(s, plots / "cuda_transfer_breakdown.png", 0.55, 1.38, 5.20, 2.74,
                    card=True, caption="Dekompozycja faz GPU (H2D = 22–23 ms ≈ 53% czasu dla W=64K)")

    # Top Right: Block size impact
    add_picture_fit(s, plots / "cuda_block_size_impact.png", 5.90, 1.38, 6.88, 2.74,
                    card=True, caption="Wpływ liczby wątków w bloku CUDA (64..1024) na czas kernela i przepustowość")

    # Bottom Left: Data size scaling
    add_picture_fit(s, plots / "gpu_datasize_scaling.png", 0.55, 4.24, 7.15, 2.70,
                    card=True, caption="Skalowanie względem rozmiaru pliku N = 10..180 MB (W=65536, S=16384)")

    # Bottom Right: 3 Crisp Engineering Findings
    add_card(s, 7.86, 4.24, 4.92, 2.70, "Wnioski architektoniczne dla GPU", accent_color=C_GREEN)
    add_bullets(s, 8.02, 4.66, 4.60, 2.22, [
        ("Transfer H2D dominuje przy dużych W: ",
         "Przesłanie 150 MB przez PCIe 3.0 (bufor pageable) trwa stałe ~22–23 ms (~7,7 GB/s) — dłużej niż sam kernel CUDA-Inc (11,5 ms) i aż 52,9% czasu faz GPU!"),
        ("Optymalny rozmiar bloku zależy od trybu: ",
         "CUDA-Naive preferuje 256–512 wątków (18,83 ms); CUDA-Inc i dane jednorodne wygrywają dla 64–128 wątków (10,02 ms = 59,9 GB/s log.), bo mniej wątków konkuruje o jedną tablicę s_hist!"),
        ("Próg opłacalności T_total i efekt skali: ",
         "Przepustowość rdzenia GPU rośnie od 6,89 GB/s (10 MB) do 29,09 GB/s (180 MB = 116 GB/s log.). W T_total narzut kontekstu WDDM (~200 ms) faworyzuje OMP-Inc na 1 pliku, ale GPU pokonuje 1-wątkowy CPU już od N ≥ 25 MB i wygrywa w przetwarzaniu wsadowym."),
    ], size=8.7, space_after=4.0, bullet_color=C_GREEN)

    add_footer(s, 13, TOTAL)
    add_notes(s, (
        "SLAJD 13 (Analiza GPU 1/2 — Jan):\n"
        "• Wskazać 3 rzeczy bardzo cenione na PRiR:\n"
        "  1) Uczciwe pokazanie narzutu PCIe (H2D ~22,9 ms, D2H tylko ~0,23 ms bo wyników jest K=9152 double = 73 KB) oraz narzutu inicjalizacji kontekstu CUDA w WSL2/WDDM (~200 ms).\n"
        "  2) Dlaczego blok 1024 wątków pogarsza wynik? Na SM mieści się wtedy tylko 1 blok, a 32 warpy w obrębie jednego bloku tłoczą się przy tej samej tablicy s_hist[256].\n"
        "  3) Jak można by dalej przyspieszyć T_total w produkcji? Pamięć przypięta (cudaMallocHost / pinned memory) + potokowe strumienie CUDA (cudaStream_t) nakładające H2D na obliczenia oraz przetwarzanie wsadowe wielu plików w jednym kontekście."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 14: GPU ANALYSIS 2/2 — ATOMIC CONTENTION STUDY
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "13  •  Analiza wydajności GPU (2/2)",
               "Wpływ koncentracji bajtów na konkurencję operacji atomowych",
               "Badanie parametryczne na serii 9 plików ksym_{1..256}.bin (50 MB, W = 65536, S = 16384, blok = 256)",
               tag_color=C_RED, tag_bg=C_RED_BG)

    # Top: Full-width high-res atomic contention plot (ratio 2.567 fits beautifully!)
    add_picture_fit(s, plots / "atomic_contention_impact.png", 0.55, 1.38, 12.23, 3.24,
                    card=True, caption="Lewy panel: skrajne rozkłady k=1 vs k=256 względem W  •  Prawy panel: ciągła krzywa koncentracji 1/k dla k ∈ {1..256}")

    # Bottom Left: Compact k-symbol table
    add_table(s, 0.55, 4.74, 7.25, 2.20,
              ["Symbole k", "Kolizja 1/k", "H_teor", "CUDA Naive [ms]", "CUDA-Inc [ms]", "Log. GB/s (Inc)", "vs k=1 (N / Inc)"],
              [
                  ["**1 (0x00)**", "**100,0%**", "0,0", "15,195", "6,558", "30,46 GB/s", "1,00× / 1,00×"],
                  ["2", "50,0%", "1,0", "7,858", "4,027", "49,60 GB/s", "1,93× / 1,63×"],
                  ["4", "25,0%", "2,0", "5,397", "2,471", "80,82 GB/s", "2,82× / 2,65×"],
                  ["8", "12,5%", "3,0", "4,163", "1,799", "111,05 GB/s", "3,65× / 3,65×"],
                  ["**64**", "1,56%", "6,0", "3,946", "**1,425**", "**140,18 GB/s**", "3,85× / **4,60×**"],
                  ["**128**", "0,78%", "7,0", "**3,682**", "1,550", "128,90 GB/s", "**4,13×** / 4,23×"],
                  ["256 (losowy)", "0,39%", "8,0", "3,804", "2,200", "90,78 GB/s", "3,99× / 2,98×"],
              ],
              col_widths=[1.05, 0.88, 0.62, 1.15, 1.05, 1.25, 1.25],
              right_cols=(1, 2, 3, 4, 5, 6), header_bg=C_RED, font_size=8.0, header_size=8.3,
              highlight_rows=(4,))

    # Bottom Right: Hardware explanation & CPU vs GPU paradox
    add_card(s, 7.96, 4.74, 4.82, 2.20, "Mechanizm sprzętowy i paradoks CPU vs GPU", accent_color=C_RED)
    add_bullets(s, 8.12, 5.16, 4.50, 1.72, [
        ("Serializacja w 32 bankach shared memory: ",
         "Dla k=1 (same zera) wszystkie 8 warpów bloku uderzają w ten sam koszyk s_hist[0] ⇒ spowolnienie ponad 4× (15,20 vs 3,68 ms), mimo tylko 1 wywołania log₂ zamiast 256!"),
        ("Szybkie odblokowanie od k ≥ 8: ",
         "Już przy 2 symbolach zysk 1,93×; przy 8 symbolach (12,5% kolizji) osiągamy > 88% szczytowej przepustowości."),
        ("Dlaczego lekki spadek CUDA-Inc przy k = 256? ",
         "Wzrost liczby niezerowych koszyków zwiększa liczbę wywołań FP64 log₂(p) w każdym kroku przesunięcia (na GTX 1660 przepustowość FP64 to 1/32 FP32)."),
    ], size=8.7, space_after=3.5, bullet_color=C_RED)

    add_footer(s, 14, TOTAL)
    add_notes(s, (
        "SLAJD 14 (Kolizje atomowe GPU — Jan & Adrian):\n"
        "• To jeden z najciekawszych wyników eksperymentalnych w projekcie: paradoks CPU vs GPU! Na CPU plik samych zer liczy się NAJSZYBCIEJ (bo w shannon_entropy liczymy tylko 1 log2 zamiast 256), a na GPU plik samych zer liczy się PONAD 4x WOLNIEJ przez kolejkowanie atomicAdd na jednym adresie s_hist[0].\n"
        "• Świetny detal architektoniczny Turinga (GTX 1660): jednostek FP64 jest 32x mniej niż FP32, dlatego przy k=256 w bardzo szybkim CUDA-Incremental koszt 256 wywołań log2(double) staje się widoczny (spadek ze 140 GB/s przy k=64 do 91 GB/s przy k=256).\n"
        "• Jak można by jeszcze bardziej zbić kolizje przy k=1? Np. przez wielokrotne kopie histogramu (privatization / replication per warp) lub redukcję warp-level (__match_any_sync) przed atomicAdd."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 15: SEGMENTATION & BOUNDARY LOCALIZATION ACCURACY
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_LIGHT_BG)
    add_header(s, "14  •  Dokładność segmentacji",
               "Lokalizacja granic segmentów i skuteczność scalania (--events)",
               "100% wykrytych granic bez fałszywych alarmów (FP = 0) dla nakładających się okien S = W/4 na plikach 150 MB",
               tag_color=C_VIOLET, tag_bg=C_VIOLET_BG)

    # Left: Large vertical plot segmentation_accuracy.png (ratio 1.337 -> ideal for left 56% of slide!)
    add_picture_fit(s, plots / "segmentation_accuracy.png", 0.55, 1.38, 7.15, 5.56,
                    card=True, caption="Profil entropii (niebieski), surowe flagi (pomarańczowe), scalone zdarzenia (czerwone) i ground-truth (czarne)")

    # Right Top: Boundary localization summary table
    add_card(s, 7.86, 1.38, 4.92, 3.24, "Wyniki lokalizacji (próg T = 0,1)", accent_color=C_VIOLET)
    add_table(s, 7.98, 1.84, 4.68, 2.66,
              ["Plik (W / S)", "GT", "Surowe", "Zdarz.", "FP", "Śr. błąd"],
              [
                  ["mixed (1024 / 1024)", "2", "34", "31", "29", "192 B"],
                  ["**mixed (1024 / 256)**", "**2**", "8", "**2**", "**0**", "**64 B**"],
                  ["mixed (4096 / 1024)", "2", "9", "**2**", "**0**", "192 B"],
                  ["mixed (16K / 4K)", "2", "9", "**2**", "**0**", "960 B"],
                  ["mixed (64K / 16K)", "2", "10", "**2**", "**0**", "2240 B"],
                  ["**multi (1024 / 256)**", "**4**", "17", "**4**", "**0**", "**32 B**"],
                  ["multi (16K / 4K)", "4", "17", "**4**", "**0**", "1152 B"],
                  ["multi (64K / 16K)", "4", "18", "**4**", "**0**", "5760 B"],
              ],
              col_widths=[1.62, 0.40, 0.66, 0.60, 0.44, 0.96],
              right_cols=(1, 2, 3, 4, 5), header_bg=C_VIOLET, font_size=8.0, header_size=8.3,
              highlight_rows=(1, 5))

    # Right Bottom: Practical conclusions
    add_card(s, 7.86, 4.74, 4.92, 2.20, "Kluczowe obserwacje z detekcji", accent_color=C_BLUE)
    add_bullets(s, 8.02, 5.16, 4.60, 1.72, [
        ("Rola scalania (--events): ",
         "Redukuje 8–10 surowych flag w mixed do dokładnie 2 zdarzeń (i 17–18 flag w multi do dokładnie 4 zdarzeń) — Recall = 100%, Precision = 100%."),
        ("Precyzja podkrokowa (≤ 0,5–0,8 · S): ",
         "Dla W=1024, S=256 granice w pliku 150 000 000 B lokalizowane są ze średnim błędem zaledwie 32–64 bajtów!"),
        ("Kompromis W vs szum losowy: ",
         "Dla rozłącznych małych okien (W=1024, S=1024) i T=0,1 fluktuacje próbkowe dają 29 FP; optymalnie: W = 4096..16384 przy S = W/4."),
    ], size=8.8, space_after=4.0, bullet_color=C_BLUE)

    add_footer(s, 15, TOTAL)
    add_notes(s, (
        "SLAJD 15 (Dokładność segmentacji — Adrian):\n"
        "• Wykres pokazuje dwa różne pliki testowe: 3-segmentowy mixed_segmented.bin (granice 50 i 100 MB) oraz trudniejszy 5-segmentowy multi_segmented.bin o zmiennych długościach i 5 poziomach entropii (0, 2, 4.52, 8, 4 bit/B).\n"
        "• Podkreślić, dlaczego w pierwszym wierszu tabeli (W=1024, S=1024) pojawiło się 29 FP, a w drugim (W=1024, S=256) już 0 FP: przy nakładaniu S=W/4 okna dzielą 75% tych samych bajtów, więc skok szumu losowego między sąsiednimi oknami jest wygładzony i nie przekracza T=0.1!"
    ))

    # -------------------------------------------------------------------------
    # SLIDE 16: LIVE DEMO PLAN (demo.ps1 — DARK TERMINAL THEME)
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_DARK_BG)
    add_header(s, "15  •  Pokaz na żywo (Live Demo)",
               "Uruchomienie projektu na żywo poprzez skrypt demo.ps1",
               "Konteneryzacja Docker + NVIDIA Container Toolkit (c05-env: CUDA 12.6 + GCC 13.3 + OpenMPI 4.1.6)",
               tag_color=C_GREEN_LT, tag_bg=C_DARK_CARD, dark=True)

    # Left: Commands & Mode cards
    demo_cmd = (
        "# 1. Domyślny pokaz 6-krokowy na obronę (kompilacja + 4 backendy + wykres):\n"
        ".\\demo.ps1\n"
        "\n"
        "# 2. Automatyczna kompilacja i wszystkie 4 zestawy testów CTest (CPU/MPI/GPU):\n"
        ".\\demo.ps1 -Mode test\n"
        "\n"
        "# 3. Pełne odtworzenie wszystkich 134 pomiarów i 9 wykresów ze sprawozdania:\n"
        ".\\demo.ps1 -Mode reproduce\n"
        "\n"
        "# 4. Interaktywna powłoka bash w kontenerze z dostępem do GPU (--gpus all):\n"
        ".\\demo.ps1 -Mode shell"
    )
    add_code_block(s, 0.55, 1.42, 6.25, 2.75, demo_cmd,
                   filename="PowerShell  •  demo.ps1  (Dockerfile + NVIDIA GPU passthrough)",
                   lang="ps1", font_size=9.2, start_line=None, accent_color=C_GREEN_LT)

    # Left Bottom: Measured live demo preview numbers (24 MB demo fixture)
    add_card(s, 0.55, 4.32, 6.25, 2.58,
             "Co zobaczymy w terminalu? (zbiór demo 24 MB, W=4096, S=1024)",
             accent_color=C_SKY, bg_color=C_DARK_CARD, border_color=C_DARK_BORDER, title_color=C_DARK_TEXT)
    add_table(s, 0.72, 4.78, 5.91, 1.98,
              ["Krok pokazu", "Tryb / Zasoby", "T_core", "Zdarzenia"],
              [
                  ["[1/5] Reference CPU", "naive, 1 wątek", "228,2 ms", "2 granice (8, 16 MB)"],
                  ["[2/5] OpenMP CPU", "naive, 12 wątków", "45,4 ms (5,0×)", "2 granice"],
                  ["[3/5] OpenMP CPU", "incremental, 12 wątków", "26,9 ms (8,5×)", "2 granice"],
                  ["[4/5] MPI (halo)", "incremental, 4 procesy", "38,0 ms (6,0×)", "2 granice"],
                  ["[5/5] CUDA GTX 1660", "shared mem, blok 256", "**5,19 ms (44×)**", "2 granice"],
                  ["Krok 4: S = 64 (375K ok.)", "OMP-Naive vs OMP-Inc", "471 → **114 ms (4,1×)**", "374 937 okien"],
              ],
              col_widths=[1.62, 1.65, 1.38, 1.26], right_cols=(2,),
              header_bg=C_BLUE, font_size=8.2, header_size=8.5, highlight_rows=(4, 5))

    # Right: 6 numbered live demo steps
    add_card(s, 7.00, 1.42, 5.78, 5.48, "Przebieg 6 kroków pokazu .\\demo.ps1",
             accent_color=C_GREEN_LT, bg_color=C_DARK_CARD, border_color=C_DARK_BORDER, title_color=C_DARK_TEXT)

    demo_steps = [
        ("1", "Kompilacja CMake + Ninja (Release)",
         "Buduje wszystkie 4 binaria (entropy_ref, entropy_omp, entropy_mpi, entropy_cuda) z flagami -O3 -march=native -Wall -Wextra (0 ostrzeżeń).", C_SKY),
        ("2", "Weryfikacja zbiorów testowych",
         "Używa wygenerowanych zbiorów w datasets/ (lub deterministycznie generuje zbiory demo przez generate_datasets.py).", C_BLUE),
        ("3", "Porównanie 4 technologii na tym samym pliku",
         "Uruchamia kolejno Reference, OpenMP Naive, OpenMP Incremental, MPI (4 procesy z halo) oraz CUDA GPU (po rozgrzewce kontekstu).", C_GREEN_LT),
        ("4", "Demonstracja siły trybu Incremental (S = 64)",
         "Przetwarza 374 937 gęsto nakładających się okien: pokazuje >4× przyspieszenie samego algorytmu na tych samych 12 wątkach.", C_AMBER_LT),
        ("5", "Wydruk scalonych granic (demo_events.csv)",
         "Wyświetla wykryte zdarzenia zmiany z dokładnym estimated_offset trafiającym w 8 MB (7 998 976 B) i 16 MB (16 000 000 B).", RGBColor(196, 181, 253)),
        ("6", "Wizualizacja profilu 5-segmentowego pliku",
         "Uruchamia plot_entropy.py na multi_segmented.bin i zapisuje wykres results/logs/demo_entropy_profile.png.", C_SKY),
    ]
    for idx, (num, stitle, sdesc, scol) in enumerate(demo_steps):
        sy = 1.92 + idx * 0.81
        add_badge(s, 7.18, sy + 0.04, 0.36, 0.36, num, scol, C_DARK_BG, size=11)
        add_textbox(s, 7.66, sy, 4.95, 0.25, stitle, size=10.5, color=C_DARK_TEXT, bold=True, font=FONT_SEMI)
        add_textbox(s, 7.66, sy + 0.25, 4.95, 0.50, sdesc, size=8.8, color=C_DARK_MUTED, line_spacing=1.14)

    add_footer(s, 16, TOTAL, dark=True)
    add_notes(s, (
        "SLAJD 16 (Przejście do demo na żywo — Jan & Adrian):\n"
        "• W tym momencie przełączamy ekran na terminal PowerShell i uruchamiamy: .\\demo.ps1\n"
        "• Cały pokaz trwa kilkanaście sekund! Po zakończeniu warto otworzyć świeżo wygenerowany obraz results/logs/demo_entropy_profile.png.\n"
        "• Jeśli prowadzący poprosi o testy poprawności, uruchamiamy: .\\demo.ps1 -Mode test (wszystkie 4 zestawy CTest trwają ~88 s i kończą się 100% PASS)."
    ))

    # -------------------------------------------------------------------------
    # SLIDE 17: CONCLUSIONS, TEAM WORK SPLIT & Q&A (DARK CLOSING)
    # -------------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, C_DARK_BG)
    add_header(s, "16  •  Podsumowanie i zespół",
               "Wnioski końcowe oraz podział pracy w zespole 2-osobowym",
               "Kompletna, zweryfikowana i w pełni reprodukowalna realizacja projektu C05",
               tag_color=C_SKY, tag_bg=C_DARK_CARD, dark=True)

    # Top: 4 Conclusion Cards in a 2x2 grid
    conclusions = [
        ("1. Poprawność i powtarzalność (100%)", C_GREEN_LT,
         "Wszystkie 8 wariantów (4 technologie × 2 tryby) przechodzą 4/4 testy CTest i generują identyczne granice. "
         "Wersje CPU są bitowo identyczne (max_err = 0,0), a błąd GPU wynosi zaledwie 1,69×10⁻¹⁴ bit/B."),
        ("2. OpenMP i MPI — ponad 91% efektywności", C_SKY,
         "Prywatne histogramy w L1 cache (OpenMP) i podział blokowy z zakładką halo W−S (MPI) skalują się niemal liniowo "
         "do 6 rdzeni fizycznych (S(6) = 5,54× / 5,46×) i do 6,76–7,23× z SMT. Na 1 węźle OpenMP wygrywa w T_total."),
        ("3. CUDA GPU — rdzeń vs czas całkowity i atomiki", C_AMBER_LT,
         "W T_core GTX 1660 osiąga 69,1× (Naive) i 129,8× (Incremental) względem CPU (do 140,2 GB/s log.). "
         "Koncentracja k=1 spowalnia GPU >4× (kolizje atomic), a transfer PCIe i setup WDDM decydują o T_total dla 1 pliku."),
        ("4. Wkład rozszerzenia 2-osobowego (Incremental)", RGBColor(196, 181, 253),
         "Aktualizacja przyrostowa O(W + KS) skraca czas rdzenia o 1,78–2,28× dla S = W/4 (i >4× dla S = W/64) "
         "we wszystkich 4 technologiach bez utraty dokładności; scalanie --events daje 100% dokładności lokalizacji."),
    ]
    for idx, (ct, cc, cb) in enumerate(conclusions):
        cx = 0.55 + (idx % 2) * 6.22
        cy = 1.40 + (idx // 2) * 1.42
        add_card(s, cx, cy, 6.01, 1.30, ct, accent_color=cc,
                 bg_color=C_DARK_CARD, border_color=C_DARK_BORDER, title_color=C_DARK_TEXT, title_size=11.0)
        add_textbox(s, cx + 0.16, cy + 0.42, 5.69, 0.82, cb, size=9.4, color=C_DARK_MUTED, line_spacing=1.18)

    # Bottom: 2 Team Role Cards
    add_card(s, 0.55, 4.34, 6.01, 2.08, "Jan Wojnicki  (nr albumu: 147020)",
             accent_color=C_BLUE, bg_color=C_DARK_CARD, border_color=C_DARK_BORDER, title_color=C_SKY, title_size=12.0)
    add_bullets(s, 0.72, 4.78, 5.68, 1.56, [
        ("Wspólny moduł common/: ", "CLI, walidacja ścieżek, odczyt binarny, układ pełnych okien, zapis CSV."),
        ("CPU Reference & OpenMP Naive: ", "compute_naive(), src/reference.cpp oraz OMP z prywatnymi histogramami wątków."),
        ("CUDA Naive & Redukcja: ", "histogram_entropy_kernel, boundary_flags_kernel, drzewiasta redukcja block_entropy."),
        ("Dane i testy: ", "generate_datasets.py, verify_correctness.py, verify_cuda.py."),
    ], size=9.0, color=C_DARK_MUTED, bullet_color=C_SKY, space_after=3.5)

    add_card(s, 6.77, 4.34, 6.01, 2.08, "Adrian Sajdak  (nr albumu: 146992)",
             accent_color=C_VIOLET, bg_color=C_DARK_CARD, border_color=C_DARK_BORDER,
             title_color=RGBColor(196, 181, 253), title_size=12.0)
    add_bullets(s, 6.94, 4.78, 5.68, 1.56, [
        ("Rozszerzenie 2-osobowe (--mode incremental): ", "compute_incremental, OMP-Inc, MPI-Inc oraz CUDA incremental_entropy_kernel."),
        ("Backend MPI (src/mpi.cpp): ", "podział blokowy okien, nieblokująca dystrybucja halo (MPI_Isend/Recv) i MPI_Gatherv."),
        ("Scalanie granic i weryfikacja: ", "BoundaryEvent (--events), verify_mpi.py, check_cuda_model.py."),
        ("Eksperymenty i wykresy: ", "benchmark_runner.py, plot_results.py, plot_entropy.py."),
    ], size=9.0, color=C_DARK_MUTED, bullet_color=RGBColor(196, 181, 253), space_after=3.5)

    # Bottom Q&A banner
    add_rect(s, 0.55, 6.54, 12.23, 0.42, C_BLUE, rounded=True, radius=0.25)
    add_textbox(s, 0.70, 6.58, 11.93, 0.34,
                "Dziękujemy za uwagę!   •   Czas na pokaz na żywo (.\\demo.ps1) oraz pytania   •   Repozytorium: github.com/JanWojnicki/Projekt_C05",
                size=11.0, color=C_WHITE, bold=True, font=FONT_SEMI, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_footer(s, 17, TOTAL, dark=True)
    add_notes(s, (
        "SLAJD 17 (Podsumowanie i Q&A — Jan & Adrian):\n"
        "• Obaj autorzy uczestniczyli w przeglądzie kodu, testowaniu, pomiarach oraz redakcji sprawozdania i rozumieją całość rozwiązania.\n"
        "• Najczęstsze pytania komisji, na które jesteśmy gotowi:\n"
        "  1) Jak dalej przyspieszyć CUDA przy k=1? (warp-level aggregation / __match_any_sync lub kilka kopii histogramu per warp).\n"
        "  2) Jak ukryć koszt transferu PCIe H2D? (pinned host memory + double-buffering ze strumieniami CUDA / przetwarzanie kawałkami).\n"
        "  3) Co gdy plik ma np. 50 GB i nie mieści się w RAM/VRAM? (streaming/chunking po dużych partiach z zakładką W-S, dokładnie jak w naszym MPI!)."
    ))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    print(f"Saved presentation ({len(prs.slides)} slides): {output_path}")


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parents[1]
    build_presentation(repo_root, repo_root / "PREZENTACJA.pptx")
