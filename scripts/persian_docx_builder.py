"""
Persian DOCX Builder & Styler Module
------------------------------------
Author: Antigravity Persian DOCX Skill
Description: A high-quality Python utility to generate Microsoft Word (.docx)
documents with proper Persian RTL & LTR support, Vazirmatn font styling,
comprehensive Paragraph Indents and Spacing controls, custom color palettes,
and professional table/callout elements.
"""

import os
import sys

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from typing import List, Dict, Optional, Tuple, Any, Union
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn


# ==============================================================================
# 1. COLOR PALETTES DEFINITIONS
# ==============================================================================

PALETTES: Dict[str, Dict[str, str]] = {
    "navy": {
        "name": "سرمه‌ای کلاسیک سازمانی",
        "primary": "1E3A8A",      # Deep Navy
        "secondary": "3B82F6",    # Bright Blue
        "accent": "D97706",       # Amber Gold
        "text_dark": "1E293B",    # Slate 800
        "text_muted": "64748B",   # Slate 500
        "bg_light": "F1F5F9",     # Slate 100
        "border": "CBD5E1",       # Slate 300
        "callout_info": "E0F2FE",
        "callout_warn": "FEF3C7",
        "callout_succ": "DCFCE7",
        "callout_dang": "FEE2E2",
    },
    "emerald": {
        "name": "سبز زمردی مدرن",
        "primary": "065F46",      # Deep Emerald
        "secondary": "10B981",    # Emerald Green
        "accent": "F59E0B",       # Warm Amber
        "text_dark": "132E27",    # Dark Forest
        "text_muted": "6B7280",   # Gray 500
        "bg_light": "ECFDF5",     # Emerald 50
        "border": "A7F3D0",       # Emerald 200
        "callout_info": "E0F2FE",
        "callout_warn": "FEF3C7",
        "callout_succ": "D1FAE5",
        "callout_dang": "FEE2E2",
    },
    "ruby": {
        "name": "یاقوتی لوکس مدیریتی",
        "primary": "881337",      # Rose 900
        "secondary": "E11D48",    # Rose 600
        "accent": "CA8A04",       # Yellow 600
        "text_dark": "27181C",    # Dark Burgundy
        "text_muted": "71717A",   # Zinc 500
        "bg_light": "FFF1F2",     # Rose 50
        "border": "FECDD3",       # Rose 200
        "callout_info": "F0F9FF",
        "callout_warn": "FEF3C7",
        "callout_succ": "DCFCE7",
        "callout_dang": "FFE4E6",
    },
    "ocean": {
        "name": "اقیانوسی و فیروزه‌ای",
        "primary": "0F766E",      # Teal 700
        "secondary": "14B8A6",    # Teal 500
        "accent": "F97316",       # Orange 500
        "text_dark": "1E293B",    # Slate 800
        "text_muted": "64748B",   # Slate 500
        "bg_light": "F0FDFA",     # Teal 50
        "border": "CCFBF1",       # Teal 100
        "callout_info": "E0F2FE",
        "callout_warn": "FEF3C7",
        "callout_succ": "CCFBF1",
        "callout_dang": "FEE2E2",
    },
    "violet": {
        "name": "بنفش خلاقانه و نوآوری",
        "primary": "4C1D95",      # Violet 900
        "secondary": "8B5CF6",    # Violet 500
        "accent": "EC4899",       # Pink 500
        "text_dark": "1E1B4B",    # Indigo 950
        "text_muted": "6B7280",   # Gray 500
        "bg_light": "F5F3FF",     # Violet 50
        "border": "DDD6FE",       # Violet 200
        "callout_info": "EDE9FE",
        "callout_warn": "FEF3C7",
        "callout_succ": "DCFCE7",
        "callout_dang": "FEE2E2",
    },
    "slate": {
        "name": "طوسی زغالی مینیمال و تکنیکال",
        "primary": "334155",      # Slate 700
        "secondary": "64748B",    # Slate 500
        "accent": "2563EB",       # Blue 600
        "text_dark": "0F172A",    # Slate 900
        "text_muted": "94A3B8",   # Slate 400
        "bg_light": "F8FAFC",     # Slate 50
        "border": "E2E8F0",       # Slate 200
        "callout_info": "F1F5F9",
        "callout_warn": "FEF3C7",
        "callout_succ": "DCFCE7",
        "callout_dang": "FEE2E2",
    },
}


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert HEX color string (e.g. '1E3A8A') to docx RGBColor."""
    hex_str = hex_str.lstrip('#')
    return RGBColor(
        int(hex_str[0:2], 16),
        int(hex_str[2:4], 16),
        int(hex_str[4:6], 16)
    )


# ==============================================================================
# 2. OPENXML HELPER FUNCTIONS FOR PARAGRAPH INDENTS, SPACING & DIRECTION
# ==============================================================================

def set_paragraph_direction(p, direction: str = "right-to-left"):
    """
    Set paragraph direction to either 'right-to-left' (RTL) or 'left-to-right' (LTR).
    For RTL paragraphs, ensures <w:bidi/> is present in <w:pPr> ([<-¶] in Word ribbon).
    For LTR paragraphs, sets explicit <w:bidi w:val="0"/> in <w:pPr>.
    """
    pPr = p._element.get_or_add_pPr()
    bidi = pPr.find(qn('w:bidi'))
    
    dir_clean = direction.lower().strip()
    is_rtl = dir_clean in ["right-to-left", "rtl", "right_to_left"]
    
    if is_rtl:
        if bidi is None:
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
        else:
            bidi.attrib.pop(qn('w:val'), None)
    else:
        # Left-to-right explicit override
        if bidi is None:
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
        bidi.set(qn('w:val'), '0')


def set_run_direction_and_font(
    run,
    font_name: str = "Vazirmatn",
    ascii_font: str = "Segoe UI",
    direction: str = "right-to-left"
):
    """
    Configure Run-level direction (w:rtl) and complex script + Latin font mapping.
    """
    rPr = run._element.get_or_add_rPr()
    rtl = rPr.find(qn('w:rtl'))
    
    dir_clean = direction.lower().strip()
    is_rtl = dir_clean in ["right-to-left", "rtl", "right_to_left"]

    if is_rtl:
        if rtl is None:
            rtl = OxmlElement('w:rtl')
            rPr.append(rtl)
        rtl.set(qn('w:val'), '1')
    else:
        if rtl is not None:
            rtl.set(qn('w:val'), '0')

    # Configure font mapping
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)


def configure_paragraph_indents_and_spacing(
    p,
    direction: str = "right-to-left",
    align: Optional[WD_ALIGN_PARAGRAPH] = None,
    left_indent_cm: Optional[float] = None,
    right_indent_cm: Optional[float] = None,
    first_line_indent_cm: Optional[float] = None,
    space_before_pt: Optional[float] = None,
    space_after_pt: Optional[float] = None,
    line_spacing: Optional[float] = None,
    keep_with_next: bool = False
):
    """
    Comprehensive Word Paragraph Settings:
    - Direction: 'left-to-right' (LTR) or 'right-to-left' (RTL)
    - Indents: left, right, first-line
    - Spacing: before, after, line_spacing
    - Alignment: Left, Right, Center, Justify
    """
    dir_clean = direction.lower().strip()
    is_rtl = dir_clean in ["right-to-left", "rtl", "right_to_left"]

    # 1. Set Direction (adds <w:bidi/> for RTL)
    set_paragraph_direction(p, direction=direction)

    # 2. Set Alignment in OpenXML:
    # In Word, an RTL paragraph with <w:bidi/> naturally starts on the physical RIGHT.
    # If <w:jc w:val="right"/> is explicitly added, Word's BiDi engine inverts it to physical Left!
    # Therefore:
    # - For RTL Right alignment (or None): omit/remove <w:jc> to default to physical Right.
    # - For RTL Center alignment: set <w:jc w:val="center"/>.
    # - For RTL Justify alignment: set <w:jc w:val="both"/>.
    # - For RTL Left alignment: set <w:jc w:val="left"/>.
    pPr = p._element.get_or_add_pPr()
    jc = pPr.find(qn('w:jc'))
    if jc is not None:
        pPr.remove(jc)

    if is_rtl:
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
        elif align in [WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.DISTRIBUTE]:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'both')
            pPr.append(jc)
        elif align == WD_ALIGN_PARAGRAPH.LEFT:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'left')
            pPr.append(jc)
        # For align == WD_ALIGN_PARAGRAPH.RIGHT or None: omit <w:jc>, Word aligns to physical Right
    else:
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
        elif align == WD_ALIGN_PARAGRAPH.RIGHT:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'right')
            pPr.append(jc)
        elif align in [WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.DISTRIBUTE]:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'both')
            pPr.append(jc)
        else:  # LEFT or None
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'left')
            pPr.append(jc)

    pf = p.paragraph_format

    # 3. Indents
    if left_indent_cm is not None:
        pf.left_indent = Cm(left_indent_cm)
    if right_indent_cm is not None:
        pf.right_indent = Cm(right_indent_cm)
    if first_line_indent_cm is not None:
        pf.first_line_indent = Cm(first_line_indent_cm)

    # 4. Spacing
    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)
    if line_spacing is not None:
        pf.line_spacing = line_spacing

    if keep_with_next:
        pf.keep_with_next = True


def set_cell_shading(cell, color_hex: str):
    """Set background color of a table cell."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    """Set inner padding of a table cell in dxa (1 pt = 20 dxa)."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, border_color_hex="CBD5E1", sz="4", val="single"):
    """Set elegant, subtle borders on an entire table."""
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{border_color_hex}"/>\n'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{border_color_hex}"/>\n'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{border_color_hex}"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)


# ==============================================================================
# 3. PERSIAN DOCX BUILDER CLASS (RTL & LTR SUPPORT)
# ==============================================================================

class PersianDocxBuilder:
    """
    High-level builder for standardized Persian Word documents with full RTL & LTR controls.
    """

    def __init__(
        self,
        palette: str = "navy",
        font_name: str = "Vazirmatn",
        ascii_font: str = "Segoe UI",
        margin_cm: float = 2.5,
        default_direction: str = "right-to-left"
    ):
        self.palette_key = palette if palette in PALETTES else "navy"
        self.palette = PALETTES[self.palette_key]
        self.font_name = font_name
        self.ascii_font = ascii_font
        self.default_direction = default_direction

        self.doc = Document()
        self._init_document_defaults(margin_cm)

    def _init_document_defaults(self, margin_cm: float):
        """Configure page orientation, margins, and section RTL/LTR."""
        is_rtl = self.default_direction.lower() in ["right-to-left", "rtl"]
        for section in self.doc.sections:
            section.top_margin = Cm(margin_cm)
            section.bottom_margin = Cm(margin_cm)
            section.left_margin = Cm(margin_cm)
            section.right_margin = Cm(margin_cm)
            section.page_width = Cm(21.0)   # A4
            section.page_height = Cm(29.7)  # A4
            
            # Section Bidi
            sectPr = section._sectPr
            bidi = sectPr.find(qn('w:bidi'))
            if is_rtl:
                if bidi is None:
                    bidi = OxmlElement('w:bidi')
                    sectPr.append(bidi)
                bidi.set(qn('w:val'), '1')
            else:
                if bidi is not None:
                    bidi.set(qn('w:val'), '0')

        # Configure Normal style
        normal_style = self.doc.styles['Normal']
        normal_style.font.name = self.ascii_font
        normal_style.font.size = Pt(11.5)
        normal_style.font.color.rgb = hex_to_rgb(self.palette['text_dark'])

    def set_header_footer(
        self,
        header_text: Optional[str] = None,
        footer_text: Optional[str] = None,
        direction: Optional[str] = None,
        header_align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
        footer_align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER
    ):
        """Set header and footer texts with center alignment by default."""
        dir_val = direction if direction else self.default_direction
        section = self.doc.sections[0]
        
        if header_text:
            header = section.header
            hp = header.paragraphs[0]
            configure_paragraph_indents_and_spacing(
                hp,
                direction=dir_val,
                align=header_align
            )
            hrun = hp.add_run(header_text)
            hrun.font.size = Pt(9)
            hrun.font.color.rgb = hex_to_rgb(self.palette['text_muted'])
            set_run_direction_and_font(hrun, self.font_name, self.ascii_font, direction=dir_val)

        if footer_text:
            footer = section.footer
            fp = footer.paragraphs[0]
            configure_paragraph_indents_and_spacing(
                fp,
                direction=dir_val,
                align=footer_align
            )
            frun = fp.add_run(footer_text)
            frun.font.size = Pt(9)
            frun.font.color.rgb = hex_to_rgb(self.palette['text_muted'])
            set_run_direction_and_font(frun, self.font_name, self.ascii_font, direction=dir_val)

    # --------------------------------------------------------------------------
    # Titles & Headings
    # --------------------------------------------------------------------------

    def add_title(
        self,
        title: str,
        subtitle: Optional[str] = None,
        author: Optional[str] = None,
        date: Optional[str] = None,
        direction: Optional[str] = None,
        align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
        meta_align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.RIGHT
    ):
        """Add a polished Persian document header or title block (Title/Subtitle centered, Meta right-aligned by default)."""
        dir_val = direction if direction else self.default_direction
        align_val = align

        # Top spacing
        p_top = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p_top, direction=dir_val, align=align_val, space_before_pt=18, space_after_pt=4
        )

        # Title text
        p_title = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p_title, direction=dir_val, align=align_val, space_before_pt=0, space_after_pt=8
        )
        
        run_title = p_title.add_run(title)
        run_title.font.size = Pt(24)
        run_title.bold = True
        run_title.font.color.rgb = hex_to_rgb(self.palette['primary'])
        set_run_direction_and_font(run_title, self.font_name, self.ascii_font, direction=dir_val)

        # Subtitle
        if subtitle:
            p_sub = self.doc.add_paragraph()
            configure_paragraph_indents_and_spacing(
                p_sub, direction=dir_val, align=align_val, space_before_pt=0, space_after_pt=12
            )
            
            run_sub = p_sub.add_run(subtitle)
            run_sub.font.size = Pt(14)
            run_sub.font.color.rgb = hex_to_rgb(self.palette['secondary'])
            set_run_direction_and_font(run_sub, self.font_name, self.ascii_font, direction=dir_val)

        # Meta (Author, Date) - Right-aligned as requested
        if author or date:
            meta_items = []
            if author:
                meta_items.append(f"نویسنده / تدوین: {author}" if dir_val in ["right-to-left", "rtl"] else f"Author: {author}")
            if date:
                meta_items.append(f"تاریخ: {date}" if dir_val in ["right-to-left", "rtl"] else f"Date: {date}")
            meta_text = " | ".join(meta_items)

            p_meta = self.doc.add_paragraph()
            configure_paragraph_indents_and_spacing(
                p_meta, direction=dir_val, align=meta_align, space_before_pt=0, space_after_pt=20
            )

            run_meta = p_meta.add_run(meta_text)
            run_meta.font.size = Pt(10)
            run_meta.font.color.rgb = hex_to_rgb(self.palette['text_muted'])
            set_run_direction_and_font(run_meta, self.font_name, self.ascii_font, direction=dir_val)

        self.add_divider(color_hex=self.palette['secondary'], size_pt=1.5)

    def add_heading_1(self, text: str, direction: Optional[str] = None):
        """Add Level 1 Heading (18pt Bold, Primary Color, Right-aligned)."""
        dir_val = direction if direction else self.default_direction
        align_val = WD_ALIGN_PARAGRAPH.RIGHT if dir_val in ["right-to-left", "rtl"] else WD_ALIGN_PARAGRAPH.LEFT
        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p, direction=dir_val, align=align_val, space_before_pt=18, space_after_pt=6, keep_with_next=True
        )

        run = p.add_run(text)
        run.font.size = Pt(18)
        run.bold = True
        run.font.color.rgb = hex_to_rgb(self.palette['primary'])
        set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_heading_2(
        self,
        text: str,
        direction: Optional[str] = None,
        align: Optional[WD_ALIGN_PARAGRAPH] = None
    ):
        """Add Level 2 Heading (15pt Bold, Secondary Color, Right-aligned by default)."""
        dir_val = direction if direction else self.default_direction
        default_align = WD_ALIGN_PARAGRAPH.RIGHT if dir_val in ["right-to-left", "rtl"] else WD_ALIGN_PARAGRAPH.LEFT
        align_val = align if align is not None else default_align
        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p, direction=dir_val, align=align_val, space_before_pt=14, space_after_pt=4, keep_with_next=True
        )

        run = p.add_run(text)
        run.font.size = Pt(15)
        run.bold = True
        run.font.color.rgb = hex_to_rgb(self.palette['secondary'])
        set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_heading_3(self, text: str, direction: Optional[str] = None):
        """Add Level 3 Heading (13pt SemiBold, Dark Color)."""
        dir_val = direction if direction else self.default_direction
        align_val = WD_ALIGN_PARAGRAPH.RIGHT if dir_val in ["right-to-left", "rtl"] else WD_ALIGN_PARAGRAPH.LEFT
        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p, direction=dir_val, align=align_val, space_before_pt=10, space_after_pt=3, keep_with_next=True
        )

        run = p.add_run(text)
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
        set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_heading_4(self, text: str, direction: Optional[str] = None):
        """Add Level 4 Heading (11.5pt Bold, Slate Color)."""
        dir_val = direction if direction else self.default_direction
        align_val = WD_ALIGN_PARAGRAPH.RIGHT if dir_val in ["right-to-left", "rtl"] else WD_ALIGN_PARAGRAPH.LEFT
        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p, direction=dir_val, align=align_val, space_before_pt=8, space_after_pt=2, keep_with_next=True
        )

        run = p.add_run(text)
        run.font.size = Pt(11.5)
        run.bold = True
        run.font.color.rgb = hex_to_rgb(self.palette['text_muted'])
        set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)
        return p

    # --------------------------------------------------------------------------
    # Body Paragraphs with Full Indents and Spacing & Direction Control
    # --------------------------------------------------------------------------

    def add_paragraph(
        self,
        text: str,
        bold_prefix: Optional[str] = None,
        direction: Optional[str] = None,
        align: Optional[WD_ALIGN_PARAGRAPH] = None,
        left_indent_cm: Optional[float] = None,
        right_indent_cm: Optional[float] = None,
        first_line_indent_cm: Optional[float] = None,
        space_before_pt: float = 0.0,
        space_after_pt: float = 6.0,
        line_spacing: float = 1.3
    ):
        """
        Add a body paragraph with complete Word settings:
        - direction: 'right-to-left' (RTL) or 'left-to-right' (LTR)
        - indents: left, right, first-line
        - spacing: before, after, line_spacing (1.3x recommended for Persian)
        - alignment: Justify, Right, Left, Center
        """
        dir_val = direction if direction else self.default_direction
        default_align = WD_ALIGN_PARAGRAPH.JUSTIFY if dir_val in ["right-to-left", "rtl"] else WD_ALIGN_PARAGRAPH.LEFT
        align_val = align if align is not None else default_align

        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p,
            direction=dir_val,
            align=align_val,
            left_indent_cm=left_indent_cm,
            right_indent_cm=right_indent_cm,
            first_line_indent_cm=first_line_indent_cm,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_spacing=line_spacing
        )

        if bold_prefix:
            run_pre = p.add_run(bold_prefix + " ")
            run_pre.font.size = Pt(11.5)
            run_pre.bold = True
            run_pre.font.color.rgb = hex_to_rgb(self.palette['primary'])
            set_run_direction_and_font(run_pre, self.font_name, self.ascii_font, direction=dir_val)

        run = p.add_run(text)
        run.font.size = Pt(11.5)
        run.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
        set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_ltr_paragraph(
        self,
        text: str,
        bold_prefix: Optional[str] = None,
        align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
        left_indent_cm: Optional[float] = None,
        right_indent_cm: Optional[float] = None,
        first_line_indent_cm: Optional[float] = None,
        space_before_pt: float = 0.0,
        space_after_pt: float = 6.0,
        line_spacing: float = 1.15
    ):
        """
        Convenience method to add an explicitly Left-To-Right (LTR) English paragraph
        with Word indents and spacing.
        """
        return self.add_paragraph(
            text=text,
            bold_prefix=bold_prefix,
            direction="left-to-right",
            align=align,
            left_indent_cm=left_indent_cm,
            right_indent_cm=right_indent_cm,
            first_line_indent_cm=first_line_indent_cm,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_spacing=line_spacing
        )

    def add_bullet_item(
        self,
        text: str,
        bold_title: Optional[str] = None,
        level: int = 0,
        direction: Optional[str] = None
    ):
        """Add a clean bullet list item with RTL/LTR indentation."""
        dir_val = direction if direction else self.default_direction
        is_rtl = dir_val in ["right-to-left", "rtl"]
        align_val = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT

        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p,
            direction=dir_val,
            align=align_val,
            right_indent_cm=(0.6 * (level + 1)) if is_rtl else 0,
            left_indent_cm=(0.6 * (level + 1)) if not is_rtl else 0,
            space_before_pt=2,
            space_after_pt=3,
            line_spacing=1.25
        )

        bullet_symbol = "▪  " if level == 0 else "▫  "
        run_bullet = p.add_run(bullet_symbol)
        run_bullet.font.size = Pt(10)
        run_bullet.font.color.rgb = hex_to_rgb(self.palette['secondary'])
        set_run_direction_and_font(run_bullet, self.font_name, self.ascii_font, direction=dir_val)

        if bold_title:
            run_bold = p.add_run(bold_title + ": ")
            run_bold.font.size = Pt(11)
            run_bold.bold = True
            run_bold.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
            set_run_direction_and_font(run_bold, self.font_name, self.ascii_font, direction=dir_val)

        run_text = p.add_run(text)
        run_text.font.size = Pt(11)
        run_text.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
        set_run_direction_and_font(run_text, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_numbered_item(
        self,
        number: str,
        text: str,
        bold_title: Optional[str] = None,
        direction: Optional[str] = None
    ):
        """Add a numbered list item."""
        dir_val = direction if direction else self.default_direction
        is_rtl = dir_val in ["right-to-left", "rtl"]
        align_val = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT

        p = self.doc.add_paragraph()
        configure_paragraph_indents_and_spacing(
            p,
            direction=dir_val,
            align=align_val,
            right_indent_cm=0.6 if is_rtl else 0,
            left_indent_cm=0.6 if not is_rtl else 0,
            space_before_pt=2,
            space_after_pt=4,
            line_spacing=1.25
        )

        run_num = p.add_run(f"{number}.  ")
        run_num.font.size = Pt(11)
        run_num.bold = True
        run_num.font.color.rgb = hex_to_rgb(self.palette['primary'])
        set_run_direction_and_font(run_num, self.font_name, self.ascii_font, direction=dir_val)

        if bold_title:
            run_bold = p.add_run(bold_title + ": ")
            run_bold.font.size = Pt(11)
            run_bold.bold = True
            run_bold.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
            set_run_direction_and_font(run_bold, self.font_name, self.ascii_font, direction=dir_val)

        run_text = p.add_run(text)
        run_text.font.size = Pt(11)
        run_text.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
        set_run_direction_and_font(run_text, self.font_name, self.ascii_font, direction=dir_val)
        return p

    def add_divider(self, color_hex: Optional[str] = None, size_pt: float = 1.0):
        """Add a clean horizontal divider."""
        c = color_hex if color_hex else self.palette['border']
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(8)
        pPr = p._element.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>\n'
            f'  <w:bottom w:val="single" w:sz="{int(size_pt * 8)}" w:space="1" w:color="{c}"/>\n'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)

    # --------------------------------------------------------------------------
    # Callout / Note Boxes
    # --------------------------------------------------------------------------

    def add_callout(
        self,
        text: str,
        title: Optional[str] = None,
        callout_type: str = "info",  # 'info', 'warning', 'success', 'danger', 'quote'
        direction: Optional[str] = None
    ):
        """
        Add a highlighted callout box with border accent and tinted background.
        """
        dir_val = direction if direction else self.default_direction
        is_rtl = dir_val in ["right-to-left", "rtl"]

        bg_colors = {
            "info": self.palette['callout_info'],
            "warning": self.palette['callout_warn'],
            "success": self.palette['callout_succ'],
            "danger": self.palette['callout_dang'],
            "quote": self.palette['bg_light'],
        }
        border_colors = {
            "info": self.palette['secondary'],
            "warning": self.palette['accent'],
            "success": "10B981",
            "danger": "EF4444",
            "quote": self.palette['primary'],
        }
        icons = {
            "info": "ℹ️ ",
            "warning": "⚠️ ",
            "success": "✅ ",
            "danger": "⛔ ",
            "quote": "💬 ",
        }

        bg_hex = bg_colors.get(callout_type, self.palette['bg_light'])
        border_hex = border_colors.get(callout_type, self.palette['primary'])

        # Create single-cell table for callout box
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        tblPr = table._element.xpath('w:tblPr')[0]
        if is_rtl:
            tblBidi = OxmlElement('w:tblBidi')
            tblBidi.set(qn('w:val'), '1')
            tblPr.append(tblBidi)

        # Style cell
        cell = table.cell(0, 0)
        cell.width = Cm(16.0)
        set_cell_shading(cell, bg_hex)
        set_cell_margins(cell, top=160, bottom=160, left=200, right=200)

        # Border on the leading edge (right for RTL, left for LTR)
        tcPr = cell._element.get_or_add_tcPr()
        border_pos = "right" if is_rtl else "left"
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>\n'
            f'  <w:{border_pos} w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>\n'
            f'  <w:top w:val="none"/>\n'
            f'  <w:{"left" if is_rtl else "right"} w:val="none"/>\n'
            f'  <w:bottom w:val="none"/>\n'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        # Add content inside callout
        p = cell.paragraphs[0]
        configure_paragraph_indents_and_spacing(
            p,
            direction=dir_val,
            align=WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT,
            space_before_pt=0,
            space_after_pt=2,
            line_spacing=1.3
        )

        if title:
            icon = icons.get(callout_type, "")
            run_title = p.add_run(f"{icon}{title}\n")
            run_title.font.size = Pt(11)
            run_title.bold = True
            run_title.font.color.rgb = hex_to_rgb(border_hex)
            set_run_direction_and_font(run_title, self.font_name, self.ascii_font, direction=dir_val)

        run_body = p.add_run(text)
        run_body.font.size = Pt(10.5)
        run_body.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
        set_run_direction_and_font(run_body, self.font_name, self.ascii_font, direction=dir_val)

        # Spacing after callout table
        p_after = self.doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(4)

    # --------------------------------------------------------------------------
    # Tables
    # --------------------------------------------------------------------------

    def add_table(
        self,
        headers: List[str],
        rows: List[List[str]],
        col_widths: Optional[List[float]] = None,
        zebra_striping: bool = True,
        direction: Optional[str] = None
    ):
        """
        Add a beautifully styled table with RTL or LTR support.
        - Header row has palette primary background with white text.
        - Alternate rows have soft zebra background.
        """
        dir_val = direction if direction else self.default_direction
        is_rtl = dir_val in ["right-to-left", "rtl"]

        # Build the table as a left-to-right grid and reverse the column order for RTL,
        # so the first logical column (e.g. 'نقش') sits on the right in every viewer.
        if is_rtl:
            headers = list(headers)[::-1]
            rows = [list(r)[::-1] for r in rows]
            if col_widths:
                col_widths = list(col_widths)[::-1]
        table = self.doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        tblPr = table._element.xpath('w:tblPr')[0]

        # Style Header Row
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            cell = hdr_cells[i]
            set_cell_shading(cell, self.palette['primary'])
            set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            p = cell.paragraphs[0]
            configure_paragraph_indents_and_spacing(
                p, direction=dir_val, align=WD_ALIGN_PARAGRAPH.CENTER, space_before_pt=0, space_after_pt=0
            )

            run = p.add_run(header_text)
            run.font.size = Pt(10.5)
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)

        # Style Data Rows
        for r_idx, row_data in enumerate(rows):
            row_cells = table.rows[r_idx + 1].cells
            is_even = (r_idx % 2 == 1)
            row_bg = self.palette['bg_light'] if (is_even and zebra_striping) else "FFFFFF"

            for c_idx, cell_value in enumerate(row_data):
                cell = row_cells[c_idx]
                if row_bg != "FFFFFF":
                    set_cell_shading(cell, row_bg)
                set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p = cell.paragraphs[0]
                # Center numbers/short text, justify/align longer text
                cell_str = str(cell_value)
                if len(cell_str) < 15:
                    align_c = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    align_c = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT

                configure_paragraph_indents_and_spacing(
                    p, direction=dir_val, align=align_c, space_before_pt=0, space_after_pt=0
                )

                run = p.add_run(cell_str)
                run.font.size = Pt(10)
                run.font.color.rgb = hex_to_rgb(self.palette['text_dark'])
                set_run_direction_and_font(run, self.font_name, self.ascii_font, direction=dir_val)

        # Set Column Widths if provided
        if col_widths and len(col_widths) == len(headers):
            for row in table.rows:
                for c_idx, width_cm in enumerate(col_widths):
                    row.cells[c_idx].width = Cm(width_cm)
            # Also set the table grid and a fixed layout so Word and LibreOffice honor the widths
            grid_cols = table._tbl.tblGrid.findall(qn('w:gridCol'))
            for gc, width_cm in zip(grid_cols, col_widths):
                gc.set(qn('w:w'), str(int(Cm(width_cm) / 635)))
            layout = OxmlElement('w:tblLayout')
            layout.set(qn('w:type'), 'fixed')
            tblPr.append(layout)

        # Subtle borders
        set_table_borders(table, border_color_hex=self.palette['border'], sz="4")

        # Spacing after table
        p_after = self.doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)

    def add_key_value_table(self, data_dict: Dict[str, str], title: Optional[str] = None):
        """Add a sleek 2-column key-value summary table."""
        if title:
            self.add_heading_3(title)

        headers = ["مشخصه / عنوان", "مقدار / توضیحات"]
        rows = [[k, v] for k, v in data_dict.items()]
        self.add_table(headers=headers, rows=rows, col_widths=[5.0, 11.0], zebra_striping=True)

    # --------------------------------------------------------------------------
    # Code Blocks (Always Left-To-Right)
    # --------------------------------------------------------------------------

    def add_code_block(self, code_text: str, language: Optional[str] = None):
        """Add a clean code / monospace technical block (Left-To-Right)."""
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Cm(16.0)
        set_cell_shading(cell, "1E293B")  # Dark editor background
        set_cell_margins(cell, top=140, bottom=140, left=160, right=160)

        p = cell.paragraphs[0]
        configure_paragraph_indents_and_spacing(
            p, direction="left-to-right", align=WD_ALIGN_PARAGRAPH.LEFT, space_before_pt=0, space_after_pt=0
        )

        run = p.add_run(code_text)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(241, 245, 249)

        # Spacing after code block
        p_after = self.doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)

    def add_page_break(self):
        """Insert a page break to advance to the next page."""
        return self.doc.add_page_break()

    def save(self, filepath: str) -> str:
        """Save the generated .docx file to disk. If file is locked by Word, saves to a fallback path."""
        directory = os.path.dirname(filepath)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
        try:
            self.doc.save(filepath)
            print(f"✅ Document successfully created: {filepath}")
            return filepath
        except PermissionError:
            base, ext = os.path.splitext(filepath)
            fallback_path = f"{base}_updated{ext}"
            self.doc.save(fallback_path)
            print(f"⚠️ Original file was locked by Word. Saved to fallback: {fallback_path}")
            return fallback_path


# ==============================================================================
# 4. CLI INTERFACE
# ==============================================================================

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Persian DOCX Builder CLI with RTL & LTR Paragraph Controls")
    parser.add_argument("--output", "-o", default="persian_document.docx", help="Output .docx file path")
    parser.add_argument("--palette", "-p", choices=list(PALETTES.keys()), default="navy", help="Color palette")
    parser.add_argument("--direction", "-d", choices=["right-to-left", "left-to-right", "rtl", "ltr"], default="right-to-left", help="Default paragraph direction")
    parser.add_argument("--title", "-t", default="سند استاندارد فارسی و انگلیسی", help="Document title")
    parser.add_argument("--subtitle", "-s", default="پشتیبانی کامل از Paragraph Indents, Spacing و Direction", help="Document subtitle")
    args = parser.parse_args()

    builder = PersianDocxBuilder(palette=args.palette, default_direction=args.direction)
    builder.set_header_footer("سند رسمی - تنظیمات پاراگراف", "صفحه ۱ از ۱")
    builder.add_title(args.title, subtitle=args.subtitle, author="دستیار Antigravity", date="۱۴۰۴/۱۲/۰۱")
    
    builder.add_heading_1("۱. معرفی قابلیت‌های تنظیم پاراگراف (Indents & Spacing)")
    builder.add_paragraph(
        "در این نسخه، کلیه تنظیمات پنجره Paragraph مایکروسافت ورد شامل Indents (چپ، راست، خط اول)، "
        "Spacing (قبل، بعد، و فاصله بین خطوط) و Direction (تنظیم Left-to-Right و Right-to-Left) به صورت کامل پیاده‌سازی شده‌اند.",
        bold_prefix="پشتیبانی کامل:"
    )

    # Example of Left-to-Right English paragraph
    builder.add_heading_2("2. Left-to-Right Paragraph Example (LTR)")
    builder.add_ltr_paragraph(
        "This is an English paragraph configured with direction: left-to-right, left alignment, "
        "and standard English typography spacing. It seamlessly integrates into Persian documents.",
        bold_prefix="English Section:"
    )

    builder.add_callout(
        text="امکان تغییر جهت پاراگراف به Left-to-Right برای درج متون انگلیسی، فرمول‌ها، ارجاعات و مستندات فنی فراهم است.",
        title="انعطاف‌پذیری جهات متن",
        callout_type="info"
    )

    builder.save(args.output)


if __name__ == "__main__":
    main()
