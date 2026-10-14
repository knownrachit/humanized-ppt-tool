#!/usr/bin/env python3
"""
humanized_ppt_tool.py
======================
A reusable generator for slide decks that follow a fixed, hand-picked design
system instead of looking AI-generated. Two modes:

  1. SCRATCH MODE  - no template given. Builds slides from a curated palette
     + layout library (title, stat, icon-rows, two-column, closing).
  2. TEMPLATE MODE - a .pptx/.potx template is supplied. The tool reuses the
     template's own slide layouts (so it inherits the template's fonts,
     colors and branding) and fills them with your content instead of
     inventing a new look.

Design rules baked in (do not change without a reason):
  - Max 2 font families total (one for headers, one for body).
  - Body text >= 24pt, always.
  - Max 3 colors total, including the background.
  - No decorative boxes, accent stripes, or underlines under titles.
  - One consistent layout grammar reused across every slide.
  - Left-aligned body copy; centered titles only.
  - Every slide carries at least one non-text visual element.

Usage
-----
    from humanized_ppt_tool import Deck

    deck = Deck()                          # scratch mode, default palette
    deck.add_title_slide("Q3 Retro", "What worked, what didn't")
    deck.add_stat_slide("94%", "of tickets closed within SLA")
    deck.add_icon_rows_slide("Key wins", [
        ("1", "Faster onboarding", "Median time-to-first-ticket dropped 40%"),
        ("2", "Fewer escalations", "Escalation rate down to 3.1%"),
    ])
    deck.save("out.pptx")

    # Template mode
    deck = Deck(template="brand_template.pptx")
    deck.add_title_slide("Q3 Retro", "What worked, what didn't")
    deck.save("out.pptx")

Run this file directly for a self-test that builds three demo decks.
"""

from __future__ import annotations
import copy
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# Palettes: exactly 3 colors each (primary / background / accent), chosen to
# fit a topic rather than defaulting to generic blue. Pick one that fits the
# subject matter when you build a deck.
#
# "navy_orange" and "campus_blue" are lifted directly (exact hex values) from
# two decks the user liked — not generic inspiration, the actual palette.
# --------------------------------------------------------------------------
PALETTES = {
    "midnight":     {"primary": "1E2761", "background": "FFFFFF", "accent": "CADCFC"},
    "forest":       {"primary": "2C5F2D", "background": "FFFFFF", "accent": "97BC62"},
    "charcoal":     {"primary": "212121", "background": "FFFFFF", "accent": "36454F"},
    "terracotta":   {"primary": "B85042", "background": "FFFFFF", "accent": "E7E8D1"},
    "teal":         {"primary": "028090", "background": "FFFFFF", "accent": "02C39A"},
    "berry":        {"primary": "6D2E46", "background": "FFFFFF", "accent": "ECE2D0"},
    # lifted from the "Data vs Intuition" deck: navy + orange, line-icon style
    "navy_orange":  {"primary": "1B3B6F", "background": "FFFFFF", "accent": "E8873A"},
    # lifted from the "A-Day: Apex" deck: blue + soft-tint sidebar boxes
    "campus_blue":  {"primary": "269FD6", "background": "FFFFFF", "accent": "DDE3EF"},
}

# Safe-list fonts only (render true-to-width, ship with Office).
HEADER_FONT = "Bookman Old Style"   # serif with personality, for titles/headers
BODY_FONT = "Calibri"               # clean sans, for everything else

# Per-palette font overrides. "navy_orange" reuses the source deck's own pair
# (Cambria + Calibri — both safe-list). "campus_blue" was built in Poppins,
# which isn't on the safe list, so QA text-fit there is approximate — sized
# with extra slack in the layouts that use it.
PALETTE_FONTS = {
    "navy_orange": {"header": "Cambria", "body": "Calibri"},
    "campus_blue": {"header": "Poppins", "body": "Poppins"},
}


def _fonts_for(pal_name: str):
    f = PALETTE_FONTS.get(pal_name, {})
    return f.get("header", HEADER_FONT), f.get("body", BODY_FONT)


# Body text on a dark-gray-on-white palette (deck2 style) reads better a
# shade softer than pure black; keep everything else as user-stated colors.
TEXT_DARK = "323232"

TITLE_SIZE = Pt(40)
SECTION_SIZE = Pt(28)
BODY_SIZE = Pt(24)          # user's hard floor for body text
CAPTION_SIZE = Pt(16)       # still bumped up from the usual 10-12pt default
STAT_SIZE = Pt(96)

MARGIN = Inches(0.6)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _hex(c: str) -> RGBColor:
    return RGBColor.from_string(c)


def _text_block_height(text: str, width_in: float, font_pt: float, line_spacing=1.15, pad_in=0.12,
                        bold=False):
    """Rough line-count estimate so callers can size/position boxes without
    text overflowing past the box edge into whatever sits below it. Not a
    real text-metrics engine — pad generously (default ~0.12in) rather than
    trust this to the pixel, especially for QA-unreliable fonts. A 1.2x
    safety factor is applied on top: bold/serif headers in particular
    render taller line-boxes than a plain 1.15x-leading estimate predicts,
    and overlap is a worse failure than a little extra whitespace."""
    import math
    # bold serif headers render wider per character than regular body copy
    char_w_in = (font_pt / 72) * (0.64 if bold else 0.56)
    chars_per_line = max(8, int(width_in / char_w_in))
    lines = 0
    for para in text.split("\n"):
        lines += max(1, math.ceil(len(para) / chars_per_line))
    line_h_in = (font_pt / 72) * line_spacing
    return (lines * line_h_in + pad_in) * 1.2


def _set_background(slide, color_hex: str):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _hex(color_hex)


def _no_line(shape):
    shape.line.fill.background()


def _textbox(slide, l, t, w, h, text, size, color, font=BODY_FONT,
             bold=False, italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.15):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.size = size
        run.font.name = font
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = _hex(color)
    return box


def _circle_badge(slide, cx, cy, diameter, label, fill_hex, text_hex):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, diameter, diameter)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _hex(fill_hex)
    _no_line(shp)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.name = HEADER_FONT
    run.font.color.rgb = _hex(text_hex)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return shp


# --------------------------------------------------------------------------
# Line-icon "sketches" — simple outlined vector icons built from primitive
# shapes, in the spirit of the reference decks (outlined database/bulb/person
# glyphs, not stock clip-art or photos).
# --------------------------------------------------------------------------

def _outline_shape(slide, shape_type, x, y, w, h, line_hex, weight_pt=2.25):
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    shp.fill.background()
    shp.line.color.rgb = _hex(line_hex)
    shp.line.width = Pt(weight_pt)
    shp.shadow.inherit = False
    return shp


def icon_database(slide, x, y, size, color):
    """Cylinder / stacked-disc icon — matches the reference deck's 'Data' glyph."""
    return _outline_shape(slide, MSO_SHAPE.CAN, x, y, size, size, color)


def icon_bulb(slide, x, y, size, color):
    """Lightbulb — bulb (oval) over a small base (rectangle), outline only."""
    bulb_h = size * 0.72
    _outline_shape(slide, MSO_SHAPE.OVAL, x, y, size, bulb_h, color)
    base_w = size * 0.34
    _outline_shape(slide, MSO_SHAPE.RECTANGLE,
                    x + (size - base_w) / 2, y + bulb_h - Emu(30000),
                    base_w, size * 0.14, color, weight_pt=1.75)


def icon_person(slide, x, y, size, color):
    """Simple avatar — head (oval) over shoulders (trapezoid), outline only."""
    head_d = size * 0.4
    _outline_shape(slide, MSO_SHAPE.OVAL, x + (size - head_d) / 2, y, head_d, head_d, color)
    body_w = size * 0.8
    body_h = size * 0.42
    shp = _outline_shape(slide, MSO_SHAPE.TRAPEZOID,
                          x + (size - body_w) / 2, y + head_d * 0.92, body_w, body_h, color)
    shp.rotation = 180


def icon_target(slide, x, y, size, color):
    """Concentric-circle bullseye — matches the reference deck's 'decision' glyph."""
    _outline_shape(slide, MSO_SHAPE.OVAL, x, y, size, size, color, weight_pt=2.25)
    m1 = size * 0.28
    _outline_shape(slide, MSO_SHAPE.OVAL, x + m1, y + m1, size - 2 * m1, size - 2 * m1, color, weight_pt=2)
    dot = size * 0.14
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + (size - dot) / 2, y + (size - dot) / 2, dot, dot)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _hex(color)
    _no_line(shp)
    shp.shadow.inherit = False


def icon_arrow_right(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _hex(color)
    _no_line(shp)
    shp.shadow.inherit = False
    return shp


def icon_badge_glyph(slide, x, y, size, glyph, color):
    """Checkmark / cross badge — plain glyph in an outlined circle, matching
    the reference deck's check/x bullet icons."""
    _outline_shape(slide, MSO_SHAPE.OVAL, x, y, size, size, color, weight_pt=2)
    box = slide.shapes.add_textbox(x, y, size, size)
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = glyph
    run.font.size = Pt(size / Emu(12700) * 0.5)  # scale relative to badge size
    run.font.bold = True
    run.font.color.rgb = _hex(color)


ICONS = {
    "database": icon_database,
    "bulb": icon_bulb,
    "person": icon_person,
    "target": icon_target,
}


def _tag_chip(slide, x, y, w, h, text, fill_hex, text_hex, font, size=BODY_SIZE):
    """Small solid-fill rounded-rectangle chip behind a keyword — the
    reference deck's highlight style, used sparingly (not as a full card)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _hex(fill_hex)
    _no_line(shp)
    shp.shadow.inherit = False
    shp.adjustments[0] = 0.18
    tf = shp.text_frame
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.margin_top = tf.margin_bottom = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.name = font
    run.font.bold = True
    run.font.color.rgb = _hex(text_hex)
    return shp


def _soft_panel(slide, x, y, w, h, fill_hex):
    """Soft-tint, no-border, rounded rectangle used as a fixed sidebar panel
    — the reference deck's 'quick facts' box. Never used as a random
    decorative card; only as this one consistent, reused element."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _hex(fill_hex)
    _no_line(shp)
    shp.shadow.inherit = False
    shp.adjustments[0] = 0.06
    return shp


def _eyebrow(slide, x, y, w, text, color, font):
    """Small caps label above a title — 'GROUP CASE · 03' style."""
    return _textbox(slide, x, y, w, Inches(0.35), text.upper(), Pt(14), color,
                     font=font, bold=True)


class Deck:
    """Scratch-mode or template-mode deck builder."""

    def __init__(self, template: str | None = None, palette: str = "midnight"):
        self.template_path = template
        self.palette_name = palette
        self.pal = PALETTES[palette]
        self.header_font, self.body_font = _fonts_for(palette)
        if template:
            self.prs = Presentation(template)
            # Template mode: strip any slides that shipped inside the
            # template file itself (we only want its layouts, not its demo
            # content) unless the caller wants to edit them directly.
            self._layouts = list(self.prs.slide_masters[0].slide_layouts)
        else:
            self.prs = Presentation()
            self.prs.slide_width = SLIDE_W
            self.prs.slide_height = SLIDE_H

    # ---------------- scratch-mode layouts ----------------

    def add_title_slide(self, title: str, subtitle: str = ""):
        if self.template_path:
            return self._template_title_slide(title, subtitle)
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank
        _set_background(s, self.pal["primary"])
        box_w_in = SLIDE_W.inches - 2 * MARGIN.inches
        title_h_in = _text_block_height(title, box_w_in, TITLE_SIZE.pt, pad_in=0.15, bold=True)
        # vertically center the title+subtitle block as a unit instead of a
        # fixed y — a wrapped (two-line) title used to run into the subtitle
        block_h_in = title_h_in + (0.55 if subtitle else 0)
        top_in = max(1.2, (SLIDE_H.inches - block_h_in) / 2)
        _textbox(s, MARGIN, Inches(top_in), Inches(box_w_in), Inches(title_h_in),
                 title, TITLE_SIZE, "FFFFFF", font=self.header_font, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        if subtitle:
            _textbox(s, MARGIN, Inches(top_in + title_h_in + 0.15), Inches(box_w_in), Inches(0.8),
                      subtitle, BODY_SIZE, self.pal["accent"], font=self.body_font,
                      align=PP_ALIGN.CENTER)
        return s

    def add_stat_slide(self, stat: str, caption: str, title: str = ""):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        if title:
            _textbox(s, MARGIN, Inches(0.6), SLIDE_W - 2 * MARGIN, Inches(0.8),
                      title, SECTION_SIZE, self.pal["primary"], font=self.header_font,
                      bold=True)
        _textbox(s, MARGIN, Inches(2.3), SLIDE_W - 2 * MARGIN, Inches(2.2),
                  stat, STAT_SIZE, self.pal["primary"], font=self.header_font,
                  bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _textbox(s, Inches(2.5), Inches(4.7), SLIDE_W - Inches(5), Inches(0.9),
                  caption, BODY_SIZE, "444444", font=self.body_font,
                  align=PP_ALIGN.CENTER)
        return s

    def add_icon_rows_slide(self, title: str, rows: list[tuple[str, str, str]]):
        """rows = [(badge_label, header, description), ...] up to 4 rows.
        badge_label should be short (1-3 chars) — it's rendered inside a
        fixed-size circle, not a free-form text box."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        top = self._header(s, title) + 0.1
        row_h = Inches(1.2)
        gap = Inches(0.35)
        badge_d = Inches(0.75)
        for i, (label, header, desc) in enumerate(rows[:4]):
            y = Inches(top) + i * (row_h + gap)
            _circle_badge(s, MARGIN, y, badge_d, label,
                          self.pal["primary"], "FFFFFF")
            tx = MARGIN + badge_d + Inches(0.4)
            if header:
                _textbox(s, tx, y - Inches(0.05), SLIDE_W - tx - MARGIN, Inches(0.5),
                          header, Pt(22), self.pal["primary"], font=self.header_font,
                          bold=True)
                _textbox(s, tx, y + Inches(0.42), SLIDE_W - tx - MARGIN, Inches(0.7),
                          desc, BODY_SIZE, "444444", font=self.body_font)
            else:
                _textbox(s, tx, y, SLIDE_W - tx - MARGIN, Inches(0.75),
                          desc, BODY_SIZE, "444444", font=self.body_font,
                          anchor=MSO_ANCHOR.MIDDLE)
        return s

    def add_two_column_slide(self, title: str, bullets: list[str],
                              badge_label: str = ""):
        """Text on the left, a simple visual (badge/shape motif) on the right
        — used when no image asset is supplied, so we never ship a bare
        text-only slide."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        col_w_in = 6.6
        col_w = Inches(col_w_in)
        title_h_in = _text_block_height(title, col_w_in, TITLE_SIZE.pt, pad_in=0.15, bold=True)
        _textbox(s, MARGIN, Inches(0.6), col_w, Inches(title_h_in),
                  title, TITLE_SIZE, self.pal["primary"], font=self.header_font,
                  bold=True)
        y_in = 0.6 + title_h_in + 0.45
        bullet_w_in = col_w_in - 0.45
        for b in bullets:
            b_h_in = _text_block_height(b, bullet_w_in, BODY_SIZE.pt, pad_in=0.15)
            _circle_badge(s, MARGIN, Inches(y_in + 0.05), Inches(0.22), "",
                          self.pal["accent"], self.pal["accent"])
            _textbox(s, MARGIN + Inches(0.45), Inches(y_in), Inches(bullet_w_in), Inches(b_h_in + 0.1),
                      b, BODY_SIZE, "333333", font=self.body_font)
            y_in += b_h_in + 0.35
        # right-side visual motif: one large soft circle, on-brand, no stripe
        _circle_badge(s, Inches(9.2), Inches(2.0), Inches(3.2), badge_label,
                      self.pal["primary"], "FFFFFF")
        return s

    def add_closing_slide(self, title: str, subtitle: str = ""):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["primary"])
        box_w_in = SLIDE_W.inches - 2 * MARGIN.inches
        title_h_in = _text_block_height(title, box_w_in, TITLE_SIZE.pt, pad_in=0.15, bold=True)
        block_h_in = title_h_in + (0.55 if subtitle else 0)
        top_in = max(1.2, (SLIDE_H.inches - block_h_in) / 2)
        _textbox(s, MARGIN, Inches(top_in), Inches(box_w_in), Inches(title_h_in),
                  title, TITLE_SIZE, "FFFFFF", font=self.header_font, bold=True,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        if subtitle:
            _textbox(s, MARGIN, Inches(top_in + title_h_in + 0.15), Inches(box_w_in), Inches(0.8),
                      subtitle, BODY_SIZE, self.pal["accent"], font=self.body_font,
                      align=PP_ALIGN.CENTER)
        return s

    def add_process_flow_slide(self, title: str, steps: list[tuple[str, str, str]]):
        """steps = [(icon_name, label, sublabel), ...], up to 4, connected by
        arrows — the reference deck's 'Data -> Analysis -> Decision' layout.
        icon_name is one of ICONS (database/bulb/person/target)."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        _textbox(s, MARGIN, Inches(0.55), SLIDE_W - 2 * MARGIN, Inches(0.9),
                  title, TITLE_SIZE, self.pal["primary"], font=self.header_font, bold=True)
        n = len(steps[:4])
        icon_size = Inches(1.1)
        total_w = SLIDE_W - 2 * MARGIN
        slot_w = total_w / n
        y = Inches(2.4)
        for i, (icon_name, label, sublabel) in enumerate(steps[:4]):
            cx = MARGIN + i * slot_w + (slot_w - icon_size) / 2
            fn = ICONS.get(icon_name, icon_target)
            fn(s, cx, y, icon_size, self.pal["primary"])
            _textbox(s, MARGIN + i * slot_w, y + icon_size + Inches(0.25), slot_w, Inches(0.5),
                      label, Pt(20), self.pal["primary"], font=self.header_font,
                      bold=True, align=PP_ALIGN.CENTER)
            _textbox(s, MARGIN + i * slot_w, y + icon_size + Inches(0.75), slot_w, Inches(0.8),
                      sublabel, CAPTION_SIZE, TEXT_DARK, font=self.body_font,
                      align=PP_ALIGN.CENTER)
            if i < n - 1:
                icon_arrow_right(s, MARGIN + (i + 1) * slot_w - Inches(0.35),
                                  y + icon_size / 2 - Inches(0.14), Inches(0.7), Inches(0.28),
                                  self.pal["accent"])
        return s

    def add_facts_slide(self, title: str, paragraph: str, bullets: list[str],
                          facts: list[tuple[str, str]], eyebrow: str = ""):
        """Left: eyebrow + title + a short paragraph + bullets.
        Right: a fixed soft-tint 'quick facts' panel (label/value pairs) —
        the reference deck's sidebar, not a decorative card.

        Every element's vertical position is derived from the actual
        estimated height of the element above it, so long paragraphs or
        wrapped fact values push later content down instead of overlapping
        it (the box height passed to _textbox is generous headroom, not a
        clip boundary — PowerPoint text boxes overflow past their bottom
        edge if you under-size them)."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        left_w_in = 7.6
        left_w = Inches(left_w_in)
        y_in = 0.55
        eyebrow_color = self.pal["primary"] if self.palette_name == "campus_blue" else self.pal["accent"]
        if eyebrow:
            _eyebrow(s, MARGIN, Inches(y_in), left_w, eyebrow, eyebrow_color, self.body_font)
            y_in += 0.45
        title_h = _text_block_height(title, left_w_in, 34, pad_in=0.15, bold=True)
        _textbox(s, MARGIN, Inches(y_in), left_w, Inches(title_h + 0.15), title, Pt(34),
                  self.pal["primary"], font=self.header_font, bold=True)
        y_in += title_h + 0.35
        para_pt = BODY_SIZE / Pt(1)
        para_h = _text_block_height(paragraph, left_w_in, para_pt, pad_in=0.2)
        _textbox(s, MARGIN, Inches(y_in), left_w, Inches(para_h + 0.2), paragraph,
                  BODY_SIZE, TEXT_DARK, font=self.body_font)
        y_in += para_h + 0.45
        for b in bullets:
            b_h = _text_block_height(b, left_w_in - 0.4, para_pt, pad_in=0.15)
            _circle_badge(s, MARGIN, Inches(y_in + 0.08), Inches(0.16), "",
                          self.pal["primary"], self.pal["primary"])
            _textbox(s, MARGIN + Inches(0.4), Inches(y_in), left_w - Inches(0.4),
                      Inches(b_h + 0.15), b, BODY_SIZE, TEXT_DARK, font=self.body_font)
            y_in += b_h + 0.35

        panel_x_in = 0.6 + left_w_in + 0.4
        panel_x = Inches(panel_x_in)
        panel_w_in = 13.333 - panel_x_in - 0.6
        panel_w = Inches(panel_w_in)
        panel_y_in = 0.9
        val_w_in = panel_w_in - 0.6
        row_heights = []
        for label, value in facts:
            v_h = _text_block_height(value, val_w_in, para_pt, pad_in=0.1)
            row_heights.append(0.35 + v_h + 0.25)  # label row + value block + row gap
        panel_h_in = min(0.45 + sum(row_heights), 6.6)
        _soft_panel(s, panel_x, Inches(panel_y_in), panel_w, Inches(panel_h_in), self.pal["accent"])
        fy_in = panel_y_in + 0.3
        for (label, value), rh in zip(facts, row_heights):
            _textbox(s, panel_x + Inches(0.3), Inches(fy_in), panel_w - Inches(0.6), Inches(0.32),
                      label.upper(), Pt(14), self.pal["primary"], font=self.body_font, bold=True)
            v_h = rh - 0.35 - 0.25
            _textbox(s, panel_x + Inches(0.3), Inches(fy_in + 0.35), panel_w - Inches(0.6),
                      Inches(v_h + 0.15), value, BODY_SIZE, TEXT_DARK, font=self.body_font)
            fy_in += rh
        return s

    def add_tagged_list_slide(self, title: str, icon_name: str,
                                items: list[tuple[str, bool]]):
        """items = [(text, tagged), ...]. Tagged items get a solid chip
        (the reference deck's highlight style); the rest sit as plain text.
        Use tags sparingly — 2-3 per slide, never every line."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        _textbox(s, MARGIN, Inches(0.55), SLIDE_W - 2 * MARGIN, Inches(0.9),
                  title, TITLE_SIZE, self.pal["primary"], font=self.header_font, bold=True)
        fn = ICONS.get(icon_name, icon_target)
        fn(s, Inches(9.6), Inches(2.0), Inches(2.6), self.pal["primary"])
        y_in = 2.1
        col_w_in = 4.6
        col_w = Inches(col_w_in)
        for text, tagged in items:
            if tagged:
                w = min(Inches(0.55) + Inches(0.14) * len(text), col_w)
                _tag_chip(s, MARGIN, Inches(y_in), w, Inches(0.55), text,
                          self.pal["primary"], "FFFFFF", self.body_font, size=Pt(18))
                y_in += 0.55 + 0.3
            else:
                t_h_in = _text_block_height(text, col_w_in, BODY_SIZE.pt, pad_in=0.15)
                _textbox(s, MARGIN, Inches(y_in + 0.1), col_w, Inches(t_h_in + 0.1),
                          text, BODY_SIZE, TEXT_DARK, font=self.body_font)
                y_in += t_h_in + 0.1 + 0.3
        return s

    def add_stat_list_slide(self, title: str, eyebrow: str,
                              items: list[tuple[str, str]], note_heading: str = "",
                              note_text: str = "", footer: str = "", page_num=None):
        """A vertical list of (stat, description) rows — the stat as plain
        bold colored text, NOT squeezed into a circle badge (badges are for
        short 1-2 character labels only; a stat like '5,450 -> 5,200' needs
        room to breathe)."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        stat_w = 2.3
        desc_x = 0.6 + stat_w + 0.35
        desc_w = 13.333 - desc_x - 0.6
        for stat, desc in items:
            s_h = _text_block_height(stat, stat_w, 22, pad_in=0.05)
            d_h = _text_block_height(desc, desc_w, 16, pad_in=0.05)
            row_h = max(s_h, d_h)
            _textbox(s, MARGIN, Inches(y), Inches(stat_w), Inches(row_h + 0.1),
                      stat, Pt(22), self.pal["primary"], font=self.header_font, bold=True)
            _textbox(s, Inches(desc_x), Inches(y + 0.05), Inches(desc_w), Inches(row_h + 0.1),
                      desc, Pt(16), TEXT_DARK, font=self.body_font)
            y += row_h + 0.35
        if note_heading:
            _textbox(s, MARGIN, Inches(y), Inches(12.1), Inches(0.35), note_heading,
                      Pt(15), self.pal["primary"], font=self.header_font, bold=True)
            y += 0.38
        if note_text:
            _textbox(s, MARGIN, Inches(y), Inches(12.1), Inches(0.35), note_text,
                      Pt(13), "666666", font=self.body_font, italic=True)
        self._footer(s, footer, page_num)
        return s

    # ---------------- header/footer/image helpers ----------------

    def _header(self, slide, title: str, eyebrow: str = "", title_pt: float = 34,
                width_in: float = None, y_start_in: float = 0.5):
        """Shared eyebrow+title block used by the composite report-style
        layouts below. Title height is estimated dynamically so a long,
        wrapping title never overlaps the content beneath it."""
        width_in = width_in if width_in is not None else (13.333 - 1.2)
        y = y_start_in
        if eyebrow:
            ec = self.pal["primary"] if self.palette_name == "campus_blue" else self.pal["accent"]
            _eyebrow(slide, MARGIN, Inches(y), Inches(width_in), eyebrow, ec, self.body_font)
            y += 0.38
        title_h = _text_block_height(title, width_in, title_pt, pad_in=0.1)
        _textbox(slide, MARGIN, Inches(y), Inches(width_in), Inches(title_h + 0.1),
                  title, Pt(title_pt), self.pal["primary"], font=self.header_font, bold=True)
        y += title_h + 0.3
        return y

    def _footer(self, slide, text: str = "", page_num=None):
        if text:
            _textbox(slide, MARGIN, Inches(7.08), Inches(10), Inches(0.3),
                      text, Pt(11), "999999", font=self.body_font)
        if page_num is not None:
            _textbox(slide, Inches(12.3), Inches(7.08), Inches(0.45), Inches(0.3),
                      str(page_num), Pt(11), "999999", font=self.body_font, align=PP_ALIGN.RIGHT)

    def _add_image(self, slide, path, x_in, y_in, w_in=None, h_in=None):
        """Inserts the image unmodified, scaling by width OR height only so
        the original aspect ratio (and therefore the figure itself) is never
        distorted."""
        kwargs = {}
        if w_in:
            kwargs["width"] = Inches(w_in)
        if h_in:
            kwargs["height"] = Inches(h_in)
        return slide.shapes.add_picture(path, Inches(x_in), Inches(y_in), **kwargs)

    # ---------------- composite report layouts ----------------
    # Built for dense, fact-heavy report/academic decks: preserve every
    # number and every figure exactly, just lay them out cleanly. These
    # still hold the line on 24pt body text, <=3 colors, and no decorative
    # boxes (soft-tint panels/tag chips excepted, as elsewhere in this file).

    def add_title_roster_slide(self, title: str, subtitle: str, eyebrow: str,
                                 roster: list[tuple[str, str]], footer_note: str = ""):
        """Dark title slide + a team roster list + a small dataset/footer
        note — the report cover-page pattern."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["primary"])
        y = 0.7
        if eyebrow:
            _eyebrow(s, MARGIN, Inches(y), Inches(12.1), eyebrow, self.pal["accent"], self.body_font)
            y += 0.5
        _textbox(s, MARGIN, Inches(y), Inches(11), Inches(1.0), title, Pt(40),
                  "FFFFFF", font=self.header_font, bold=True)
        y += 1.15
        sub_h = _text_block_height(subtitle, 9.5, 22, pad_in=0.1)
        _textbox(s, MARGIN, Inches(y), Inches(9.5), Inches(sub_h + 0.1), subtitle,
                  Pt(22), self.pal["accent"], font=self.body_font)
        y += sub_h + 0.55
        _eyebrow(s, MARGIN, Inches(y), Inches(10), "TEAM", self.pal["accent"], self.body_font)
        y += 0.4
        col_w = 12.1 / max(1, min(len(roster), 5))
        for i, (name, ident) in enumerate(roster):
            cx = MARGIN + Inches(i * col_w)
            _textbox(s, cx, Inches(y), Inches(col_w - 0.15), Inches(0.4),
                      name, Pt(18), "FFFFFF", font=self.body_font, bold=True)
            _textbox(s, cx, Inches(y + 0.38), Inches(col_w - 0.15), Inches(0.35),
                      ident, Pt(14), self.pal["accent"], font=self.body_font)
        if footer_note:
            _textbox(s, MARGIN, Inches(6.85), Inches(12.1), Inches(0.4),
                      footer_note, Pt(13), self.pal["accent"], font=self.body_font)
        return s

    def add_two_image_slide(self, title: str, eyebrow: str, intro: str,
                              images: list[tuple[str, str]], footer: str = "", page_num=None):
        """Two figures side by side, each with a short caption — for
        before/after or method-comparison charts (scree/cumulative-variance,
        elbow/silhouette, etc.)."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        if intro:
            intro_h = _text_block_height(intro, 12.1, 20, pad_in=0.1)
            _textbox(s, MARGIN, Inches(y), Inches(12.1), Inches(intro_h + 0.1),
                      intro, Pt(20), TEXT_DARK, font=self.body_font)
            y += intro_h + 0.3
        CONTENT_BOTTOM = 6.75  # leaves clear room before the footer at 7.08in
        gap = 0.4
        col_w = (12.1 - gap) / 2
        cap_h = 0.32 if any(c for _, c in images[:2]) else 0.0
        avail_h = max(1.4, CONTENT_BOTTOM - y - cap_h - 0.1)
        from PIL import Image as _PILImage
        for i, (path, caption) in enumerate(images[:2]):
            iw, ih = _PILImage.open(path).size
            aspect = iw / ih
            w = col_w
            h = w / aspect
            if h > avail_h:
                h = avail_h
                w = h * aspect
            x_in = 0.6 + i * (col_w + gap) + (col_w - w) / 2
            self._add_image(s, path, x_in, y, w_in=w)
            if caption:
                _textbox(s, Inches(0.6 + i * (col_w + gap)), Inches(y + h + 0.08), Inches(col_w),
                          Inches(0.3), caption, CAPTION_SIZE, "666666", font=self.body_font,
                          align=PP_ALIGN.CENTER)
        self._footer(s, footer, page_num)
        return s

    def add_image_stats_slide(self, title: str, eyebrow: str, image_path: str,
                                image_caption: str, stats: list[tuple[str, str]],
                                bullets: list[str], footer: str = "", page_num=None,
                                image_w_in: float = 4.2):
        """Figure on the left; a compact single-row stat strip + observation
        bullets on the right — for an EDA/overview slide anchored by one
        chart plus headline numbers. image_w_in is tunable: a near-square
        figure (e.g. a correlation matrix) needs to run narrower than a
        wide chart so the text column keeps enough width to avoid deep
        wrapping when there are several long observation bullets."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        from PIL import Image as _PILImage
        iw, ih = _PILImage.open(image_path).size
        aspect = iw / ih
        img_w = image_w_in
        img_h = img_w / aspect
        self._add_image(s, image_path, 0.6, y, w_in=img_w)
        if image_caption:
            _textbox(s, Inches(0.6), Inches(y + img_h + 0.08), Inches(img_w), Inches(0.35),
                      image_caption, CAPTION_SIZE, "666666", font=self.body_font)
        right_x = 0.6 + img_w + 0.4
        right_w = 13.333 - right_x - 0.6
        ry = y
        n_stats = max(1, len(stats[:4]))
        stat_col_w = (right_w - 0.25 * (n_stats - 1)) / n_stats
        for i, (val, lbl) in enumerate(stats[:4]):
            sx = right_x + i * (stat_col_w + 0.25)
            _textbox(s, Inches(sx), Inches(ry), Inches(stat_col_w), Inches(0.42),
                      val, Pt(22), self.pal["primary"], font=self.header_font, bold=True)
            lbl_h = _text_block_height(lbl, stat_col_w, 12, pad_in=0.03)
            _textbox(s, Inches(sx), Inches(ry + 0.44), Inches(stat_col_w), Inches(lbl_h + 0.1),
                      lbl, Pt(12), TEXT_DARK, font=self.body_font)
        by = ry + 0.44 + max(_text_block_height(l, stat_col_w, 12, pad_in=0.03) for _, l in stats[:4]) + 0.3
        if bullets:
            _eyebrow(s, Inches(right_x), Inches(by), Inches(right_w), "OBSERVATIONS",
                      self.pal["accent"], self.body_font)
            by += 0.32
            for b in bullets:
                b_h = _text_block_height(b, right_w - 0.32, 14, pad_in=0.06)
                _circle_badge(s, Inches(right_x), Inches(by + 0.05), Inches(0.1), "",
                              self.pal["primary"], self.pal["primary"])
                _textbox(s, Inches(right_x + 0.24), Inches(by), Inches(right_w - 0.24),
                          Inches(b_h + 0.1), b, Pt(14), TEXT_DARK, font=self.body_font)
                by += b_h + 0.14
        self._footer(s, footer, page_num)
        return s

    def add_bullets_image_slide(self, title: str, eyebrow: str, intro: str,
                                  bullets: list[str], image_path: str, image_caption: str,
                                  footer: str = "", page_num=None, left_w: float = 7.2):
        """Argument/explanation as bulleted text on the left, one supporting
        figure on the right — for 'here's the evidence' slides. left_w is
        tunable: dense, multi-sentence bullets need a wide text column to
        avoid deep wrapping, even if that means a smaller figure. Font size
        adapts down (15pt -> 12pt) if the bullets don't fit the vertical
        budget at the default size."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        avail = 6.55 - y

        def _layout_height(pt, gap):
            h = 0.0
            if intro:
                h += _text_block_height(intro, left_w, pt, pad_in=0.06) + gap
            for b in bullets:
                h += _text_block_height(b, left_w - 0.32, pt, pad_in=0.05) + gap
            return h

        pt, gap = 15, 0.16
        for cand_pt, cand_gap in ((15, 0.16), (14, 0.14), (13, 0.12), (12, 0.1)):
            if _layout_height(cand_pt, cand_gap) <= avail:
                pt, gap = cand_pt, cand_gap
                break
        else:
            pt, gap = 12, 0.1  # smallest allowed; may still be tight, but never below this floor

        ty = y
        if intro:
            i_h = _text_block_height(intro, left_w, pt, pad_in=0.06)
            _textbox(s, MARGIN, Inches(ty), Inches(left_w), Inches(i_h + 0.1),
                      intro, Pt(pt), TEXT_DARK, font=self.body_font, italic=True)
            ty += i_h + gap
        for b in bullets:
            b_h = _text_block_height(b, left_w - 0.32, pt, pad_in=0.05)
            _circle_badge(s, MARGIN, Inches(ty + 0.05), Inches(0.11), "",
                          self.pal["primary"], self.pal["primary"])
            _textbox(s, MARGIN + Inches(0.26), Inches(ty), Inches(left_w - 0.26),
                      Inches(b_h + 0.1), b, Pt(pt), TEXT_DARK, font=self.body_font)
            ty += b_h + gap
        from PIL import Image as _PILImage
        iw, ih = _PILImage.open(image_path).size
        aspect = iw / ih
        img_x = 0.6 + left_w + 0.35
        img_w = 13.333 - img_x - 0.6
        img_h = img_w / aspect
        max_h = 6.55 - y
        if img_h > max_h:
            img_h = max_h
            img_w = img_h * aspect
        self._add_image(s, image_path, img_x, y, w_in=img_w)
        if image_caption:
            _textbox(s, Inches(img_x), Inches(y + img_h + 0.08), Inches(img_w), Inches(0.3),
                      image_caption, CAPTION_SIZE, "666666", font=self.body_font,
                      align=PP_ALIGN.CENTER)
        self._footer(s, footer, page_num)
        return s

    def add_stat_image_slide(self, title: str, eyebrow: str, intro: str, rule_text: str,
                               stat_value: str, stat_caption: str, note: str,
                               image_path: str, image_caption: str,
                               footer: str = "", page_num=None):
        """Headline stat + supporting rule/note on the left, one figure on
        the right — for 'here's the number, here's the evidence' slides."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        left_w = 5.6
        ty = y
        if intro:
            i_h = _text_block_height(intro, left_w, 15, pad_in=0.1)
            _textbox(s, MARGIN, Inches(ty), Inches(left_w), Inches(i_h + 0.1),
                      intro, Pt(15), TEXT_DARK, font=self.body_font)
            ty += i_h + 0.2
        if rule_text:
            r_h = _text_block_height(rule_text, left_w - 0.4, 14, pad_in=0.15)
            _soft_panel(s, Inches(0.6), Inches(ty), Inches(left_w), Inches(r_h + 0.3),
                        self.pal["accent"])
            _textbox(s, Inches(0.6 + 0.2), Inches(ty + 0.15), Inches(left_w - 0.4),
                      Inches(r_h + 0.1), rule_text, Pt(14), self.pal["primary"], font=self.body_font, bold=True)
            ty += r_h + 0.5
        _textbox(s, MARGIN, Inches(ty), Inches(left_w), Inches(1.0),
                  stat_value, Pt(64), self.pal["primary"], font=self.header_font, bold=True)
        ty += 1.0
        cap_h = _text_block_height(stat_caption, left_w, 18, pad_in=0.1)
        _textbox(s, MARGIN, Inches(ty), Inches(left_w), Inches(cap_h + 0.1),
                  stat_caption, Pt(18), TEXT_DARK, font=self.body_font)
        ty += cap_h + 0.25
        if note:
            n_h = _text_block_height(note, left_w, 15, pad_in=0.1)
            _textbox(s, MARGIN, Inches(ty), Inches(left_w), Inches(n_h + 0.1),
                      note, Pt(15), "666666", font=self.body_font, italic=True)
        from PIL import Image as _PILImage
        iw, ih = _PILImage.open(image_path).size
        aspect = iw / ih
        img_x = 0.6 + left_w + 0.5
        img_w = 13.333 - img_x - 0.6
        img_h = img_w / aspect
        max_h = 6.55 - y
        if img_h > max_h:
            img_h = max_h
            img_w = img_h * aspect
        img_y = y + (max_h - img_h) / 2
        self._add_image(s, image_path, img_x, img_y, w_in=img_w)
        if image_caption:
            _textbox(s, Inches(img_x), Inches(img_y + img_h + 0.08), Inches(img_w), Inches(0.5),
                      image_caption, CAPTION_SIZE, "666666", font=self.body_font,
                      align=PP_ALIGN.CENTER)
        self._footer(s, footer, page_num)
        return s

    def add_grid_cards_slide(self, title: str, eyebrow: str,
                               cards: list[dict], cols: int = 2,
                               footer: str = "", page_num=None):
        """Up to 4 cards in a grid, each: badge, heading, optional stat line,
        and a few detail lines. No card borders/fills — separation comes
        from the badge + consistent spacing, per the no-decorative-boxes
        rule. cards: [{'badge': '1', 'heading': str, 'stat': str,
        'lines': [str, ...]}, ...]"""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["background"])
        y = self._header(s, title, eyebrow)
        n = len(cards)
        rows = (n + cols - 1) // cols
        gap_x, gap_y = 0.5, 0.35
        cell_w = (12.1 - gap_x * (cols - 1)) / cols
        cell_h = (6.7 - y - gap_y * (rows - 1)) / rows
        for i, card in enumerate(cards):
            r, c = divmod(i, cols)
            cx = MARGIN + c * Inches(cell_w + gap_x)
            cy = Inches(y) + r * Inches(cell_h + gap_y)
            badge_d = Inches(0.5)
            _circle_badge(s, cx, cy, badge_d, card.get("badge", ""), self.pal["primary"], "FFFFFF")
            tx = cx + badge_d + Inches(0.25)
            tw = Inches(cell_w) - badge_d - Inches(0.25)
            tw_in = cell_w - (badge_d / Emu(914400)) - 0.25
            head_h = _text_block_height(card["heading"], tw_in, 18, pad_in=0.05)
            _textbox(s, tx, cy - Inches(0.03), tw, Inches(head_h + 0.1),
                      card["heading"], Pt(18), self.pal["primary"], font=self.header_font, bold=True)
            ny = cy + Inches(max(head_h, 0.4) + 0.05)
            if card.get("stat"):
                _textbox(s, tx, ny, tw, Inches(0.35), card["stat"], Pt(15),
                          self.pal["accent"] if self.palette_name != "campus_blue" else self.pal["primary"],
                          font=self.body_font, bold=True)
                ny += Inches(0.38)
            for line in card.get("lines", []):
                l_h = _text_block_height(line, cell_w - 0.05, 14, pad_in=0.06)
                _textbox(s, cx, ny, Inches(cell_w), Inches(l_h + 0.1), line, Pt(14),
                          TEXT_DARK, font=self.body_font)
                ny += Inches(l_h + 0.14)
        self._footer(s, footer, page_num)
        return s

    def add_closing_bullets_slide(self, title: str, eyebrow: str, bullets: list[str],
                                    summary_label: str, summary_text: str,
                                    footer: str = "", page_num=None):
        """Dark closing/limitations slide: a bulleted list plus one
        highlighted 'bottom line' paragraph. Bullet/summary font sizes
        adapt down a tier if needed so nothing runs past the slide edge."""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        _set_background(s, self.pal["primary"])
        y = 0.6
        if eyebrow:
            _eyebrow(s, MARGIN, Inches(y), Inches(12.1), eyebrow, self.pal["accent"], self.body_font)
            y += 0.4
        title_h = _text_block_height(title, 12.1, 32, pad_in=0.1)
        _textbox(s, MARGIN, Inches(y), Inches(12.1), Inches(title_h + 0.1), title, Pt(32),
                  "FFFFFF", font=self.header_font, bold=True)
        y += title_h + 0.35
        avail = 6.75 - y - (0.35 if summary_label else 0) - 0.15

        def _fit_height(b_pt, b_gap, s_pt):
            h = 0.0
            for b in bullets:
                h += _text_block_height(b, 11.6, b_pt, pad_in=0.06) + b_gap
            if summary_text:
                h += _text_block_height(summary_text, 12.0, s_pt, pad_in=0.08)
            return h

        tiers = [(17, 0.22, 18), (16, 0.18, 17), (15, 0.15, 16), (14, 0.12, 15)]
        b_pt, b_gap, s_pt = tiers[-1]
        for cand in tiers:
            if _fit_height(*cand) <= avail:
                b_pt, b_gap, s_pt = cand
                break

        for b in bullets:
            b_h = _text_block_height(b, 11.6, b_pt, pad_in=0.06)
            _circle_badge(s, MARGIN, Inches(y + 0.06), Inches(0.13), "",
                          self.pal["accent"], self.pal["accent"])
            _textbox(s, MARGIN + Inches(0.3), Inches(y), Inches(11.6), Inches(b_h + 0.1),
                      b, Pt(b_pt), "E8ECF4", font=self.body_font)
            y += b_h + b_gap
        y += 0.12
        if summary_label:
            _eyebrow(s, MARGIN, Inches(y), Inches(12.1), summary_label, self.pal["accent"], self.body_font)
            y += 0.32
        if summary_text:
            s_h = _text_block_height(summary_text, 12.0, s_pt, pad_in=0.08)
            _textbox(s, MARGIN, Inches(y), Inches(12.0), Inches(s_h + 0.1),
                      summary_text, Pt(s_pt), "FFFFFF", font=self.body_font, bold=True)
        self._footer_light(s, footer, page_num)
        return s

    def _footer_light(self, slide, text: str = "", page_num=None):
        if text:
            _textbox(slide, MARGIN, Inches(7.08), Inches(10), Inches(0.3),
                      text, Pt(11), self.pal["accent"], font=self.body_font)
        if page_num is not None:
            _textbox(slide, Inches(12.3), Inches(7.08), Inches(0.45), Inches(0.3),
                      str(page_num), Pt(11), self.pal["accent"], font=self.body_font, align=PP_ALIGN.RIGHT)

    # ---------------- template-mode helpers ----------------

    def _pick_layout(self, kind: str):
        """Very small heuristic matcher: look at placeholder types on each
        layout in the template and pick the closest fit. Falls back to the
        layout at the conventional PowerPoint index for that kind."""
        want = {
            "title": 0,      # "Title Slide"
            "content": 1,    # "Title and Content"
            "section": 2,    # "Section Header"
            "closing": 0,
        }
        idx = want.get(kind, 1)
        if idx < len(self._layouts):
            return self._layouts[idx]
        return self._layouts[0]

    def _fill_placeholder(self, slide, idx_type, text):
        for ph in slide.placeholders:
            if ph.placeholder_format.type == idx_type or ph.placeholder_format.idx == idx_type:
                tf = ph.text_frame
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = text
                else:
                    tf.paragraphs[0].text = text
                # enforce the body-size floor even inside template placeholders
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.size is None or r.font.size < BODY_SIZE:
                            if ph.placeholder_format.idx != 0:  # don't shrink titles
                                r.font.size = BODY_SIZE
                return True
        return False

    def _template_title_slide(self, title, subtitle):
        layout = self._pick_layout("title")
        s = self.prs.slides.add_slide(layout)
        for ph in s.placeholders:
            if ph.placeholder_format.idx == 0:
                ph.text_frame.paragraphs[0].text = title
            elif ph.placeholder_format.idx == 1:
                ph.text_frame.paragraphs[0].text = subtitle
        return s

    def add_template_content_slide(self, title: str, bullets: list[str]):
        """Fill the template's own 'Title and Content' layout with a title
        + bullet list, keeping the template's native fonts/colors."""
        layout = self._pick_layout("content")
        s = self.prs.slides.add_slide(layout)
        title_ph, body_ph = None, None
        for ph in s.placeholders:
            if ph.placeholder_format.idx == 0:
                title_ph = ph
            elif ph.placeholder_format.type is not None and ph.has_text_frame and ph.placeholder_format.idx != 0:
                body_ph = body_ph or ph
        if title_ph:
            title_ph.text_frame.paragraphs[0].text = title
        if body_ph:
            tf = body_ph.text_frame
            tf.clear()
            for i, b in enumerate(bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = b
                for r in p.runs or [p.add_run()]:
                    if r.font.size is None or r.font.size < BODY_SIZE:
                        r.font.size = BODY_SIZE
        return s

    # ---------------- output ----------------

    def save(self, path: str):
        self.prs.save(path)
        return path


if __name__ == "__main__":
    # Self-test: three one-slide demos, scratch mode, no template.
    import sys

    d1 = Deck(palette="midnight")
    d1.add_title_slide("Humanized Deck Demo", "Title layout — dark, centered, no accent line")
    d1.save("/home/claude/demo_1_title.pptx")

    d2 = Deck(palette="teal")
    d2.add_icon_rows_slide("Icon-row layout", [
        ("1", "One idea per row", "Numbered badge, bold header, 24pt description — scannable in 3-5 seconds."),
        ("2", "No boxes", "Content sits on open whitespace instead of a bordered card."),
        ("3", "One accent color", "Badges reuse the same accent everywhere for consistency."),
    ])
    d2.save("/home/claude/demo_2_iconrows.pptx")

    d3 = Deck(palette="terracotta")
    d3.add_two_column_slide("Two-column layout", [
        "Left column carries the argument in short, left-aligned lines.",
        "Right side carries a single visual motif instead of a stock photo.",
        "Same margins and type scale as every other slide in the deck.",
    ], badge_label="")
    d3.save("/home/claude/demo_3_twocolumn.pptx")

    print("Built 3 demo decks.")
