# -*- coding: utf-8 -*-
"""Welcome Kit - Induction Project (4 slides).
Perfil: tecnico de soporte IT independiente, base en David, Chiriqui.
Alineado con la rubrica 'Final Project Induction (Instructions + Rubric)'."""

import os
import re
import shutil

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

# ---------------------------------------------------------------- paleta
NAVY = RGBColor(0x0E, 0x27, 0x3F)
NAVY_SOFT = RGBColor(0x17, 0x3A, 0x59)
BLUE = RGBColor(0x1F, 0x6F, 0x9E)
ACCENT = RGBColor(0xF2, 0xA6, 0x3B)
BG = RGBColor(0xF5, 0xF7, 0xFA)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xDD, 0xE5, 0xEC)
TEXT = RGBColor(0x1B, 0x2B, 0x3A)
MUTED = RGBColor(0x55, 0x69, 0x7C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

MAP_BG = RGBColor(0xEB, 0xF0, 0xE8)
ROAD = RGBColor(0xDC, 0xE2, 0xE7)
BLOCK = RGBColor(0xD2, 0xDC, 0xD3)
PARK = RGBColor(0xC7, 0xDF, 0xBB)

FONT = "Segoe UI"
W, H = 13.333, 7.5
M = 0.85

WORDS = []


def _count(text):
    WORDS.extend(re.findall(r"[A-Za-z']+", text))


def tx(slide, text, l, t, w, h, size=18, bold=False, color=TEXT, font=FONT,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0, space_after=0,
       italic=False, caps=False, count=True):
    if count:
        _count(text)
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text.upper() if caps else text
    f = r.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    return box


def card(slide, l, t, w, h, fill=CARD, radius=0.035, line=LINE):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t),
                                 Inches(w), Inches(h))
    shp.adjustments[0] = radius
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    shp.text_frame.word_wrap = True
    return shp


def rect(slide, l, t, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t),
                                 Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def circle(slide, l, t, d, fill, line=None, lw=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(d), Inches(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(lw)
    shp.shadow.inherit = False
    return shp


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def header(slide, kicker, title):
    tx(slide, kicker, M, 0.55, 8.0, 0.3, size=11, bold=True, color=ACCENT, caps=True)
    tx(slide, title, M, 0.88, 11.4, 0.75, size=36, bold=True, color=NAVY)
    rect(slide, M + 0.02, 1.72, 0.9, 0.055, ACCENT)


def footer(slide, n):
    tx(slide, "Induction Project  \u00b7  Davis Fern\u00e1ndez", M, H - 0.55, 6.0, 0.3,
       size=9.5, color=MUTED, count=False)
    tx(slide, f"{n:02d}", W - M - 0.5, H - 0.55, 0.5, 0.3, size=9.5, color=MUTED,
       align=PP_ALIGN.RIGHT, count=False)


# --------------------------------------------------------------- mapa simple
def draw_map(slide, L, T, MW, MH):
    """Esquema del mapa: base en La Riviera (C. 5) y cliente PikTa (C. Aristides Romero)."""
    import math

    def X(u):
        return L + u * MW

    def Y(v):
        return T + v * MH

    def mrect(u, v, w, h, fill, line=None):
        return rect(slide, X(u), Y(v), w * MW, h * MH, fill, line)

    def road(u1, v1, u2, v2, thick, fill=ROAD):
        x1, y1, x2, y2 = X(u1), Y(v1), X(u2), Y(v2)
        length = math.hypot(x2 - x1, y2 - y1)
        ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
        shp = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches((x1 + x2) / 2 - length / 2),
            Inches((y1 + y2) / 2 - thick / 2),
            Inches(length), Inches(thick))
        shp.rotation = ang
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
        shp.line.fill.background()
        shp.shadow.inherit = False
        return shp

    def label(text, u, v, size=8, color=MUTED, bold=False, w=0.0, caps=False):
        if w == 0:
            w = 0.30
        tx(slide, text, X(u - w / 2), Y(v), w * MW, 0.2, size=size, color=color,
           bold=bold, align=PP_ALIGN.CENTER, count=False, caps=caps)

    card(slide, L, T, MW, MH, fill=MAP_BG, radius=0.03, line=RGBColor(0xD8, 0xE2, 0xD5))

    # --- manzanas
    mrect(0.02, 0.74, 0.30, 0.14, PARK)          # La Riviera
    mrect(0.36, 0.55, 0.16, 0.13, BLOCK)
    mrect(0.30, 0.90, 0.18, 0.08, BLOCK)
    mrect(0.04, 0.10, 0.14, 0.20, BLOCK)
    mrect(0.22, 0.06, 0.16, 0.15, BLOCK)
    mrect(0.52, 0.05, 0.16, 0.14, BLOCK)
    mrect(0.82, 0.30, 0.14, 0.14, BLOCK)
    mrect(0.86, 0.62, 0.12, 0.12, BLOCK)
    mrect(0.60, 0.62, 0.13, 0.12, BLOCK)

    # --- calles
    road(0.0, 0.40, 1.0, 0.14, 0.115)            # C. Aristides Romero
    road(0.0, 0.92, 0.86, 0.78, 0.095)           # C. 5
    road(0.552, 0.829, 0.615, 0.240, 0.075)      # C. 4

    # --- ruta
    pts = [(0.140, 0.897), (0.552, 0.829), (0.615, 0.240), (0.775, 0.1985)]
    for (u1, v1), (u2, v2) in zip(pts, pts[1:]):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(X(u1)), Inches(Y(v1)),
                                       Inches(X(u2)), Inches(Y(v2)))
        c.line.color.rgb = ACCENT
        c.line.width = Pt(4.0)

    # --- base (La Riviera)
    circle(slide, X(0.140) - 0.075, Y(0.897) - 0.075, 0.15, NAVY, WHITE, 1.6)
    label("MY BASE", 0.140, 0.925, w=0.26, bold=True, color=NAVY)
    label("LA RIVERA", 0.140, 0.775, w=0.30, size=7, bold=True,
          color=RGBColor(0x3B, 0x63, 0x2F), caps=True)

    # --- PikTa
    pk = circle(slide, X(0.775) - 0.075, Y(0.1985) - 0.075, 0.15, ACCENT, WHITE, 1.6)
    label("PIKTA", 0.775, 0.228, w=0.20, bold=True, color=RGBColor(0x8E, 0x57, 0x06),
          caps=True)

    # --- Ozaga Market
    mrect(0.700, 0.300, 0.085, 0.075, RGBColor(0xCC, 0xD7, 0xE1))
    label("OZAGA MARKET", 0.742, 0.385, w=0.26, color=RGBColor(0x2C, 0x40, 0x52))

    # --- nombres de calles
    label("C. ARISTIDES ROMERO", 0.300, 0.400, w=0.40, size=7.5,
          color=RGBColor(0x5C, 0x6E, 0x7F), bold=True, caps=True)
    label("C. 5", 0.500, 0.770, w=0.12, size=7.5, color=RGBColor(0x5C, 0x6E, 0x7F),
          bold=True)
    label("C. 4", 0.662, 0.550, w=0.12, size=7.5, color=RGBColor(0x5C, 0x6E, 0x7F),
          bold=True)

    # --- brujula
    circle(slide, X(0.930), Y(0.850), 0.30, WHITE, LINE, 0.75)
    tx(slide, "N", X(0.930), Y(0.850) + 0.045, 0.30, 0.2, size=9, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER, count=False)


prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]

# ================================================================ SLIDE 1
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, NAVY)
circle(s, W - 3.55, -1.55, 5.6, NAVY_SOFT)
circle(s, W - 1.85, H - 1.75, 3.2, ACCENT)
circle(s, -1.15, H - 1.55, 2.9, NAVY_SOFT)
rect(s, 0, H - 0.11, W, 0.11, ACCENT)

tx(s, "New team member", M, 1.45, 6.0, 0.32, size=12.5, bold=True,
   color=ACCENT, caps=True)
tx(s, "Welcome to\nthe Team!", M, 1.95, 8.4, 2.0, size=63, bold=True, color=WHITE,
   spacing=0.95)
rect(s, M, 3.95, 1.05, 0.06, ACCENT)
tx(s, "Davis Fern\u00e1ndez", M, 4.35, 7.0, 0.55, size=27, bold=True, color=WHITE)
tx(s, "Independent IT Technician  \u00b7  Chiriqu\u00ed", M, 4.95, 8.0, 0.4, size=17,
   color=ACCENT)
tx(s, "Welcome Kit  \u00b7  David, Chiriqu\u00ed  \u00b7  October 2026", M, 5.75, 8.0, 0.35,
   size=13, color=RGBColor(0xB8, 0xC6, 0xD4))
notes(s, "Good morning, everyone. My name is Davis Fernandez and I am an independent "
         "IT technician. Thank you for the warm welcome. Today I am going to present "
         "my welcome kit for a new colleague from another country. I will talk about "
         "four things: who I am and my routine, how to get to my office in David, the "
         "places to eat near my office, and the safety rules I use when I visit "
         "clients all over Chiriqui. It will take about five minutes, and at the end "
         "you can ask me questions.")

# ================================================================ SLIDE 2
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
header(s, "About me", "This is me, my work and my week")

cards = [
    ("01", "Who I am",
     "Friendly, organized and a little shy. I like solving problems and helping "
     "people."),
    ("02", "Where I work",
     "Independent IT technician. I visit clients all over Chiriqu\u00ed; my main client "
     "is PikTa."),
    ("03", "My schedule",
     "I start at 8:00 a.m. and finish at 5:00 p.m., from Monday to Friday."),
    ("04", "This week",
     "This week I am setting up a network and visiting clients in Bugaba and "
     "Boquete."),
]
cw, gap = 2.82, 0.2
for i, (num, title, body) in enumerate(cards):
    l = M + i * (cw + gap)
    card(s, l, 2.25, cw, 2.2)
    circle(s, l + 0.28, 2.48, 0.52, NAVY)
    tx(s, num, l + 0.28, 2.48, 0.52, 0.52, size=13, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    tx(s, title, l + 0.28, 3.16, cw - 0.5, 0.3, size=15.5, bold=True, color=NAVY)
    tx(s, body, l + 0.28, 3.5, cw - 0.52, 1.0, size=12.5, color=MUTED, spacing=1.18)

card(s, M, 4.78, W - 2 * M, 1.68, fill=NAVY, line=None)
tx(s, "Daily routine", M + 0.42, 5.0, 2.1, 0.3, size=11, bold=True, color=ACCENT,
   caps=True)
tx(s, "8:00 a.m.  \u2192  5:00 p.m.", M + 0.42, 5.3, 3.45, 0.5, size=20, bold=True,
   color=WHITE)
steps = [("8:00", "Check messages and\nplan the visits"),
         ("10:00", "Visit clients and\nfix equipment"),
         ("4:30", "Update the tickets\nand plan tomorrow")]
for i, (hour, desc) in enumerate(steps):
    l = 4.62 + i * 2.62
    circle(s, l, 5.07, 0.2, ACCENT)
    tx(s, hour, l + 0.32, 4.98, 1.4, 0.3, size=12, bold=True, color=ACCENT)
    tx(s, desc, l + 0.32, 5.33, 2.15, 0.75, size=11, color=RGBColor(0xD3, 0xDE, 0xE8),
       spacing=1.15)
    if i < 2:
        rect(s, l + 2.1, 5.14, 0.4, 0.02, RGBColor(0x3C, 0x5B, 0x78))

footer(s, 2)
notes(s, "Let me start with me. My name is Davis Fernandez and I am an independent IT "
         "technician. I am friendly and organized, and I am a little shy, so please "
         "be patient with me at the beginning. I do not work in one office all day: "
         "I visit clients all over Chiriqui province, and my base is in David. My "
         "schedule is from eight in the morning to five in the afternoon, Monday to "
         "Friday. This week I am setting up a new network and I am visiting clients "
         "in Bugaba and Boquete. My daily routine is like this. At eight o'clock I "
         "check my messages and I plan the visits. Then, from ten to four, I visit "
         "clients and I fix their equipment. At four thirty I update the tickets and "
         "I plan tomorrow.")

# ================================================================ SLIDE 3
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
header(s, "How to get there", "How to get to my workplace")

card(s, M, 2.2, 4.5, 1.75, fill=NAVY, line=None)
tx(s, "My main client \u00b7 David", M + 0.35, 2.38, 3.9, 0.3, size=10, bold=True,
   color=ACCENT, caps=True)
tx(s, "Restaurante PikTa", M + 0.35, 2.66, 3.85, 0.62, size=17, bold=True,
   color=WHITE, spacing=1.1)
tx(s, "C. Ar\u00edstides Romero, David", M + 0.35, 3.3, 3.9, 0.3, size=11.5,
   color=ACCENT)
tx(s, "My base: La Riviera, Calle 5.", M + 0.35, 3.58, 3.9, 0.28, size=9.5,
   color=RGBColor(0xC9, 0xD5, 0xE1))

tx(s, "By car \u00b7 step by step", M, 4.12, 4.2, 0.3, size=10, bold=True, color=ACCENT,
   caps=True)
steps = [("First,", "leave my base and go east on Calle 5."),
         ("Then,", "turn left on Calle 4."),
         ("After that,", "go north to Calle Ar\u00edstides Romero."),
         ("Finally,", "turn right; PikTa is near Ozaga Market.")]
y = 4.42
for i, (word, text) in enumerate(steps):
    circle(s, M, y + 0.03, 0.3, ACCENT)
    tx(s, str(i + 1), M, y + 0.03, 0.3, 0.3, size=10.5, bold=True, color=NAVY,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box = s.shapes.add_textbox(Inches(M + 0.44), Inches(y - 0.02), Inches(4.15), Inches(0.62))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.line_spacing = 1.12
    r1 = p.add_run()
    r1.text = word + " "
    _count(r1.text)
    r1.font.name = FONT
    r1.font.size = Pt(12.5)
    r1.font.bold = True
    r1.font.color.rgb = NAVY
    r2 = p.add_run()
    r2.text = text
    _count(r2.text)
    r2.font.name = FONT
    r2.font.size = Pt(12.5)
    r2.font.color.rgb = MUTED
    y += 0.52

draw_map(s, 5.6, 2.2, 6.88, 4.45)
tx(s, "From my base in La Riviera to PikTa \u00b7 5 minutes by car",
   5.6, 6.68, 6.9, 0.25, size=9.5, color=MUTED, italic=True)

footer(s, 3)
notes(s, "Now, how to get to my workplace. My base is in La Riviera, on Calle five, "
         "but I do not stay there all day: my main client is Restaurante PikTa, on "
         "Calle Aristides Romero, and I also visit clients in Bugaba, Boquete and "
         "Volcan. On the map, the orange line is the route from my base to PikTa. "
         "First, I leave my base and I go east on Calle five. Then, I turn left on "
         "Calle four. After that, I go north to Calle Aristides Romero. Finally, I "
         "turn right, and PikTa is there, near Ozaga Market. It takes about five "
         "minutes by car. When I work in the restaurant, I park on the street behind "
         "the building. My advice is to leave early, because the traffic is heavy at "
         "eight in the morning.")

# ================================================================ SLIDE 4
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
header(s, "Around the office  \u00b7  Safety first", "Food, family and safety on the road")

card(s, M, 2.1, 5.95, 2.2)
tx(s, "Places to eat", M + 0.3, 2.28, 3.0, 0.28, size=10, bold=True, color=BLUE, caps=True)
tx(s, "The bakery on the corner", M + 0.3, 2.56, 5.1, 0.32, size=15, bold=True,
   color=NAVY)
tx(s, "Coffee, bread and fruit for about $2.", M + 0.3, 2.86, 5.35, 0.3, size=11,
   color=MUTED)
tx(s, "The market", M + 0.3, 3.22, 2.6, 0.32, size=15, bold=True, color=NAVY)
tx(s, "Rice with chicken, beans and plantain for about $5.",
   M + 0.3, 3.52, 5.35, 0.3, size=11, color=MUTED)
tx(s, "There aren't any restaurants inside my building, but there are a lot of "
      "places near. Some are cheap and many people eat at the market; I don't have "
      "much time.",
   M + 0.3, 3.83, 5.4, 0.3, size=9.5, italic=True, color=RGBColor(0x3D, 0x5A, 0x74))

card(s, 6.98, 2.1, 5.5, 2.2)
tx(s, "Yesterday and tomorrow", 7.28, 2.28, 4.6, 0.28, size=10, bold=True, color=BLUE,
   caps=True)
tx(s, "Yesterday", 7.28, 2.56, 1.2, 0.28, size=10.5, bold=True, color=ACCENT, caps=True)
tx(s, "I finished work and I went to dinner with my family.",
   7.28, 2.83, 4.9, 0.3, size=11.5, color=TEXT)
tx(s, "Tomorrow", 7.28, 3.2, 1.2, 0.28, size=10.5, bold=True, color=ACCENT, caps=True)
tx(s, "I will have a long day: three clients in David, Bugaba and Boquete.",
   7.28, 3.47, 4.9, 0.55, size=11.5, color=TEXT, spacing=1.15)

# --- seguridad
tx(s, "Safety first", M, 4.5, 3.0, 0.3, size=10, bold=True, color=ACCENT, caps=True)

card(s, M, 4.82, 2.55, 1.84)
tx(s, "Emergency", M + 0.24, 5.0, 2.1, 0.26, size=9.5, bold=True, color=BLUE, caps=True)
tx(s, "911", M + 0.24, 5.26, 2.0, 0.6, size=34, bold=True, color=NAVY)
tx(s, "Police, ambulance and fire department.", M + 0.24, 5.94, 2.15, 0.5, size=9.5,
   color=MUTED, spacing=1.1)

card(s, M + 2.75, 4.82, 3.95, 1.84)
tx(s, "First-aid kit", M + 2.99, 5.0, 3.2, 0.26, size=9.5, bold=True, color=BLUE,
   caps=True)
tx(s, "In my car, in a red box", M + 2.99, 5.26, 3.4, 0.3, size=14, bold=True,
   color=NAVY)
tx(s, "Ask me for the key. If someone is bleeding, apply pressure and call 911.",
   M + 2.99, 5.64, 3.5, 0.75, size=10, color=MUTED, spacing=1.15)

card(s, M + 6.9, 4.82, 3.73, 1.84)
tx(s, "Three simple rules", M + 7.14, 5.0, 3.2, 0.26, size=9.5, bold=True,
   color=BLUE, caps=True)
rules = ["Check the car before a long trip.",
         "Share your route with your family.",
         "Back up your data before every visit."]
for i, r in enumerate(rules):
    yy = 5.3 + i * 0.42
    rect(s, M + 7.14, yy + 0.07, 0.07, 0.22, ACCENT)
    tx(s, r, M + 7.32, yy, 3.2, 0.3, size=10, color=TEXT)

footer(s, 4)
notes(s, "This last slide is about the office and about safety. There are a lot of "
         "places to eat near my office. The bakery on the corner is cheap: coffee, "
         "bread and fruit for about two dollars. The market is also good: a typical "
         "lunch is rice with chicken, beans and plantain for about five dollars. "
         "Some places are expensive and many people eat at the market, but I do not "
         "have much time, so I usually eat near the office. And this was my week: "
         "yesterday I finished work and I went to dinner with my family. Tomorrow I "
         "will have a long day: I will visit three clients in David, Bugaba and "
         "Boquete, so I will leave home very early. Now, safety. The emergency "
         "number is 911. I have the first-aid kit in my car, in a red box, and I "
         "have the key. If someone is bleeding, apply pressure and call 911. "
         "Finally, three simple rules: check the car before a long trip, share your "
         "route with your family, and back up your data before every visit. Thank "
         "you very much. Do you have any questions?")

base = r"C:\Users\DAVIS\Documents\Induction_Project"
out = os.path.join(base, "Induction_Project_Davis_Fernandez_FINAL.pptx")
tmp = os.path.join(base, "_build.pptx")
try:
    prs.save(out)
    saved = out
except PermissionError:
    prs.save(tmp)
    saved = tmp
print("Saved:", saved)

target = os.path.join(base, "Induction_Project_Davis_Fernandez.pptx")
for dst in (out, target):
    if os.path.abspath(dst) == os.path.abspath(saved):
        print("Guardado directamente:", dst)
        continue
    try:
        shutil.copyfile(saved, dst)
        print("Actualizado:", dst)
    except PermissionError:
        print("Abierto en PowerPoint (no se pudo actualizar):", dst)

print("Palabras visibles (sin etiquetas del mapa):", len(WORDS))
