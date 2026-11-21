# -*- coding: utf-8 -*-
"""Welcome Kit - Induction Project (4 slides).
Perfil: tecnico de soporte IT independiente, base en David, Chiriqui.
Estilo visual: menta / Canva (chips, tarjetas redondeadas, sombras suaves).
Alineado con la rubrica 'Final Project Induction (Instructions + Rubric)'."""

import os
import re
import shutil

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- paleta
NAVY = RGBColor(0x0E, 0x1B, 0x4D)
NAVY_SOFT = RGBColor(0x1C, 0x2E, 0x6E)
TEAL = RGBColor(0x12, 0xC2, 0xA0)
TEAL_DARK = RGBColor(0x0B, 0x8F, 0x77)
YELLOW = RGBColor(0xFF, 0xC5, 0x3D)
CORAL = RGBColor(0xFF, 0x6B, 0x8A)
BLUE = RGBColor(0x3D, 0x6B, 0xFF)
BG = RGBColor(0xE9, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD = WHITE
LINE = RGBColor(0xD9, 0xF0, 0xE7)
TEXT = RGBColor(0x18, 0x2A, 0x4A)
MUTED = RGBColor(0x5C, 0x6E, 0x7E)
SHADOW = RGBColor(0xCD, 0xEE, 0xE2)
ACCENT = TEAL
MINT_SOFT = RGBColor(0xD6, 0xF6, 0xEB)

FONT = "Segoe UI"
W, H = 13.333, 7.5
M = 0.85

WORDS = []


def _count(text):
    WORDS.extend(re.findall(r"[A-Za-z']+", text))


# ---------------------------------------------------------------- helpers
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


def soft_shadow(shp, blur=16, dist=5, color=SHADOW, alpha=70):
    spPr = shp._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    effect = spPr.makeelement(qn("a:effectLst"), {})
    shd = effect.makeelement(qn("a:outerShdw"), {
        "blurRad": str(int(blur * 12700)),
        "dist": str(int(dist * 12700)),
        "dir": "5400000",
        "algn": "ctr",
        "rotWithShape": "0",
    })
    clr = shd.makeelement(qn("a:srgbClr"), {"val": str(color)})
    clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 1000))}))
    shd.append(clr)
    effect.append(shd)
    spPr.append(effect)


def card(slide, l, t, w, h, fill=CARD, radius=0.13, line=None, sh=True):
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
    if sh:
        soft_shadow(shp)
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


def pill(slide, l, t, text, fill=WHITE, dot=None, color=NAVY, size=10.0, h=0.4,
         bold=True, caps=False):
    cw = 0.080 * size / 10.0
    w = 0.46 + len(text) * cw + (0.22 if dot is not None else 0.0)
    card(slide, l, t, w, h, fill=fill, radius=0.5, line=None, sh=False)
    x = l + 0.23
    if dot is not None:
        d = 0.12
        circle(slide, l + 0.18, t + (h - d) / 2, d, dot)
        x = l + 0.42
    tx(slide, text, x, t, w - (x - l) - 0.06, h, size=size, bold=bold,
       color=color, anchor=MSO_ANCHOR.MIDDLE, caps=caps)
    return w


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def header(slide, kicker, title, accent=ACCENT):
    pill(slide, M, 0.48, kicker, dot=accent, size=9.5, h=0.36, caps=True)
    tx(slide, title, M, 0.95, 11.6, 0.78, size=37, bold=True, color=NAVY,
       spacing=0.98)


def footer(slide, n):
    tx(slide, "Induction Project  \u00b7  Davis Fern\u00e1ndez", M, H - 0.5, 6.0, 0.3,
       size=9, color=MUTED, count=False)
    tx(slide, f"{n:02d}", W - M - 0.5, H - 0.5, 0.5, 0.3, size=9, color=MUTED,
       align=PP_ALIGN.RIGHT, count=False)


def rich(slide, l, t, w, h, head, rest, size=10.5, head_color=NAVY, color=MUTED,
         spacing=1.12):
    _count(head + " ")
    _count(rest)
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.line_spacing = spacing
    r1 = p.add_run()
    r1.text = head + " "
    r1.font.name = FONT
    r1.font.size = Pt(size)
    r1.font.bold = True
    r1.font.color.rgb = head_color
    r2 = p.add_run()
    r2.text = rest
    r2.font.name = FONT
    r2.font.size = Pt(size)
    r2.font.color.rgb = color
    return box


prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]
AVATAR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "avatar-davis.png")

# ================================================================ SLIDE 1
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
circle(s, 9.2, -2.2, 5.6, WHITE)
circle(s, -1.6, 5.2, 2.8, WHITE)

# retrato a la derecha
circle(s, 7.85, 1.1, 5.0, YELLOW)
if os.path.exists(AVATAR):
    s.shapes.add_picture(AVATAR, Inches(7.85), Inches(1.1), Inches(5.0), Inches(5.0))

# mini icono de laptop
card(s, 7.3, 1.55, 0.72, 0.72, fill=WHITE, radius=0.24)
rect(s, 7.44, 1.72, 0.44, 0.28, NAVY)
rect(s, 7.4, 2.03, 0.52, 0.06, NAVY)

circle(s, 11.95, 1.3, 0.26, YELLOW)
circle(s, 7.42, 4.95, 0.2, CORAL)
circle(s, 12.2, 5.55, 0.34, TEAL)

pill(s, M, 1.15, "New team member", dot=TEAL, size=9.5, h=0.36, caps=True)
tx(s, "Welcome to\nthe Team!", M, 1.75, 7.0, 1.95, size=58, bold=True, color=NAVY,
   spacing=0.95)
rect(s, M + 0.03, 3.78, 1.1, 0.09, TEAL)
tx(s, "Davis Fern\u00e1ndez", M, 4.05, 6.6, 0.5, size=26, bold=True, color=NAVY)
tx(s, "Independent IT Technician  \u00b7  Chiriqu\u00ed", M, 4.62, 6.8, 0.35, size=15,
   color=TEAL_DARK)
tx(s, "David, Chiriqu\u00ed  \u00b7  October 2026", M, 5.02, 6.4, 0.3, size=12,
   color=MUTED)
pill(s, M, 5.52, "4 sections  \u00b7  5 minutes", fill=WHITE, size=10, h=0.36)

# franja inferior de color
rect(s, 0, H - 0.14, 3.34, 0.14, TEAL)
rect(s, 3.34, H - 0.14, 3.33, 0.14, YELLOW)
rect(s, 6.67, H - 0.14, 3.33, 0.14, CORAL)
rect(s, 10.0, H - 0.14, 3.33, 0.14, BLUE)

notes(s, "Good morning, everyone. My name is Davis Fernández and I am an independent "
         "IT technician. Thank you for the warm welcome. Today I am going to present "
         "my welcome kit for a new colleague from another country. I will talk about "
         "four things: who I am and what my routine is, how to get to my office in "
         "David, the places to eat near my office, and the safety rules I use when I "
         "visit clients all over Chiriquí. It will take about five minutes, and at "
         "the end you can ask me questions.")

# ================================================================ SLIDE 2
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
circle(s, 11.9, 5.7, 2.8, WHITE)

header(s, "About me", "Hello, I'm Davis")
tx(s, "Independent IT technician  \u00b7  Chiriqu\u00ed", M, 1.8, 6.6, 0.3, size=15,
   color=MUTED)

# chips de personalidad
x = M
for label, dotc in (("Friendly", TEAL), ("Organized", YELLOW), ("A little shy", CORAL)):
    x += pill(s, x, 2.24, label, dot=dotc, size=10.5, h=0.4) + 0.18

# tres tarjetas de rutina
cards = [
    ("01", "Who I am",
     "I visit clients all over Chiriqu\u00ed and I love solving problems.", TEAL),
    ("02", "Where I work",
     "My main client is Restaurante PikTa, in David.", YELLOW),
    ("03", "My schedule",
     "Monday to Friday\n8:00 a.m. \u2013 5:00 p.m.", CORAL),
]
cw, gap = 2.1, 0.16
for i, (num, title, body, col) in enumerate(cards):
    l = M + i * (cw + gap)
    card(s, l, 2.9, cw, 1.75)
    circle(s, l + 0.22, 3.06, 0.42, col)
    tx(s, num, l + 0.22, 3.06, 0.42, 0.42, size=10.5, bold=True,
       color=NAVY if col == YELLOW else WHITE, align=PP_ALIGN.CENTER,
       anchor=MSO_ANCHOR.MIDDLE)
    tx(s, title, l + 0.22, 3.6, cw - 0.4, 0.28, size=13.5, bold=True, color=NAVY)
    tx(s, body, l + 0.22, 3.9, cw - 0.42, 0.7, size=10.5, color=MUTED, spacing=1.14)

# tarjeta "This week"
card(s, M, 4.85, 6.62, 1.55, fill=NAVY, radius=0.1)
tx(s, "This week", M + 0.35, 5.05, 3.0, 0.28, size=9.5, bold=True, color=YELLOW,
   caps=True)
tx(s, "I am setting up a network and visiting clients in Bugaba and Boquete.",
   M + 0.35, 5.36, 5.9, 0.9, size=15.5, color=WHITE, spacing=1.18)

# retrato a la derecha
circle(s, 7.95, 1.25, 4.9, YELLOW)
if os.path.exists(AVATAR):
    s.shapes.add_picture(AVATAR, Inches(7.95), Inches(1.25), Inches(4.9), Inches(4.9))
card(s, 7.5, 1.7, 0.72, 0.72, fill=WHITE, radius=0.24)
rect(s, 7.64, 1.87, 0.44, 0.28, NAVY)
rect(s, 7.6, 2.18, 0.52, 0.06, NAVY)
circle(s, 11.75, 5.55, 0.28, CORAL)
circle(s, 12.08, 2.0, 0.2, TEAL)

footer(s, 2)
notes(s, "Let me start by telling you about myself. My name is Davis Fernández and "
         "I am an independent IT technician. I am friendly and organized, and I am a "
         "little shy, so please be patient with me at the beginning. I do not work in "
         "one office all day: I visit clients all over Chiriquí province, and my base "
         "is in David. My schedule is from eight in the morning to five in the "
         "afternoon, Monday to Friday. This week I am setting up a new network and I "
         "am visiting clients in Bugaba and Boquete. "
         "At eight o'clock I check my messages and I plan the visits. Then, from ten "
         "to four, I visit clients and I fix their equipment. At four thirty I update "
         "the tickets and I plan tomorrow's visits.")

# ================================================================ SLIDE 3
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
header(s, "How to get there", "How to get to my workplace")

# tarjeta del cliente principal
card(s, M, 2.1, 4.45, 1.6, fill=NAVY, radius=0.1)
tx(s, "My main client  \u00b7  David", M + 0.32, 2.28, 3.9, 0.28, size=9.5, bold=True,
   color=TEAL, caps=True)
tx(s, "Restaurante PikTa", M + 0.32, 2.56, 3.85, 0.42, size=18, bold=True,
   color=WHITE)
tx(s, "C. Ar\u00edstides Romero  \u00b7  50 m from Julio Gato", M + 0.32, 3.0, 3.95,
   0.3, size=9.5, color=YELLOW)
tx(s, "My base: La Riviera, Calle 5", M + 0.32, 3.28, 3.9, 0.28, size=9.5,
   color=RGBColor(0xC9, 0xD5, 0xE1))

# tarjeta con los pasos
card(s, M, 3.9, 4.45, 2.86)
tx(s, "By car  \u00b7  step by step", M + 0.3, 4.06, 4.0, 0.28, size=9.5, bold=True,
   color=TEAL_DARK, caps=True)
steps = [("First,", "from the Romero La Riviera bus stop, go north on Carretera 48."),
         ("Then,", "continue on Carretera 48 to Av. 8a Oeste."),
         ("After that,", "go straight until Calle Ar\u00edstides Romero."),
         ("Finally,", "turn left; PikTa is on the left, next to Hospital Chiriqu\u00ed.")]
y = 4.42
for i, (word, text) in enumerate(steps):
    circle(s, M + 0.3, y + 0.02, 0.3, TEAL)
    tx(s, str(i + 1), M + 0.3, y + 0.02, 0.3, 0.3, size=10, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rich(s, M + 0.74, y - 0.03, 3.45, 0.55, word, text, size=10.5)
    if i < 3:
        rect(s, M + 0.74, y + 0.56, 3.4, 0.008, LINE)
    y += 0.6

# tarjeta del mapa
card(s, 5.6, 2.1, 6.88, 4.67, radius=0.07)
_map = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mapa-pikta.png")
if os.path.exists(_map):
    _mh = 3.9
    _mw = _mh * 841 / 726
    s.shapes.add_picture(_map, Inches(5.6 + (6.88 - _mw) / 2), Inches(2.28),
                         Inches(_mw), Inches(_mh))
tx(s, "From La Riviera to PikTa  \u00b7  about 10 minutes by car", 5.85, 6.38,
   4.6, 0.28, size=9.5, color=MUTED, italic=True)
pill(s, 10.48, 6.34, "5 km", fill=MINT_SOFT, color=NAVY, size=9.5, h=0.32)

footer(s, 3)
notes(s, "Now, let me tell you how to get to my workplace. My base is in La Riviera, "
         "on Calle 5, but I do not stay there all day: my main client is Restaurante "
         "PikTa, on Calle Arístides Romero, 50 meters from Julio Gato, and I also "
         "visit clients in Bugaba, "
         "Boquete and Volcán. On the map, the blue line is the route from my base to "
         "PikTa. First, I leave my base and I drive to the Romero La Riviera bus stop. "
         "From there, I go north on Carretera 48, and I continue on Avenida 8a Oeste, "
         "also called Avenida La Libertad. After that, I go straight until Calle "
         "Arístides Romero. Finally, I turn left, and PikTa is on the left, next to "
         "Hospital Chiriquí. It takes about ten minutes by car, about five kilometers. "
         "When I work in the restaurant, I park on the street behind the building. My advice "
         "is to leave early: the traffic is heavy at eight.")

# ================================================================ SLIDE 4
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, BG)
header(s, "Around the office  \u00b7  Safety first", "Food, family and safety")

# --- lugares para comer
card(s, M, 2.0, 6.05, 2.6)
tx(s, "Places to eat", M + 0.3, 2.18, 3.0, 0.28, size=9.5, bold=True,
   color=TEAL_DARK, caps=True)

card(s, M + 0.3, 2.52, 0.5, 0.5, fill=CORAL, radius=0.28, sh=False)
tx(s, "DF", M + 0.3, 2.52, 0.5, 0.5, size=10, bold=True, color=WHITE,
   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
tx(s, "Panader\u00eda Don Felo", M + 0.95, 2.5, 3.8, 0.3, size=14, bold=True, color=NAVY)
tx(s, "Coffee, bread and fruit.", M + 0.95, 2.8, 3.8, 0.28, size=10, color=MUTED)
pill(s, M + 4.85, 2.58, "$2", fill=YELLOW, color=NAVY, size=9.5, h=0.34)

card(s, M + 0.3, 3.22, 0.5, 0.5, fill=TEAL, radius=0.28, sh=False)
tx(s, "SX", M + 0.3, 3.22, 0.5, 0.5, size=10, bold=True, color=WHITE,
   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
tx(s, "S\u00faper Xtra", M + 0.95, 3.2, 3.8, 0.3, size=14, bold=True, color=NAVY)
tx(s, "Rice with chicken, beans and plantain.", M + 0.95, 3.5, 4.3, 0.28, size=10,
   color=MUTED)
pill(s, M + 4.85, 3.28, "$5", fill=YELLOW, color=NAVY, size=9.5, h=0.34)

tx(s, "There aren't any restaurants inside the office building, but there are a lot "
      "of places nearby. Some are cheap and many people eat at S\u00faper Xtra; I don't "
      "have much time.",
   M + 0.3, 3.9, 5.45, 0.55, size=9.5, italic=True, color=RGBColor(0x3D, 0x5A, 0x74),
   spacing=1.12)

# --- ayer y manana
card(s, 7.05, 2.0, 5.43, 2.6)
tx(s, "Yesterday and tomorrow", 7.35, 2.18, 4.6, 0.28, size=9.5, bold=True,
   color=TEAL_DARK, caps=True)
pill(s, 7.35, 2.52, "Yesterday", fill=RGBColor(0xFF, 0xE1, 0xE7), dot=CORAL, size=9.5,
     h=0.34)
tx(s, "I finished work and I went to dinner with my family.", 7.35, 2.94, 4.85, 0.3,
   size=10.5, color=TEXT, spacing=1.12)
pill(s, 7.35, 3.36, "Tomorrow", fill=RGBColor(0xE4, 0xEA, 0xFF), dot=BLUE, size=9.5,
     h=0.34)
tx(s, "I will have a long day: three clients in David, Bugaba and Boquete.", 7.35,
   3.78, 4.85, 0.6, size=10.5, color=TEXT, spacing=1.12)

# --- seguridad
card(s, M, 4.8, 2.5, 1.82)
tx(s, "Emergency", M + 0.26, 4.98, 2.0, 0.26, size=9.5, bold=True, color=TEAL_DARK,
   caps=True)
tx(s, "911", M + 0.26, 5.24, 2.0, 0.6, size=32, bold=True, color=NAVY)
tx(s, "Police, ambulance and fire department.", M + 0.26, 5.9, 2.05, 0.55, size=9,
   color=MUTED, spacing=1.1)

card(s, M + 2.7, 4.8, 4.05, 1.82)
tx(s, "First-aid kit", M + 2.96, 4.98, 3.2, 0.26, size=9.5, bold=True, color=TEAL_DARK,
   caps=True)
tx(s, "In my car, in a red box", M + 2.96, 5.24, 3.5, 0.3, size=13.5, bold=True,
   color=NAVY)
tx(s, "Ask me for the key. If someone is bleeding, apply pressure and call 911.",
   M + 2.96, 5.6, 3.55, 0.75, size=9.5, color=MUTED, spacing=1.14)

card(s, M + 6.95, 4.8, 4.68, 1.82)
tx(s, "Three simple rules", M + 7.21, 4.98, 3.2, 0.26, size=9.5, bold=True,
   color=TEAL_DARK, caps=True)
rules = ["Check the car before a long trip.",
         "Share your route with your family.",
         "Back up your data before every visit."]
for i, r in enumerate(rules):
    yy = 5.32 + i * 0.4
    rect(s, M + 7.21, yy + 0.08, 0.07, 0.2, TEAL)
    tx(s, r, M + 7.4, yy, 3.9, 0.3, size=9.5, color=TEXT)

footer(s, 4)
notes(s, "This last slide is about the office and safety. There are a lot of "
         "places to eat near my office. Panadería Don Felo is cheap: coffee, "
         "bread and fruit for about two dollars. Súper Xtra is also good: a typical "
         "lunch is rice with chicken, beans and plantain for about five dollars. "
         "Some places are expensive and many people eat at Súper Xtra, but I do not "
         "have much time, so I usually eat near the office. Yesterday I finished "
         "work and I went to dinner with my family. Tomorrow I "
         "will have a long day: I will visit three clients in David, Bugaba and "
         "Boquete. Now, safety. The emergency "
         "number is 911. I keep the first-aid kit in my car, in a red box, and I "
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
