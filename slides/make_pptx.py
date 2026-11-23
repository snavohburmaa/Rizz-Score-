"""Build the Rizz Score proposal deck.

The deck borrows the app's own design language: the dark purple ground from
public/style.css, the pill option chips, the hot-to-cool gradient, the score
meter and the face scan overlay. Every slide carries the same frame, and the
meter along the bottom fills up as the deck moves.

Layout is measured, not guessed. Text is wrapped with the real font files
through PIL, so every text box is given the exact height its lines need, and a
check at the end asserts that no two text boxes on a slide overlap.
"""

from PIL import ImageFont

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE

OUT = "/Users/khantwaiyan/rizz score/slides/Rizz Score Proposal.pptx"
FDIR = "/Users/khantwaiyan/Library/Fonts/"

# ---------------------------------------------------------------- the palette
# Straight out of public/style.css.
BG = RGBColor(0x0D, 0x0B, 0x14)
GLOW = RGBColor(0x20, 0x16, 0x3A)
CARD = RGBColor(0x17, 0x14, 0x22)
LINE = RGBColor(0x2E, 0x27, 0x42)
INK = RGBColor(0xF2, 0xEE, 0xFB)
MUTED = RGBColor(0x9B, 0x92, 0xB4)
DIM = RGBColor(0x6B, 0x63, 0x85)
HOT = RGBColor(0xFF, 0x4D, 0x7E)
COOL = RGBColor(0x7B, 0x61, 0xFF)
PINK = RGBColor(0xFF, 0x8F, 0xB0)
TRACK = RGBColor(0x24, 0x1D, 0x36)
VIEW = RGBColor(0x0A, 0x08, 0x12)
GRIDC = RGBColor(0x26, 0x1E, 0x45)

HEAD = "Inter Display"
BODY = "Inter"
MONO = "JetBrains Mono"

FILES = {
    (HEAD, True): "InterDisplay-Bold.otf",
    (HEAD, False): "InterDisplay-Regular.otf",
    (BODY, True): "Inter-SemiBold.otf",
    (BODY, False): "Inter-Regular.otf",
    (MONO, True): "JetBrainsMono-Bold.ttf",
    (MONO, False): "JetBrainsMono-Regular.ttf",
}

# Measured against PowerPoint's own output: a line of Inter takes 1.21 to 1.24
# times the font size, so 1.25 is a safe pitch, and PowerPoint wraps a shade
# earlier than PIL does.
PITCH = 1.25
WRAPSAFE = 0.96
SCALE = 4            # measure at 4x for sub pixel accuracy

W, H = 1920, 1080
MARGIN = 112
CW = W - MARGIN * 2
TOPBAR = 58
BODYTOP = 232
SAFE_BOTTOM = 988    # nothing below this; the meter lives at 1016

_cache = {}


def _font(name, bold, size):
    key = (name, bold, size)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(FDIR + FILES[(name, bold)], int(size * SCALE))
    return _cache[key]


def P(px):
    return Emu(int(round(px * 914400 / 144)))


def FS(px):
    return Pt(px / 2.0)


def text_w(s, name, bold, size, track=0.0):
    """Width in px of one line, including letter tracking."""
    f = _font(name, bold, size)
    return f.getlength(s) / SCALE + max(0, len(s) - 1) * track * size


def wrap(text, name, bold, size, width, track=0.0):
    """Greedy word wrap, honouring hard newlines."""
    out = []
    limit = width * WRAPSAFE
    for para in text.split("\n"):
        words, line = para.split(), ""
        for word in words:
            trial = word if not line else line + " " + word
            if text_w(trial, name, bold, size, track) <= limit or not line:
                line = trial
            else:
                out.append(line)
                line = word
        out.append(line)
    return out


# ------------------------------------------------------------------ paragraphs

def par(text, size=26, color=None, font=BODY, bold=False, caps=False, track=0.0,
        space=1.45, gap=0, align=PP_ALIGN.LEFT):
    return dict(text=text.upper() if caps else text, size=size,
                color=color or MUTED, font=font, bold=bold, track=track,
                space=space, gap=gap, align=align)


def micro(text, color=None, gap=0, size=21):
    return par(text, size=size, color=color or DIM, caps=True, track=0.20,
               space=1.2, gap=gap, bold=True)


def para_h(p, width):
    lines = wrap(p["text"], p["font"], p["bold"], p["size"], width, p["track"])
    return len(lines) * p["size"] * p["space"] * PITCH


def block_h(paras, width):
    return sum(para_h(p, width) + p["gap"] for p in paras)


# --------------------------------------------------------------------- drawing

prs = Presentation()
prs.slide_width, prs.slide_height = P(W), P(H)
BLANK = prs.slide_layouts[6]
BOXES = []           # text rectangles on the current slide, for the overlap check
DEVICES = []         # phone mock rectangles, so content cannot spill out of a card


def no_shadow(shape):
    shape.shadow.inherit = False
    return shape


def rect(s, x, y, w, h, shape=MSO_SHAPE.RECTANGLE):
    return no_shadow(s.shapes.add_shape(shape, P(x), P(y), P(w), P(h)))


def fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    return shape


def fill_grad(shape, a=HOT, b=COOL, angle=0.0):
    f = shape.fill
    f.gradient()
    f.gradient_angle = angle
    f.gradient_stops[0].color.rgb = a
    f.gradient_stops[0].position = 0.0
    f.gradient_stops[1].color.rgb = b
    f.gradient_stops[1].position = 1.0
    return shape


def hairline(s, x, y, w, color=LINE, h=1.4):
    bar = rect(s, x, y, w, h)
    fill_solid(bar, color)
    bar.line.fill.background()
    return bar


def panel(s, x, y, w, h, fill=CARD, border=LINE, radius=0.07):
    shape = rect(s, x, y, w, h, MSO_SHAPE.ROUNDED_RECTANGLE)
    shape.adjustments[0] = radius
    fill_solid(shape, fill)
    if border:
        shape.line.color.rgb = border
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def text(s, x, y, w, paras, anchor=MSO_ANCHOR.TOP, h=None, label=""):
    """One text frame per call, sized to the lines it actually needs."""
    need = block_h(paras, w)
    box_h = h if h is not None else need + paras[0]["size"] * 0.4
    tb = s.shapes.add_textbox(P(x), P(y), P(w), P(box_h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.clear()
    for i, p in enumerate(paras):
        pr = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        pr.alignment = p["align"]
        pr.line_spacing = p["space"]
        pr.space_before = Pt(p["gap"] / 2.0)
        pr.space_after = Pt(0)
        run = pr.add_run()
        run.text = p["text"]
        run.font.size = FS(p["size"])
        run.font.bold = p["bold"]
        run.font.name = p["font"]
        run.font.color.rgb = p["color"]
        if p["track"]:
            run.font._rPr.set("spc", str(int(round(p["track"] * p["size"] * 50))))
    BOXES.append((x, y, w, need, label or paras[0]["text"][:28]))
    return need


def chip(s, x, y, text_s, size=24, on=False, ghost=False, dash=False,
         color=None, pad_x=22, h=None):
    """A pill, like the option buttons in the app."""
    ch = h or int(size * PITCH + 26)
    cw = text_w(text_s, BODY, on, size) + pad_x * 2
    pill = rect(s, x, y, cw, ch, MSO_SHAPE.ROUNDED_RECTANGLE)
    pill.adjustments[0] = 0.5
    if on:
        fill_grad(pill)
        pill.line.fill.background()
    else:
        if ghost:
            pill.fill.background()
        else:
            fill_solid(pill, CARD)
        pill.line.color.rgb = LINE
        pill.line.width = Pt(1)
        if dash:
            pill.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    tf = pill.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    pr = tf.paragraphs[0]
    pr.alignment = PP_ALIGN.CENTER
    run = pr.add_run()
    run.text = text_s
    run.font.size = FS(size)
    run.font.name = BODY
    run.font.bold = on
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if on else (color or INK)
    BOXES.append((x, y, cw, ch, "chip " + text_s[:22]))
    return cw, ch


def chip_row(s, x, y, w, items, size=24, gap=12, row_gap=12, **kw):
    """Chips flowing onto as many rows as they need, like .opts in the app."""
    cx, cy, h = x, y, 0
    for item in items:
        cw = text_w(item, BODY, kw.get("on", False), size) + kw.get("pad_x", 22) * 2
        if cx > x and cx + cw > x + w:
            cy += h + row_gap
            cx = x
        _, h = chip(s, cx, cy, item, size=size, **kw)
        cx += cw + gap
    return cy + h - y


def button(s, x, y, w, label, ghost=False, size=26):
    h = int(size * PITCH + 30)
    b = rect(s, x, y, w, h, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.adjustments[0] = 0.26
    if ghost:
        b.fill.background()
        b.line.color.rgb = LINE
        b.line.width = Pt(1)
    else:
        fill_grad(b)
        b.line.fill.background()
    tf = b.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    pr = tf.paragraphs[0]
    pr.alignment = PP_ALIGN.CENTER
    run = pr.add_run()
    run.text = label
    run.font.size = FS(size)
    run.font.name = BODY
    run.font.bold = not ghost
    run.font.color.rgb = MUTED if ghost else RGBColor(0xFF, 0xFF, 0xFF)
    BOXES.append((x, y, w, h, "button " + label))
    return h


def meter(s, x, y, w, frac, h=8, bg=TRACK, reg=True):
    if reg:
        BOXES.append((x, y, w, h, "meter"))
    base = rect(s, x, y, w, h, MSO_SHAPE.ROUNDED_RECTANGLE)
    base.adjustments[0] = 0.5
    fill_solid(base, bg)
    base.line.fill.background()
    if frac > 0:
        fl = rect(s, x, y, max(h, w * frac), h, MSO_SHAPE.ROUNDED_RECTANGLE)
        fl.adjustments[0] = 0.5
        fill_grad(fl)
        fl.line.fill.background()


# ------------------------------------------------------------------- the frame

N_SLIDES = 7
_slide_no = 0


def slide(tag, notes="", cover=False):
    """A new slide with the shared frame, returns the slide."""
    global _slide_no
    BOXES.clear()
    DEVICES.clear()
    _slide_no += 1
    s = prs.slides.add_slide(BLANK)
    ground = rect(s, 0, 0, W, H)
    ground.line.fill.background()
    fill_grad(ground, GLOW, BG, angle=270.0)
    if notes:
        s.notes_slide.notes_text_frame.text = notes

    # wordmark, left
    text(s, MARGIN, TOPBAR, 700, [micro("Rizz Score   /   " + tag, color=DIM)],
         label="frame wordmark")
    # slide count, right
    text(s, W - MARGIN - 300, TOPBAR, 300,
         [par("%02d / %02d" % (_slide_no, N_SLIDES), size=21, color=DIM, font=MONO,
              space=1.2, align=PP_ALIGN.RIGHT)], label="frame count")
    hairline(s, MARGIN, TOPBAR + 46, CW, LINE)
    # the app's meter, filling up as the deck moves
    meter(s, MARGIN, 1016, CW, _slide_no / N_SLIDES, h=8, reg=False)
    return s


def heading(s, title_s, kicker=None, sub=None, size=76, y=BODYTOP, pad=54):
    """Kicker, big title, optional one line of support. Returns the next y."""
    paras = []
    if kicker:
        paras.append(micro(kicker, color=HOT))
    paras.append(par(title_s, size=size, color=INK, font=HEAD, bold=True,
                     track=-0.025, space=1.06, gap=22 if kicker else 0))
    if sub:
        paras.append(par(sub, size=29, color=MUTED, space=1.4, gap=20))
    h = text(s, MARGIN, y, min(CW, 1300), paras, label="heading")
    return y + h + pad


def row_list(s, x, y, w, items, idx_color=HOT, tag_x=None, tag_w=320,
             title_size=33, body_size=25, pad=30):
    """Numbered rows with a hairline between them. Optional mono tag on the right."""
    cy = y
    text_w_avail = w if tag_x is None else tag_x - x - 40
    for i, item in enumerate(items):
        head, bodytext = item[0], item[1]
        tag = item[2] if len(item) > 2 else None
        if i:
            hairline(s, x, cy, w, LINE)
            cy += pad
        n = text(s, x, cy + 4, 70,
                 [par("%02d" % (i + 1), size=22, color=idx_color, font=MONO, space=1.2)],
                 label="idx%d" % i)
        bh = text(s, x + 80, cy, text_w_avail - 80, [
            par(head, size=title_size, color=INK, font=HEAD, bold=True,
                track=-0.015, space=1.14),
            par(bodytext, size=body_size, color=MUTED, space=1.45, gap=14),
        ], label="row%d" % i)
        if tag:
            text(s, tag_x, cy + 4, tag_w,
                 [par(tag, size=22, color=COOL, font=MONO, space=1.3,
                      align=PP_ALIGN.RIGHT)], label="tag%d" % i)
        cy += max(bh, n) + pad
    return cy - y


def mono_lines(s, x, y, w, lines, size=24, color=MUTED, space=1.6, gap=0, label=""):
    paras = []
    for i, (txt, col) in enumerate(lines):
        paras.append(par(txt, size=size, color=col or color, font=MONO,
                         space=space, gap=0 if i == 0 else gap))
    return text(s, x, y, w, paras, label=label or "mono")


# ------------------------------------------------------------- the phone mocks

def device(s, x, y, w, h, caption=None):
    """The app's card on a dark screen, with a caption under it."""
    panel(s, x, y, w, h, CARD, LINE, radius=0.045)
    DEVICES.append((x, y, w, h))
    if caption:
        text(s, x, y + h + 18, w,
             [par(caption, size=19, color=DIM, caps=True, track=0.12, space=1.25,
                  align=PP_ALIGN.CENTER, bold=True)], label="caption")
    return x + 30, y + 30, w - 60      # inner box


def check_overlaps():
    """No two text boxes share space, and card content stays inside its card."""
    for ax, ay, aw, ah, an in BOXES:
        for dx, dy, dw, dh in DEVICES:
            cx, cyc = ax + aw / 2.0, ay + ah / 2.0
            if dx < cx < dx + dw and dy < cyc < dy + dh:
                assert ay >= dy + 16 and ay + ah <= dy + dh - 16, \
                    "%s spills out of its card (%d..%d vs %d..%d)" % (
                        an, ay, ay + ah, dy, dy + dh)
    for i in range(len(BOXES)):
        ax, ay, aw, ah, an = BOXES[i]
        assert ay >= TOPBAR - 2, "%s sits above the frame" % an
        assert ay + ah <= SAFE_BOTTOM + 2, \
            "%s runs to %d, past the safe bottom" % (an, ay + ah)
        assert ax >= 0 and ax + aw <= W + 1, "%s runs off the side" % an
        for j in range(i + 1, len(BOXES)):
            bx, by, bw, bh, bn = BOXES[j]
            if (ax < bx + bw - 2 and bx < ax + aw - 2
                    and ay < by + bh - 2 and by < ay + ah - 2):
                raise AssertionError("overlap: %r and %r" % (an, bn))


# ============================================================= 1. the topic
s = slide("Topic of focus", notes=(
    "Good morning. We are Khant Wai Yan Aung and Kaung San Thu, and our topic is a web "
    "application called Rizz Score. You answer ten silly questions about yourself and the app "
    "gives you a charisma score out of 100 with a rank and a short roast. The topic is a joke on "
    "purpose. The work behind it is not: it is ordinary client side and server side web "
    "development, which is exactly the material of this subject. What you see on the right is "
    "the screen we are aiming at."))

x, cy = MARGIN, 236
cy += text(s, x, cy, 900, [micro("Project proposal", color=HOT)]) + 26
w_rizz = text_w("Rizz", HEAD, True, 148, -0.03)
text(s, x, cy, w_rizz + 20, [par("Rizz", size=148, color=HOT, font=HEAD, bold=True,
                                 track=-0.03, space=1.0)], label="rizz")
cy += text(s, x + w_rizz + 36, cy, 760, [par("Score", size=148, color=COOL, font=HEAD,
                                             bold=True, track=-0.03, space=1.0)],
           label="score") + 30
cy += text(s, x, cy, 900, [par("Ten questions. One number. No mercy.", size=38,
                               color=INK, space=1.3)]) + 44
hairline(s, x, cy, 900)
cy += 34
cy += text(s, x, cy, 900, [micro("Team")]) + 18
names = text(s, x, cy, 460, [
    par("Khant Wai Yan Aung", size=28, color=INK, bold=True, space=1.5),
    par("Kaung San Thu", size=28, color=INK, bold=True, space=1.5, gap=8),
], label="names")
text(s, x + 480, cy, 300, [
    par("6703712", size=26, color=MUTED, font=MONO, space=1.62),
    par("6703124", size=26, color=MUTED, font=MONO, space=1.62, gap=8),
], label="ids")
cy += names + 52
text(s, x, cy, 900, [par("This is a joke app. Do not change your life based on the result.",
                         size=22, color=DIM, space=1.4)], label="fine")

# the result card the whole project points at
DX, DY, DW, DH = 1300, 236, 508, 696
ix, iy, iw = device(s, DX, DY, DW, DH, "the screen we are aiming at")
cy = iy + 8
cy += text(s, ix, cy, iw, [par("Certified Yapper", size=21, color=MUTED, caps=True,
                               track=0.18, space=1.3, align=PP_ALIGN.CENTER)]) + 10
cy += text(s, ix, cy, iw, [par("72", size=132, color=HOT, font=HEAD, bold=True,
                               track=-0.03, space=1.0, align=PP_ALIGN.CENTER)]) + 2
cy += text(s, ix, cy, iw, [par("out of 100", size=22, color=DIM, space=1.3,
                               align=PP_ALIGN.CENTER)]) + 22
meter(s, ix, cy, iw, 0.72)
cy += 30
cy += text(s, ix, cy, iw, [par("You have the talking part down. The listening part is still "
                               "loading.", size=25, color=INK, space=1.45)]) + 22
cy += text(s, ix, cy, iw, [
    par("Reply in under four hours once", size=22, space=1.4),
    par("Retire the hoodie on Fridays", size=22, space=1.4, gap=10),
    par("Ask one real question per conversation", size=22, space=1.4, gap=10),
], label="tips") + 26
button(s, ix, cy, iw, "Try again", ghost=True, size=24)
check_overlaps()

# ================================================================= 2. scope
s = slide("Scope", notes=(
    "This is the scope. The chips on top are the ten functions the app can do, and that is the "
    "whole product. The dashed chips underneath are what we cut on purpose. No accounts, no "
    "stored scores and no stored photos, because the moment we keep faces this stops being a "
    "harmless joke, and because two people have to finish this in a few weeks."))
cy = heading(s, "What it does, and what it will not do", kicker="Scope and functions",
             size=64, y=160, pad=48)
cy += text(s, MARGIN, cy, 900, [micro("Ten functions the app can do", color=COOL)]) + 20
cy += chip_row(s, MARGIN, cy, CW, [
    "Serve the ten question form",
    "Load the questions from the server",
    "Tap one option per question",
    "Optional seven second face scan",
    "POST the answers as JSON",
    "Ask an LLM for the score",
    "Return score, rank, roast, tips",
    "Validate and repair the reply",
    "Animate the score card",
    "Local fallback if the model fails",
], size=24) + 46
hairline(s, MARGIN, cy, CW)
cy += 36
cy += text(s, MARGIN, cy, 900, [micro("Out of scope, on purpose", color=HOT)]) + 20
cy += chip_row(s, MARGIN, cy, CW, [
    "Accounts or login", "A database of past scores", "A leaderboard",
    "Photo upload or storage", "Knowing who you really are",
    "Training our own model", "Payments",
], size=24, ghost=True, dash=True, color=DIM) + 34
text(s, MARGIN, cy, 1240, [par("Every dashed chip is a decision, not a gap. Accounts and "
                               "stored photos would turn a joke into something we have to "
                               "look after.", size=25, color=DIM, space=1.5)],
     label="scope note")
check_overlaps()

# ===================================================== 3. techniques, client
s = slide("Techniques, client side", notes=(
    "Slides three and four are the techniques, and they are the parts of the subject we are "
    "being taught. This one is the browser. There is no framework here on purpose: plain HTML, "
    "CSS and JavaScript, so the subject material is doing all of the work. The tag on the right "
    "of each row is the actual API we use."))
cy = heading(s, "Techniques in the browser", kicker="Subject knowledge, client side",
             size=64, y=176)
row_list(s, MARGIN, cy, CW, [
    ("Document structure and the DOM",
     "JavaScript builds the ten question blocks from the data the server sends.",
     "document\ncreateElement"),
    ("Asynchronous requests",
     "async and await to GET the questions and POST the answers, with try and catch.",
     "fetch\nawait"),
    ("CSS layout and animation",
     "Flexbox cards, a media query for the phone, keyframes for the sweep and the bar.",
     "flex\n@keyframes"),
    ("Canvas and the camera stream",
     "Frames drawn into a hidden canvas at 64 by 48 and compared for movement.",
     "getUserMedia\ndrawImage"),
], tag_x=1470, tag_w=338, title_size=31, body_size=25, pad=30)
check_overlaps()

# ===================================================== 4. techniques, server
s = slide("Techniques, server side", notes=(
    "This is the server and the model. Express gives us two routes, dotenv keeps the API key out "
    "of the code and out of the browser, and the prompt forces the model to answer in the exact "
    "shape on the right. The fourth row is the one we care about most: the model can always "
    "return something odd, so the server clamps the score and forces exactly three tips before "
    "the page ever sees it, and answers with a local score if the call fails."))
cy = heading(s, "Techniques on the server", kicker="Subject knowledge, server side",
             size=64, y=176)
LCOL = 1100
rh = row_list(s, MARGIN, cy, LCOL, [
    ("Node and Express routing",
     "A static route and two JSON routes.", "express"),
    ("Configuration and secrets",
     "The key is read from an ignored .env file.", "dotenv"),
    ("Prompt design and the API call",
     "Ten answers and one scan word become one prompt.", "fetch"),
    ("Validation and fallback",
     "Score clamped, three tips forced, local score if it fails.", "try / catch"),
], tag_x=MARGIN + LCOL - 240, tag_w=240, title_size=31, body_size=25, pad=26)

RX, RW = MARGIN + LCOL + 76, CW - LCOL - 76
panel(s, RX, cy - 6, RW, rh + 12)
cyr = cy + 34
cyr += text(s, RX + 36, cyr, RW - 72,
            [micro("The reply must look like this", color=COOL, size=20)]) + 16
cyr += mono_lines(s, RX + 36, cyr, RW - 72, [
    ("{", DIM),
    ('  "score": 72,', INK),
    ('  "title": "Certified Yapper",', INK),
    ('  "roast": "one or two lines",', INK),
    ('  "tips":  ["", "", ""]', INK),
    ("}", DIM),
], size=23, space=1.5, label="contract") + 30
hairline(s, RX + 36, cyr, RW - 72)
cyr += 26
cyr += text(s, RX + 36, cyr, RW - 72, [micro("The two routes", color=COOL)]) + 16
mono_lines(s, RX + 36, cyr, RW - 72, [
    ("GET   /api/questions", INK),
    ("POST  /api/rizz", INK),
], size=23, space=1.6, label="routes")
check_overlaps()

# ========================================================= 5. the prototype
s = slide("Expected result", notes=(
    "Here is the expected result as a prototype of the interface. Screen one is the question "
    "form: one label, four option chips, tap to answer, and the bar at the bottom shows how far "
    "in you are. Screen two is the face scan, and it is deliberate theatre: a mirrored frame, a "
    "sweep line, corner brackets and a countdown. The skip button is on screen the whole time, "
    "and the only thing that leaves the browser is one word."))
cy = heading(s, "The screens before the score", kicker="Expected result, prototype",
             size=56, y=148, pad=40)

PY, PH, PW = cy, 600, 476
# ---- screen one, the questions
ix, iy, iw = device(s, MARGIN, PY, PW, PH, "screen 1  ·  ten questions")
c = iy + 2
w_r = text_w("Rizz", HEAD, True, 32, -0.03)
text(s, ix, c, w_r + 8, [par("Rizz", size=32, color=HOT, font=HEAD, bold=True,
                             track=-0.03, space=1.1)], label="s1 rizz")
c += text(s, ix + w_r + 12, c, 240, [par("Score", size=32, color=COOL, font=HEAD, bold=True,
                                         track=-0.03, space=1.1)], label="s1 score") + 6
c += text(s, ix, c, iw, [par("Ten questions. One number. No mercy.", size=19, space=1.3)],
          label="s1 sub") + 24
c += text(s, ix, c, iw, [micro("Question 3 of 10", color=HOT, size=19)]) + 16
c += text(s, ix, c, iw, [par("How many unread messages right now?", size=24, color=INK,
                             bold=True, space=1.35)], label="s1 label") + 18
c += chip_row(s, ix, c, iw, ["Zero, I am a machine"], size=19, on=True, pad_x=16) + 10
c += chip_row(s, ix, c, iw, ["Under 10", "10 to 100", "Over 1000"], size=19, pad_x=16) + 26
meter(s, ix, c, iw, 0.3, h=6)
c += 28
c += button(s, ix, c, iw, "Judge me", size=24) + 16
text(s, ix, c, iw, [par("This is a joke app. Do not change your life based on the result.",
                        size=17, color=DIM, space=1.4, align=PP_ALIGN.CENTER)], label="s1 fine")

# ---- screen two, the face scan
SX = MARGIN + PW + 34
ix, iy, iw = device(s, SX, PY, PW, PH, "screen 2  ·  the face scan")
c = iy + 2
c += text(s, ix, c, iw, [par("Face scan", size=19, color=MUTED, caps=True, track=0.2,
                             space=1.3, align=PP_ALIGN.CENTER, bold=True)], label="s2 title") + 16
VH = 272
vp = panel(s, ix, c, iw, VH, VIEW, LINE, radius=0.05)
for k in range(1, 5):                                  # the faint tech grid
    hairline(s, ix + 6, c + VH * k / 5.0, iw - 12, GRIDC, h=1.2)
for k in range(1, 4):
    bar = rect(s, ix + iw * k / 4.0, c + 6, 1.2, VH - 12)
    fill_solid(bar, GRIDC)
    bar.line.fill.background()
ov = rect(s, ix + iw * 0.23, c + VH * 0.13, iw * 0.54, VH * 0.62, MSO_SHAPE.OVAL)
ov.fill.background()
ov.line.color.rgb = COOL
ov.line.width = Pt(1.5)
ov.line.dash_style = MSO_LINE_DASH_STYLE.DASH
glowband = rect(s, ix + 2, c + VH * 0.40, iw - 4, 26)        # the sweep, trailing glow
fill_solid(glowband, RGBColor(0x3A, 0x18, 0x32))
glowband.line.fill.background()
sweep = rect(s, ix + 2, c + VH * 0.40 + 26, iw - 4, 3)
fill_solid(sweep, PINK)
sweep.line.fill.background()
B = 34
for bx, by, in ((ix + 14, c + 14), (ix + iw - 14 - B, c + 14),
                (ix + 14, c + VH - 14 - B), (ix + iw - 14 - B, c + VH - 14 - B)):
    for seg in ((bx, by, B, 3), (bx, by, 3, B)):
        br = rect(s, *seg)
        fill_solid(br, HOT)
        br.line.fill.background()
text(s, ix + iw - 140, c + VH - 100, 80, [par("7", size=34, color=INK, font=HEAD, bold=True,
                                              space=1.1, align=PP_ALIGN.RIGHT)],
     label="s2 timer")
c += VH + 20
c += text(s, ix, c, iw, [par("Now pull your silliest face.", size=21, color=INK, space=1.4,
                             align=PP_ALIGN.CENTER)], label="s2 prompt") + 20
meter(s, ix, c, iw, 0.55, h=6)
c += 26
c += button(s, ix, c, iw, "Skip the scan", ghost=True, size=22) + 14
text(s, ix, c, iw, [par("Nothing is uploaded, recorded or saved.", size=17,
                        color=DIM, space=1.4, align=PP_ALIGN.CENTER)], label="s2 fine")

# ---- the rail
RX = MARGIN + (PW + 34) * 2
RW = W - MARGIN - RX
cyr = PY + 4
cyr += text(s, RX, cyr, RW, [micro("What the prototype has to prove", color=COOL)]) + 28
for i, (head, bodytext) in enumerate([
    ("One thing at a time",
     "One question on a phone, all ten on a laptop."),
    ("The scan is never forced",
     "Skip sits under the frame the whole time."),
    ("Nothing is uploaded",
     "One word leaves the browser, never a frame."),
    ("Every state is drawn",
     "Loading, offline and result all have a screen."),
]):
    if i:
        hairline(s, RX, cyr, RW)
        cyr += 24
    cyr += text(s, RX, cyr, RW, [
        par(head, size=30, color=INK, font=HEAD, bold=True, track=-0.015, space=1.16),
        par(bodytext, size=24, color=MUTED, space=1.45, gap=12),
    ], label="rail%d" % i) + 24
check_overlaps()

# ====================================================== 6. the result screen
s = slide("Expected result", notes=(
    "This is the card people will screenshot, and the three states behind it. The score counts "
    "up and the bar fills, so the number lands instead of just appearing. Offline Rizz is the "
    "local fallback: if the model is down we still answer, with the boring maths version, and "
    "the demo carries on. The disclaimer sits under the result every time, because the whole "
    "thing is comedy and we would rather say so than imply otherwise."))
cy = heading(s, "The score card, and the states behind it",
             kicker="Expected result, prototype", size=56, y=148, pad=40)

ix, iy, iw = device(s, MARGIN, cy, 476, 600, "screen 3  ·  the result card")
c = iy + 4
c += text(s, ix, c, iw, [par("Certified Yapper", size=20, color=MUTED, caps=True, track=0.18,
                             space=1.3, align=PP_ALIGN.CENTER)]) + 8
c += text(s, ix, c, iw, [par("72", size=108, color=HOT, font=HEAD, bold=True, track=-0.03,
                             space=1.0, align=PP_ALIGN.CENTER)]) + 2
c += text(s, ix, c, iw, [par("out of 100", size=20, color=DIM, space=1.3,
                             align=PP_ALIGN.CENTER)]) + 16
meter(s, ix, c, iw, 0.72)
c += 28
c += text(s, ix, c, iw, [par("You have the talking part down. The listening part is still "
                             "loading.", size=23, color=INK, space=1.45)]) + 16
c += text(s, ix, c, iw, [
    par("Reply in under four hours once", size=20, space=1.4),
    par("Retire the hoodie on Fridays", size=20, space=1.4, gap=8),
    par("Ask one real question per conversation", size=20, space=1.4, gap=8),
], label="s3 tips") + 16
button(s, ix, c, iw, "Try again", ghost=True, size=22)

RX = MARGIN + 476 + 60
RW = W - MARGIN - RX
cyr = cy + 4
cyr += text(s, RX, cyr, RW, [micro("Three states we have to draw", color=COOL)]) + 26
states = [
    ("Loading", "Reading your aura...",
     "A spinner and one line, so the wait is part of the joke."),
    ("Offline Rizz", "The judge is not answering right now.",
     "If the model fails, the server scores the answers itself and the card still fills in."),
    ("Result", "score 72 · title · roast · tips[3]",
     "The number counts up from zero and the bar fills with it."),
]
for i, (name, quote, note) in enumerate(states):
    if i:
        hairline(s, RX, cyr, RW)
        cyr += 24
    cyr += text(s, RX, cyr, RW, [
        par(name, size=27, color=INK, font=HEAD, bold=True, track=-0.015, space=1.16),
        par(quote if quote.startswith("score") else '"' + quote + '"',
            size=23, color=PINK, font=MONO, space=1.4, gap=10),
        par(note, size=23, color=MUTED, space=1.45, gap=10),
    ], label="state%d" % i) + 24
hairline(s, RX, cyr, RW)
cyr += 26
text(s, RX, cyr, RW, [par("A plain disclaimer sits under the result every time: this is "
                          "comedy, not a measurement of anybody.", size=22, color=DIM,
                          space=1.45)], label="disclaimer")
check_overlaps()

# ============================================================ 7. conclusion
s = slide("Conclusion  /  thank you", notes=(
    "To close. The scope is small on purpose, it uses the subject from the browser through to "
    "the server, nothing is stored about anyone, and there is a fallback so the demo works even "
    "if the model is down. We are asking for approval to build it, and we will bring a live "
    "demo. The sources on the right are the documentation we are building against. Thank you, "
    "and we are ready for questions."))
cy = heading(s, "A small app, finished properly", kicker="Conclusion", size=64, y=176)

LCOL = 880
rh = row_list(s, MARGIN, cy, LCOL, [
    ("The scope is small on purpose",
     "Ten questions, one scan, one card. Two people can finish that."),
    ("It uses the subject end to end",
     "DOM and fetch in the browser, Express and JSON on the server."),
    ("Nothing is stored about anyone",
     "No accounts, no scores kept, no frame ever uploaded."),
    ("The demo cannot go blank",
     "A local score answers whenever the model will not."),
], title_size=30, body_size=23, pad=26)

RX, RW = MARGIN + LCOL + 76, CW - LCOL - 76
panel(s, RX, cy - 6, RW, rh + 12)
cyr = cy + 32
cyr += text(s, RX + 36, cyr, RW - 72, [micro("References and sources", color=COOL)]) + 24
PITCH_ROW = 56
NAME_SP = PITCH_ROW / (23 * PITCH)
URL_SP = PITCH_ROW / (19 * PITCH)
refs = [
    ("MDN Web Docs", "developer.mozilla.org"),
    ("Express", "expressjs.com/en/guide/routing.html"),
    ("Node.js", "nodejs.org/docs"),
    ("dotenv", "github.com/motdotla/dotenv"),
    ("Groq API", "console.groq.com/docs"),
    ("Lecture notes", "this subject's slides and labs"),
]
text(s, RX + 36, cyr, 230,
     [par(n, size=23, color=INK, bold=True, space=NAME_SP) for n, _ in refs],
     label="ref names")
text(s, RX + 36 + 250, cyr, RW - 72 - 250,
     [par(u, size=19, color=MUTED, font=MONO, space=URL_SP) for _, u in refs],
     label="ref urls")
cyr += PITCH_ROW * len(refs) + 24
hairline(s, RX + 36, cyr, RW - 72)
cyr += 24
text(s, RX + 36, cyr, RW - 72,
     [par("On MDN we work from the Fetch, Canvas and MediaDevices pages. The Express guide "
          "covers routing and static files.", size=21, color=DIM, space=1.45)],
     label="ref note")

check_overlaps()

prs.save(OUT)
print("wrote", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
