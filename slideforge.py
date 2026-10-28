#!/usr/bin/env python3
"""
SlideForge — turn a Markdown file into an editable PowerPoint deck, in one command.

Why: "PDF/Word → PPT" tools are everywhere, but writers draft in Markdown.
SlideForge reads plain Markdown and builds a real .pptx — native text boxes,
not screenshots — so every slide stays editable in PowerPoint / WPS / Keynote.

Syntax:
  # Deck title            -> cover slide (front matter: first H1)
  # Section name          -> section divider slide (later H1s)
  ## Slide title          -> content slide
  - bullet                -> bullet (two spaces + - for level 2)
  > quote                 -> big-quote slide content
  1. step                 -> numbered bullet

Usage:
  pip install -r requirements.txt
  python3 slideforge.py examples/deck.md --theme dark --brand "@我的账号"
"""
import argparse, os, re, sys

THEMES = {
    "dark":  {"bg": (17, 17, 20),    "title": (255, 255, 255), "body": (225, 225, 230), "accent": (208, 0, 0),   "sub": (150, 150, 160)},
    "light": {"bg": (250, 250, 248), "title": (25, 25, 30),    "body": (55, 55, 65),    "accent": (190, 40, 40), "sub": (130, 130, 135)},
    "ocean": {"bg": (10, 35, 66),    "title": (245, 250, 255), "body": (210, 228, 245), "accent": (0, 180, 220), "sub": (150, 180, 205)},
}

def parse_md(text):
    """Return list of slides: {kind: cover|section|content|quote, title, items:[(level, text)]}"""
    slides, cur = [], None
    first_h1 = True
    def flush():
        nonlocal cur
        if cur: slides.append(cur); cur = None
    for raw in text.splitlines():
        line = raw.rstrip()
        m1 = re.match(r"^#\s+(.+)", line)
        m2 = re.match(r"^##\s+(.+)", line)
        mq = re.match(r"^>\s*(.+)", line)
        mb = re.match(r"^(\s*)[-*]\s+(.+)", line)
        mn = re.match(r"^\s*\d+[.)]\s+(.+)", line)
        if m1:
            flush(); cur = {"kind": "cover" if first_h1 else "section", "title": m1.group(1).strip(), "items": []}; first_h1 = False
        elif m2:
            flush(); cur = {"kind": "content", "title": m2.group(1).strip(), "items": []}
        elif mq:
            if cur is None or cur["kind"] != "quote":
                flush(); cur = {"kind": "quote", "title": "", "items": []}
            cur["items"].append((0, mq.group(1).strip()))
        elif mb:
            if cur is None: cur = {"kind": "content", "title": "", "items": []}
            cur["items"].append((1 if len(mb.group(1)) >= 2 else 0, mb.group(2).strip()))
        elif mn:
            if cur is None: cur = {"kind": "content", "title": "", "items": []}
            cur["items"].append((0, mn.group(1).strip()))
        elif line.strip():
            if cur is None: cur = {"kind": "content", "title": "", "items": []}
            if cur["kind"] in ("cover", "section") and not cur["items"]:
                cur["items"].append((0, line.strip()))  # subtitle on cover/section
            else:
                cur["items"].append((0, line.strip()))
    flush()
    return slides

def build(slides, out_path, theme="dark", brand=""):
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    T = {k: RGBColor(*v) for k, v in THEMES[theme].items()}
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = 13.333, 7.5
    total = len(slides)
    for idx, s in enumerate(slides, 1):
        sl = prs.slides.add_slide(blank)
        bg = sl.background.fill; bg.solid(); bg.fore_color.rgb = T["bg"]
        def box(x, y, w, h):
            tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); return tb.text_frame
        # accent bar
        shp = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(W), Inches(0.09)); shp.fill.solid(); shp.fill.fore_color.rgb = T["accent"]; shp.line.fill.background()
        if s["kind"] in ("cover", "section"):
            tf = box(0.9, 2.35 if s["kind"] == "cover" else 2.7, 11.5, 1.9); tf.word_wrap = True
            p = tf.paragraphs[0]; r = p.add_run(); r.text = s["title"]; r.font.size = Pt(54 if s["kind"] == "cover" else 44); r.font.bold = True; r.font.color.rgb = T["title"]
            if s["items"]:
                tf2 = box(0.92, 4.15 if s["kind"] == "cover" else 4.3, 11.0, 0.8)
                p2 = tf2.paragraphs[0]; r2 = p2.add_run(); r2.text = s["items"][0][1]; r2.font.size = Pt(22); r2.font.color.rgb = T["sub"]
        elif s["kind"] == "quote":
            tf = box(1.1, 2.2, 11.1, 3.2); tf.word_wrap = True
            for i, (_, txt) in enumerate(s["items"]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                r = p.add_run(); r.text = "“" + txt + "”" if i == 0 else txt; r.font.size = Pt(36); r.font.italic = True; r.font.color.rgb = T["title"]
        else:
            tf = box(0.9, 0.55, 11.5, 1.1)
            p = tf.paragraphs[0]; r = p.add_run(); r.text = s["title"]; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = T["title"]
            body = box(0.95, 1.85, 11.4, 4.9); body.word_wrap = True
            for i, (lvl, txt) in enumerate(s["items"]):
                p = body.paragraphs[0] if i == 0 else body.add_paragraph()
                p.level = lvl; p.space_after = Pt(12)
                r = p.add_run(); r.text = ("• " if lvl == 0 else "– ") + txt; r.font.size = Pt(22 if lvl == 0 else 19); r.font.color.rgb = T["body"]
        # footer: brand + page
        foot = box(0.9, 7.02, 8.0, 0.35); pf = foot.paragraphs[0]; rf = pf.add_run(); rf.text = brand; rf.font.size = Pt(12); rf.font.color.rgb = T["sub"]
        pg = box(12.1, 7.02, 0.8, 0.35); pp = pg.paragraphs[0]; pp.alignment = PP_ALIGN.RIGHT; rp = pp.add_run(); rp.text = f"{idx}/{total}"; rp.font.size = Pt(12); rp.font.color.rgb = T["sub"]
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    prs.save(out_path)
    return out_path

def main():
    ap = argparse.ArgumentParser(description="SlideForge — Markdown to PowerPoint")
    ap.add_argument("md", help="input .md file")
    ap.add_argument("--out", help="output .pptx (default: same name)")
    ap.add_argument("--theme", default="dark", choices=list(THEMES))
    ap.add_argument("--brand", default="")
    a = ap.parse_args()
    slides = parse_md(open(a.md, encoding="utf-8").read())
    if not slides: sys.exit("No slides found — start the file with '# Title'.")
    out = a.out or os.path.splitext(a.md)[0] + ".pptx"
    build(slides, out, a.theme, a.brand)
    print(f"Done: {len(slides)} slide(s) -> {out} (theme={a.theme})")

if __name__ == "__main__":
    try: import pptx  # noqa
    except ImportError: sys.exit("Need python-pptx: pip install -r requirements.txt")
    main()
