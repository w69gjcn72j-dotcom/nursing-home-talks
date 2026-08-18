#!/usr/bin/env python3
"""Add a study/talk to this collection, in English and/or Chinese.

    # English only
    python3 add-study.py study.docx --group "Term 1 2026" --ref "Mark 10:35-45"

    # English + Chinese in one go
    python3 add-study.py study.docx --cn-docx 中文.docx \
        --group "Term 1 2026" --cn-group "2026 上学期" --ref "Mark 10:35-45"

    # add the Chinese version of an entry published earlier
    python3 add-study.py 中文.docx --lang cn --n 7 --cn-group "2026 上学期"

    # verify the site is consistent
    python3 add-study.py --check

Writes NNN-slug.html / NNN-slug-cn.html beside index.html and inserts the matching
entry into the STUDIES list in index.html / index-cn.html. Nothing else is touched.

Needs python-docx only when reading a .docx:  pip3 install python-docx
"""
import argparse, html, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
CONF = os.path.join(HERE, "site.json")
INDEX = {"en": os.path.join(HERE, "index.html"), "cn": os.path.join(HERE, "index-cn.html")}

FONTS = ('<link rel="preconnect" href="https://fonts.googleapis.com">\n'
         '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
         '<link href="https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,300;'
         '0,9..144,400;0,9..144,600;1,9..144,400&family=Inter+Tight:wght@300;400;500;600&display=swap" '
         'rel="stylesheet">')

CN_CSS = (
    "  .count, .group .n, .item .ref, .item .go, .eyebrow, .langbar a, .tip b,\n"
    "  .back, .lang, footer, .bigidea b { letter-spacing: 0.05em; text-transform: none; }"
)

LANGS = {
    "en": dict(htmllang="en", other_label="中文", eyebrow="St Paul's Anglican Kogarah",
               sans="'Inter Tight', system-ui, -apple-system, sans-serif",
               serif="'Fraunces', Georgia, serif",
               bigidea="Big Idea",
               page_footer="Soli Deo Gloria &middot; <span>St Paul's Anglican Kogarah</span>", lang_css=""),
    "cn": dict(htmllang="zh-Hans", other_label="English", eyebrow="高嘉华圣保罗圣公会",
               sans="'Inter Tight', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', system-ui, sans-serif",
               serif="'Fraunces', 'Songti SC', 'STSong', Georgia, serif",
               bigidea="中心思想",
               page_footer="唯独荣耀归于神 &middot; <span>高嘉华圣保罗圣公会</span>", lang_css=CN_CSS),
}

PAGE = """<!DOCTYPE html>
<html lang="{htmllang}">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<link rel="apple-touch-icon" sizes="180x180" href="{prefix}-apple-touch-icon.png">
<link rel="icon" type="image/png" sizes="32x32" href="{prefix}-favicon-32.png">
<link rel="manifest" href="{manifest}">
<link rel="alternate" hreflang="{alt_hreflang}" href="{other_page}">
<meta name="theme-color" content="#0e1a2b">
<title>{page_title} — {eyebrow}</title>
{fonts}
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  :root {{
    --night:#0e1a2b; --night-soft:#142339; --night-card:#1a2c47;
    --brass:#d4a253; --brass-soft:#b08539; --cream:#f5ead4; --cream-dim:#c9bfa8;
    --line:rgba(212,162,83,0.18);
  }}
  html {{ font-size: {base_font}; -webkit-text-size-adjust: 100%; }}
  body {{ font-family: {sans}; background:var(--night); color:var(--cream); line-height:1.75; }}
  .wrap {{ max-width: 780px; margin: 0 auto; padding: 2rem 1.4rem 4rem; }}
  .topbar {{ display:flex; align-items:center; justify-content:space-between; gap:1rem; }}
  .back {{ font-size:0.85rem; letter-spacing:0.12em; text-transform:uppercase; color:var(--brass); text-decoration:none; }}
  .back:hover {{ color: var(--cream); }}
  .lang {{ font-size:0.74rem; letter-spacing:0.16em; text-transform:uppercase; color:var(--brass); text-decoration:none;
           border:1px solid var(--line); border-radius:999px; padding:0.25rem 0.8rem; white-space:nowrap; }}
  .lang:hover {{ color: var(--cream); border-color: rgba(212,162,83,0.42); }}
  h1 {{ font-family:{serif}; font-weight:400; font-size:clamp(1.7rem,4.5vw,2.5rem); line-height:1.15; color:var(--cream); margin:1.2rem 0 0.4rem; }}
  .meta {{ color:var(--brass); font-size:0.95rem; letter-spacing:0.04em; margin-bottom:1.6rem; }}
  .meta .note {{ color:var(--cream-dim); font-style:italic; }}
  .bigidea {{ background:var(--night-card); border-left:3px solid var(--brass); border-radius:3px; padding:1rem 1.2rem; margin-bottom:2.2rem; color:var(--cream); }}
  .bigidea b {{ color:var(--brass); font-weight:600; letter-spacing:0.08em; text-transform:uppercase; font-size:0.78rem; display:block; margin-bottom:0.3rem; }}
  .ms p {{ margin-bottom:0.85rem; font-weight:300; }}
  .ms h2 {{ font-family:{serif}; font-weight:600; font-size:1.15rem; color:var(--brass); margin:1.8rem 0 0.8rem; }}
  .ms ol, .ms ul {{ margin:0 0 1rem 1.3rem; }}
  .ms li {{ margin-bottom:0.5rem; font-weight:300; }}
  .ms blockquote {{ border-left:2px solid var(--brass-soft); padding-left:1rem; color:var(--cream-dim); font-style:italic; margin:0 0 1rem; }}
  footer {{ margin-top:3rem; padding-top:1.4rem; border-top:1px solid var(--line); text-align:center; font-size:0.78rem; letter-spacing:0.18em; text-transform:uppercase; color:var(--cream-dim); opacity:0.7; }}
  footer span {{ color: var(--brass); }}
{lang_css}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <a class="back" href="{self_index}">← {title}</a>
    <a class="lang" href="{other_page}" hreflang="{alt_hreflang}">{other_label}</a>
  </div>
  <h1>{page_title}</h1>
  <div class="meta">{ref}{note_html}</div>
  {idea_html}
  <div class="ms">
{body}
  </div>
  <footer>{page_footer}</footer>
</div>
</body>
</html>
"""


def conf():
    with open(CONF, encoding="utf-8") as f:
        return json.load(f)


def slug(title):
    s = title.lower().replace("'", "").replace("’", "")
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return (s[:60].rstrip("-") or "study")


def read_studies(lang):
    src = open(INDEX[lang], encoding="utf-8").read()
    m = re.search(r"const STUDIES = (\[.*?\]);\n", src, re.S)
    if not m:
        sys.exit("Could not find the STUDIES list in %s." % os.path.basename(INDEX[lang]))
    return src, json.loads(m.group(1)), m.span(1)


def save_studies(lang, src, span, studies):
    body = json.dumps(studies, indent=2, ensure_ascii=False)
    open(INDEX[lang], "w", encoding="utf-8").write(src[:span[0]] + body + src[span[1]:])


# --- docx ---------------------------------------------------------------

def from_docx(path):
    try:
        from docx import Document
    except ImportError:
        sys.exit("python-docx is needed to read .docx files:  pip3 install python-docx")
    doc = Document(path)
    title, idea, blocks = None, None, []
    pending, kind = [], None

    def flush():
        nonlocal pending, kind
        if pending:
            tag = "ol" if kind == "ol" else "ul"
            items = "\n".join("      <li>%s</li>" % html.escape(t) for t in pending)
            blocks.append("    <%s>\n%s\n    </%s>" % (tag, items, tag))
            pending, kind = [], None

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = (p.style.name or "").lower()
        if title is None and ("title" in style or style.startswith("heading 1")):
            title = text
            continue
        low = text.lower()
        if idea is None and (low.startswith(("big idea", "main idea", "aim:"))
                             or text.startswith(("中心思想", "主旨", "大意"))):
            idea = re.split(r"[:：]", text, 1)[-1].strip() or text
            continue
        if "heading" in style:
            flush(); blocks.append("    <h2>%s</h2>" % html.escape(text)); continue
        if "list number" in style or re.match(r"^\d+[.)、]\s*", text):
            if kind not in (None, "ol"):
                flush()
            kind = "ol"; pending.append(re.sub(r"^\d+[.)、]\s*", "", text)); continue
        if "list bullet" in style or text.startswith(("-", "•", "*")):
            if kind not in (None, "ul"):
                flush()
            kind = "ul"; pending.append(text.lstrip("-•* ").strip()); continue
        if "quote" in style:
            flush(); blocks.append("    <blockquote>%s</blockquote>" % html.escape(text)); continue
        flush(); blocks.append("    <p>%s</p>" % html.escape(text))
    flush()

    if title is None:
        title = os.path.splitext(os.path.basename(path))[0]
    return title, idea, "\n".join(blocks)


def write_page(c, lang, page, other_page, title, ref, note, idea, body):
    L = LANGS[lang]
    m = c[lang]
    with open(os.path.join(HERE, page), "w", encoding="utf-8") as f:
        f.write(PAGE.format(
            htmllang=L["htmllang"], prefix=c["prefix"], manifest=m["manifest"], fonts=FONTS,
            sans=L["sans"], serif=L["serif"], eyebrow=L["eyebrow"], title=m["title"],
            self_index=m["index"], other_page=other_page, other_label=L["other_label"],
            alt_hreflang=LANGS["cn" if lang == "en" else "en"]["htmllang"],
            page_title=html.escape(title), ref=html.escape(ref or ""),
            note_html=(" &middot; <span class='note'>%s</span>" % html.escape(note)) if note else "",
            idea_html=('<div class="bigidea"><b>%s</b>%s</div>' % (L["bigidea"], html.escape(idea))) if idea else "",
            body=body, base_font=c.get("base_font", "18px"), page_footer=L["page_footer"],
            lang_css=L["lang_css"],
        ))


def insert(studies, entry):
    idx = next((i for i, s in enumerate(studies) if s.get("group") == entry["group"]), len(studies))
    studies.insert(idx, entry)


def repoint(page, other_page):
    """Update an already-written page's language toggle to point at its counterpart."""
    path = os.path.join(HERE, page)
    if not os.path.exists(path):
        return
    src = open(path, encoding="utf-8").read()
    src = re.sub(r'(<link rel="alternate" hreflang="[^"]+" href=")[^"]*(")', r"\g<1>%s\g<2>" % other_page, src)
    src = re.sub(r'(<a class="lang" href=")[^"]*(")', r"\g<1>%s\g<2>" % other_page, src)
    open(path, "w", encoding="utf-8").write(src)


# --- commands -----------------------------------------------------------

def cmd_add(args):
    c = conf()

    # --- attach a Chinese page to an existing English entry
    if args.lang == "cn" and args.n:
        src_en, en_studies, _ = read_studies("en")
        target = next((s for s in en_studies if s.get("n") == args.n), None)
        if not target:
            sys.exit("No entry n=%d in index.html." % args.n)
        stem = re.sub(r"\.html$", "", target["page"])
        cn_page = stem + "-cn.html"
        title, idea, body = (from_docx(args.docx) if args.docx else (args.cn_title, args.cn_idea, ""))
        title = args.cn_title or title
        idea = args.cn_idea or idea
        write_page(c, "cn", cn_page, target["page"], title, args.cn_ref or target.get("ref", ""),
                   args.cn_note or target.get("note", ""), idea, body)
        src_cn, cn_studies, span_cn = read_studies("cn")
        insert(cn_studies, {"n": args.n, "title": title, "ref": args.cn_ref or target.get("ref", ""),
                            "group": args.cn_group or target.get("group", ""),
                            "note": args.cn_note or target.get("note", ""),
                            "idea": idea or "", "blurb": args.cn_blurb or "", "page": cn_page})
        save_studies("cn", src_cn, span_cn, cn_studies)
        repoint(target["page"], cn_page)
        print("wrote %s and linked it from %s" % (cn_page, target["page"]))
        return

    # --- normal add
    if args.docx:
        title, idea, body = from_docx(args.docx)
    else:
        title, idea, body = args.title, args.idea, "    <p>Add the text here.</p>"
    title = args.title or title
    idea = args.idea or idea
    if not title:
        sys.exit("Give a --title, or a .docx with a Title/Heading 1 paragraph.")

    src_en, en_studies, span_en = read_studies("en")
    n = max([s.get("n", 0) for s in en_studies] + [0]) + 1
    stem = "%03d-%s" % (n, slug(title))
    en_page, cn_page = stem + ".html", stem + "-cn.html"

    have_cn = bool(args.cn_docx or args.cn_title)
    write_page(c, "en", en_page, cn_page if have_cn else c["cn"]["index"],
               title, args.ref, args.note, idea, body)
    insert(en_studies, {"n": n, "title": title, "ref": args.ref or "", "group": args.group or "Recent",
                        "note": args.note or "", "idea": idea or "", "blurb": args.blurb or "",
                        "page": en_page})
    save_studies("en", src_en, span_en, en_studies)
    print("wrote %s" % en_page)

    if have_cn:
        if args.cn_docx:
            cn_title, cn_idea, cn_body = from_docx(args.cn_docx)
        else:
            cn_title, cn_idea, cn_body = args.cn_title, args.cn_idea, "    <p>请在此加入内容。</p>"
        cn_title = args.cn_title or cn_title
        cn_idea = args.cn_idea or cn_idea
        write_page(c, "cn", cn_page, en_page, cn_title, args.cn_ref or args.ref,
                   args.cn_note or args.note, cn_idea, cn_body)
        src_cn, cn_studies, span_cn = read_studies("cn")
        insert(cn_studies, {"n": n, "title": cn_title, "ref": args.cn_ref or args.ref or "",
                            "group": args.cn_group or args.group or "Recent",
                            "note": args.cn_note or args.note or "", "idea": cn_idea or "",
                            "blurb": args.cn_blurb or "", "page": cn_page})
        save_studies("cn", src_cn, span_cn, cn_studies)
        print("wrote %s" % cn_page)
    else:
        print("no Chinese version yet — add it later with:  "
              "python3 add-study.py 中文.docx --lang cn --n %d" % n)

    print("now: git add -A && git commit -m %r && git push" % ("Add: " + title))


def cmd_check(args):
    ok = True
    all_pages = {f for f in os.listdir(HERE) if re.match(r"^\d{3}-.*\.html$", f)}
    listed = {}
    for lang in ("en", "cn"):
        _, studies, _ = read_studies(lang)
        listed[lang] = {s["page"] for s in studies if s.get("page")}
        ns = [s.get("n") for s in studies]
        if len(set(ns)) != len(ns):
            print("DUPLICATE n values in %s" % os.path.basename(INDEX[lang])); ok = False
        for s in studies:
            if not s.get("title") or not s.get("group"):
                print("INCOMPLETE  %s entry n=%s" % (lang, s.get("n"))); ok = False
        for p in sorted(listed[lang] - all_pages):
            print("MISSING     %s (listed in %s)" % (p, os.path.basename(INDEX[lang]))); ok = False

    for p in sorted(all_pages - (listed["en"] | listed["cn"])):
        print("ORPHAN      %s (no entry in either index)" % p); ok = False

    # every Chinese page must belong to index-cn, every English page to index
    for p in sorted(listed["cn"]):
        if not p.endswith("-cn.html"):
            print("MISNAMED    %s listed in index-cn.html but has no -cn suffix" % p); ok = False
    for p in sorted(listed["en"]):
        if p.endswith("-cn.html"):
            print("MISNAMED    %s listed in index.html but is a Chinese page" % p); ok = False

    # language toggles must resolve
    for p in sorted(all_pages):
        src = open(os.path.join(HERE, p), encoding="utf-8").read()
        m = re.search(r'<a class="lang" href="([^"]+)"', src)
        if m and m.group(1) not in all_pages and not m.group(1).startswith("index"):
            print("BROKEN LINK %s → %s" % (p, m.group(1))); ok = False

    n_en, n_cn = len(listed["en"]), len(listed["cn"])
    missing_cn = sorted(p for p in listed["en"] if p.replace(".html", "-cn.html") not in listed["cn"])
    if missing_cn:
        print("NO CHINESE YET (not an error):")
        for p in missing_cn:
            print("            %s" % p)

    print(("OK — %d English, %d Chinese, %d page files, all matched" % (n_en, n_cn, len(all_pages)))
          if ok else "problems found")
    sys.exit(0 if ok else 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="?", help="Word document to convert")
    ap.add_argument("--lang", choices=["en", "cn"], default="en",
                    help="with --n, attach a Chinese page to an existing entry")
    ap.add_argument("--n", type=int, help="existing entry number to attach a Chinese page to")
    ap.add_argument("--title"); ap.add_argument("--ref", help='e.g. "Mark 10:35-45"')
    ap.add_argument("--group", help='e.g. "Term 1 2026"')
    ap.add_argument("--note", help='e.g. "February 2026"')
    ap.add_argument("--idea", help="the one-sentence big idea")
    ap.add_argument("--blurb", help="two sentences shown on the browse page")
    ap.add_argument("--cn-docx", help="Chinese Word document")
    ap.add_argument("--cn-title"); ap.add_argument("--cn-ref"); ap.add_argument("--cn-group")
    ap.add_argument("--cn-note"); ap.add_argument("--cn-idea"); ap.add_argument("--cn-blurb")
    ap.add_argument("--check", action="store_true", help="verify entries, pages and toggles match")
    args = ap.parse_args()
    (cmd_check if args.check else cmd_add)(args)


if __name__ == "__main__":
    main()
