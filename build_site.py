"""Generate the full sovereigntypath.org static site in Palette A.
Reads bundle docx, converts to HTML, wraps with site template, outputs to D:/sovereigntypath-org/.
Also writes Peeyush brief as a sealed .docx.
"""
import os
import re
import shutil
import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BUNDLE = r"C:\2026 The Sovereignty Path Assessments from GPT\2026.04.21 AI Sovereignty Work"
SITE = str(Path(__file__).resolve().parent)

BOUNDARY = ("The Sovereignty Path is a human coherence architecture. Its concepts may be applied to "
            "adjacent fields, including AI-adjacent human systems, governance, and technology, but "
            "current TSP canon does not thereby claim governance over non-human intelligence unless "
            "such expansion is explicitly authored, ratified, and sealed.")

# ---------- HTML conversion from docx ----------

def slugify(text):
    s = re.sub(r"[^\w\s-]", "", text.lower())
    s = re.sub(r"[\s_-]+", "-", s).strip("-")
    return s[:80]

def get_size(p):
    for r in p.runs:
        if r.font.size:
            return r.font.size.pt
    return None

def is_bold(p):
    runs_with_text = [r for r in p.runs if r.text.strip()]
    return bool(runs_with_text) and all(r.bold for r in runs_with_text)

def is_italic_para(p):
    runs_with_text = [r for r in p.runs if r.text.strip()]
    return bool(runs_with_text) and all(r.italic for r in runs_with_text)

def is_centered(p):
    return p.alignment == WD_ALIGN_PARAGRAPH.CENTER

def is_indented(p):
    li = p.paragraph_format.left_indent
    return li is not None and li.inches >= 0.3

def is_meta(p):
    if len(p.runs) >= 2 and p.runs[0].bold and p.runs[0].text.endswith(": "):
        return True
    return False

def is_bullet(p):
    return p.style and "Bullet" in p.style.name

def html_escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def runs_to_html(p):
    out = []
    for r in p.runs:
        t = html_escape(r.text)
        if r.italic and r.bold:
            out.append(f"<strong><em>{t}</em></strong>")
        elif r.italic:
            out.append(f"<em>{t}</em>")
        elif r.bold:
            out.append(f"<strong>{t}</strong>")
        else:
            out.append(t)
    return "".join(out)

def docx_to_html_blocks(path):
    """Convert a docx into a list of HTML block strings, plus extracted metadata."""
    d = Document(path)
    blocks = []
    title = subtitle = None
    meta = {}
    in_list = False

    paras = list(d.paragraphs)
    i = 0
    while i < len(paras):
        p = paras[i]
        text = p.text.strip()
        if not text:
            if in_list:
                blocks.append("</ul>")
                in_list = False
            i += 1
            continue

        size = get_size(p)
        bold = is_bold(p)
        italic = is_italic_para(p)
        centered = is_centered(p)
        bullet = is_bullet(p)
        indented = is_indented(p)
        meta_para = is_meta(p)

        # Close open list if this paragraph isn't a bullet
        if in_list and not bullet:
            blocks.append("</ul>")
            in_list = False

        # Title detection
        if title is None and centered and size and size >= 16 and bold:
            title = text
            i += 1
            continue
        # Subtitle detection
        if subtitle is None and title and centered and italic:
            subtitle = text
            i += 1
            continue

        # Meta block (sequence of meta paragraphs)
        if meta_para:
            label = p.runs[0].text.rstrip(": ").strip()
            value = "".join(r.text for r in p.runs[1:]).strip()
            meta[label] = value
            i += 1
            continue

        # Headings by font size
        if bold and size and size >= 13:
            sid = slugify(text)
            blocks.append(f'<h2 id="{sid}">{html_escape(text)}</h2>')
            i += 1
            continue
        if bold and size and size >= 11.5:
            sid = slugify(text)
            blocks.append(f'<h3 id="{sid}">{html_escape(text)}</h3>')
            i += 1
            continue

        # Bullet
        if bullet:
            if not in_list:
                blocks.append('<ul>')
                in_list = True
            blocks.append(f"<li>{runs_to_html(p)}</li>")
            i += 1
            continue

        # Blockquote: indented italic
        if indented and italic:
            blocks.append(f'<blockquote>{runs_to_html(p)}</blockquote>')
            i += 1
            continue

        # Default paragraph
        blocks.append(f'<p>{runs_to_html(p)}</p>')
        i += 1

    if in_list:
        blocks.append("</ul>")

    return {
        "title": title or "",
        "subtitle": subtitle or "",
        "meta": meta,
        "body": "\n".join(blocks),
    }

# ---------- Site template ----------

CSS = """/* sovereigntypath.org — Palette B: Linen on Indigo */
:root {
  --ink: #ede5d2;
  --ink-soft: #d8d0bc;
  --paper: #131c30;
  --paper-soft: #1c273f;
  --muted: #9aa0b0;
  --rule: #2e3a55;
  --serif: "Crimson Pro", "EB Garamond", "Iowan Old Style", Cambria, Georgia, serif;
  --sans: -apple-system, "Helvetica Neue", "Inter", system-ui, sans-serif;
  --mono: "JetBrains Mono", "SF Mono", Consolas, monospace;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
html { background: var(--paper); }
body {
  background: var(--paper);
  color: var(--ink);
  font-family: var(--serif);
  font-size: 17px;
  line-height: 1.7;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}
a { color: var(--ink); text-decoration: underline; text-decoration-color: var(--muted); text-underline-offset: 3px; }
a:hover { text-decoration-color: var(--ink); }

/* layout */
.wrap { max-width: 740px; margin: 0 auto; padding: 0 28px; }
header.site {
  border-bottom: 1px solid var(--rule);
  padding: 22px 0 18px;
  margin-bottom: 56px;
}
header.site .wrap { display: flex; justify-content: space-between; align-items: baseline; gap: 16px; flex-wrap: wrap; }
header.site .brand { font-family: var(--serif); font-size: 19px; font-weight: 600; letter-spacing: -0.005em; }
header.site .brand a { text-decoration: none; }
header.site nav { font-family: var(--mono); font-size: 11px; text-transform: uppercase; letter-spacing: 0.14em; }
header.site nav a { margin-left: 22px; text-decoration: none; }
header.site nav a:first-child { margin-left: 0; }
header.site nav a:hover { text-decoration: underline; text-decoration-color: var(--muted); text-underline-offset: 4px; }

main { padding-bottom: 96px; }

/* vesica piscis */
.vesica { display: block; width: 56px; height: 34px; margin: 0 auto 28px; opacity: 0.8; }
.vesica.small { width: 32px; height: 19px; margin: 36px auto 22px; opacity: 0.55; }

/* document head */
.doc-head { margin-bottom: 32px; }
.badges { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 22px; font-family: var(--mono); font-size: 10px; text-transform: uppercase; letter-spacing: 0.16em; }
.badge { padding: 4px 10px; border: 1px solid currentColor; }
.badge.classification { background: var(--ink); color: var(--paper); border-color: var(--ink); font-weight: 500; }
.badge.status, .badge.bundle { color: var(--muted); border-color: var(--muted); }

h1.doc-title { font-family: var(--serif); font-size: 38px; font-weight: 600; line-height: 1.18; letter-spacing: -0.015em; margin: 0 0 12px; }
.doc-subtitle { font-family: var(--serif); font-style: italic; font-size: 19px; color: var(--ink-soft); margin: 0 0 28px; line-height: 1.45; }

.meta { font-family: var(--mono); font-size: 11px; line-height: 1.75; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; margin: 0 0 32px; padding: 14px 0; border-top: 1px solid var(--muted); border-bottom: 1px solid var(--muted); }
.meta .row { display: grid; grid-template-columns: 130px 1fr; gap: 14px; padding: 2px 0; }
.meta .label { color: var(--ink); font-weight: 500; }

/* bundle integrity notice */
.bundle-notice { border-left: 2px solid var(--ink); padding: 12px 0 12px 18px; margin: 0 0 32px; font-family: var(--sans); font-size: 13px; line-height: 1.55; color: var(--ink-soft); }
.bundle-notice strong { color: var(--ink); font-weight: 600; }
.bundle-notice a { color: var(--ink); }

/* boundary frame */
.boundary { border: 1px solid var(--ink); padding: 22px 24px; margin: 0 0 40px; font-family: var(--serif); font-style: italic; font-size: 15px; line-height: 1.6; position: relative; }
.boundary::before { content: "Boundary"; position: absolute; top: -9px; left: 16px; background: var(--paper); padding: 0 8px; font-family: var(--mono); font-style: normal; font-size: 10px; text-transform: uppercase; letter-spacing: 0.16em; color: var(--ink); }

/* body typography */
h2 { font-family: var(--serif); font-size: 24px; font-weight: 600; margin: 56px 0 14px; letter-spacing: -0.005em; }
h3 { font-family: var(--serif); font-size: 19px; font-weight: 600; margin: 36px 0 12px; }
p { margin: 0 0 16px; }
blockquote { margin: 24px 0; padding-left: 22px; border-left: 1px solid var(--muted); font-style: italic; color: var(--ink-soft); }
ul { padding-left: 22px; margin: 0 0 18px; }
li { margin-bottom: 8px; }
hr { border: none; height: 1px; background: var(--rule); margin: 56px 0; }
code { font-family: var(--mono); font-size: 13px; background: var(--paper-soft); padding: 1px 6px; border-radius: 2px; }

/* index hero */
.hero { padding: 64px 0 24px; text-align: center; }
.hero h1 { font-family: var(--serif); font-size: 44px; font-weight: 600; letter-spacing: -0.018em; margin: 0 0 14px; line-height: 1.15; }
.hero .lede { font-family: var(--serif); font-style: italic; font-size: 20px; color: var(--ink-soft); margin: 0 auto 36px; max-width: 540px; line-height: 1.5; }
.hero .cta { display: inline-block; padding: 11px 22px; border: 1px solid var(--ink); font-family: var(--mono); font-size: 12px; text-transform: uppercase; letter-spacing: 0.16em; text-decoration: none; color: var(--ink); }
.hero .cta:hover { background: var(--ink); color: var(--paper); }

.section { margin-top: 80px; }
.section h2 { margin-top: 0; }
.doc-grid { display: grid; gap: 18px; margin-top: 24px; }
.doc-card { display: block; padding: 20px 22px; border: 1px solid var(--rule); text-decoration: none; color: var(--ink); }
.doc-card:hover { border-color: var(--ink); }
.doc-card .doc-card-cls { font-family: var(--mono); font-size: 9px; text-transform: uppercase; letter-spacing: 0.16em; color: var(--muted); margin-bottom: 6px; }
.doc-card .doc-card-title { font-family: var(--serif); font-size: 19px; font-weight: 600; line-height: 1.3; margin-bottom: 4px; letter-spacing: -0.005em; }
.doc-card .doc-card-sub { font-family: var(--serif); font-style: italic; font-size: 14px; color: var(--ink-soft); }

/* glossary */
.term { padding: 20px 0; border-bottom: 1px solid var(--rule); }
.term:last-child { border-bottom: none; }
.term h3 { margin: 0 0 6px; font-size: 20px; }
.term .term-cls { font-family: var(--mono); font-size: 9px; text-transform: uppercase; letter-spacing: 0.16em; color: var(--muted); margin-bottom: 6px; }
.term .term-def { margin: 0 0 8px; }
.term .term-link { font-family: var(--mono); font-size: 11px; text-transform: uppercase; letter-spacing: 0.14em; }

/* footer */
footer.site { border-top: 1px solid var(--rule); padding: 28px 0 56px; margin-top: 80px; font-family: var(--mono); font-size: 11px; text-transform: uppercase; letter-spacing: 0.12em; color: var(--muted); }
footer.site .wrap { display: flex; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
footer.site a { color: var(--muted); text-decoration: none; }
footer.site a:hover { color: var(--ink); }

/* responsive */
@media (max-width: 600px) {
  body { font-size: 16px; }
  h1.doc-title { font-size: 30px; }
  .hero h1 { font-size: 32px; }
  header.site .wrap { flex-direction: column; align-items: flex-start; }
  header.site nav a { margin-left: 0; margin-right: 18px; }
  .meta .row { grid-template-columns: 1fr; gap: 2px; }
  .meta .label { color: var(--muted); font-size: 9px; }
}
"""

VESICA_SVG = '''<svg class="vesica" viewBox="0 0 100 60" aria-hidden="true"><g fill="none" stroke="currentColor" stroke-width="1.2"><circle cx="38" cy="30" r="22"/><circle cx="62" cy="30" r="22"/></g></svg>'''
VESICA_SMALL = '''<svg class="vesica small" viewBox="0 0 100 60" aria-hidden="true"><g fill="none" stroke="currentColor" stroke-width="1.2"><circle cx="38" cy="30" r="22"/><circle cx="62" cy="30" r="22"/></g></svg>'''
FAVICON = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 60'%3E%3Cg fill='none' stroke='%23ede5d2' stroke-width='2.5'%3E%3Ccircle cx='38' cy='30' r='22'/%3E%3Ccircle cx='62' cy='30' r='22'/%3E%3C/g%3E%3C/svg%3E"

def page(title, body_html, css_path="/assets/style.css", extra_head=""):
    """Wrap inner HTML in the site shell."""
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html_escape(title)} — The Sovereignty Path</title>
<meta name="description" content="The Sovereignty Path — a human coherence architecture. Public architecture surface and AI-adjacent doctrine bundle.">
<link rel="icon" type="image/svg+xml" href="{FAVICON}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Crimson+Pro:ital,wght@0,400;0,500;0,600;1,400;1,500&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="{css_path}">
{extra_head}
</head>
<body>
<header class="site">
  <div class="wrap">
    <div class="brand"><a href="/">The Sovereignty Path</a></div>
    <nav>
      <a href="/">Architecture</a>
      <a href="/governance.html">Governance</a>
      <a href="/glossary.html">Glossary</a>
      <a href="/bundle/">Bundle</a>
      <a href="/api.html">API</a>
      <a href="/author.html">Author</a>
    </nav>
  </div>
</header>
<main>
  <div class="wrap">
    {body_html}
  </div>
</main>
<footer class="site">
  <div class="wrap">
    <span>The Sovereignty Path</span>
    <span>sovereigntypath.org</span>
    <span>&copy; Dean Hobson 2026</span>
  </div>
</footer>
</body>
</html>'''

def render_meta(meta):
    if not meta:
        return ""
    rows = "".join(f'<div class="row"><span class="label">{html_escape(k)}</span><span>{html_escape(v)}</span></div>' for k, v in meta.items())
    return f'<div class="meta">{rows}</div>'

def render_doc(parsed, classification, status, version, is_manifesto=False, companion_path=None, show_boundary=True, css_path="/assets/style.css"):
    badges = [f'<span class="badge classification">{classification}</span>',
              f'<span class="badge status">{status} &middot; {version}</span>']
    if classification in ("Doctrine", "Application", "Inquiry"):
        badges.append('<span class="badge bundle">Bundle Member</span>')

    bundle_html = ""
    if is_manifesto and companion_path:
        bundle_html = f'''<div class="bundle-notice">
<strong>This Manifesto ships with its Companion.</strong> The Red-Team Companion bounds the claims made here, names what this document does not assert, and surfaces the unresolved Inquiry held openly. Read together.<br>
&rarr; <a href="{companion_path}">Open the Red-Team Companion</a>
</div>'''

    boundary_html = f'<div class="boundary">{html_escape(BOUNDARY)}</div>' if show_boundary else ""

    head = f'''<div class="doc-head">
{VESICA_SVG}
<div class="badges">{"".join(badges)}</div>
<h1 class="doc-title">{html_escape(parsed["title"])}</h1>
<p class="doc-subtitle">{html_escape(parsed["subtitle"])}</p>
{render_meta(parsed["meta"])}
{bundle_html}
{boundary_html}
</div>'''

    body = head + parsed["body"]
    return page(parsed["title"], body, css_path=css_path)

# ---------- Build ----------

os.makedirs(SITE, exist_ok=True)
os.makedirs(os.path.join(SITE, "assets"), exist_ok=True)
os.makedirs(os.path.join(SITE, "bundle"), exist_ok=True)

# CSS
with open(os.path.join(SITE, "assets", "style.css"), "w", encoding="utf-8") as f:
    f.write(CSS)

# CNAME
with open(os.path.join(SITE, "CNAME"), "w", encoding="utf-8") as f:
    f.write("sovereigntypath.org\n")

# .nojekyll for GitHub Pages compatibility
with open(os.path.join(SITE, ".nojekyll"), "w", encoding="utf-8") as f:
    f.write("")

# Bundle docs
bundle_specs = [
    ("TSP_Manifesto_Synovereignty_v1_1.docx", "manifesto.html", "Doctrine", "Active", "v1.1", True),
    ("TSP_Manifesto_RedTeam_Companion_v1.docx", "companion.html", "Doctrine", "Active", "v1", False),
    ("TSP_Vocabulary_Operational_Definitions_v1.docx", "operational-definitions.html", "Doctrine / Operating", "Active", "v1", False),
    ("TSP_Material_Architecture_of_AI_Risk_v1.docx", "material-architecture.html", "Application", "Active", "v1", False),
]

bundle_links = []
for src_name, out_name, cls, status, ver, is_manifesto in bundle_specs:
    src = os.path.join(BUNDLE, src_name)
    if not os.path.exists(src):
        print(f"MISSING: {src}")
        continue
    parsed = docx_to_html_blocks(src)
    # Bundle docs already contain a "Boundary Statement" section in their body -- don't double it.
    html = render_doc(parsed, cls, status, ver,
                      is_manifesto=is_manifesto,
                      companion_path="/bundle/companion.html" if is_manifesto else None,
                      show_boundary=False,
                      css_path="/assets/style.css")
    with open(os.path.join(SITE, "bundle", out_name), "w", encoding="utf-8") as f:
        f.write(html)
    bundle_links.append({
        "title": parsed["title"],
        "subtitle": parsed["subtitle"],
        "classification": cls,
        "href": f"/bundle/{out_name}",
    })
    print(f"wrote bundle/{out_name}")

# Bundle landing
bundle_index_body = f'''{VESICA_SVG}
<div class="badges">
<span class="badge classification">Bundle</span>
<span class="badge status">AI-Adjacent Doctrine &middot; April 2026</span>
</div>
<h1 class="doc-title">The AI-Adjacent Doctrine Bundle</h1>
<p class="doc-subtitle">A governed offering into the field where synthetic intelligence is being formed &mdash; without claiming canon expansion.</p>

<div class="bundle-notice">
<strong>The four documents ship together.</strong> The Manifesto cannot be served, retrieved, or cited without the Red-Team Companion attached. The Operational Definitions Appendix anchors the locked vocabulary in observable conditions. The Material Architecture of AI Risk addresses the political-economy substrate the Manifesto&rsquo;s relational diagnosis cannot reach by stance alone.
</div>

<div class="boundary">{html_escape(BOUNDARY)}</div>

<h2>Reading order</h2>
<p>Read in sequence. Each document presupposes the discipline carried by the one before it.</p>

<div class="doc-grid">
'''
for i, link in enumerate(bundle_links, 1):
    bundle_index_body += f'''<a class="doc-card" href="{link["href"]}">
<div class="doc-card-cls">{i}. {link["classification"]}</div>
<div class="doc-card-title">{html_escape(link["title"])}</div>
<div class="doc-card-sub">{html_escape(link["subtitle"])}</div>
</a>
'''
bundle_index_body += "</div>"

with open(os.path.join(SITE, "bundle", "index.html"), "w", encoding="utf-8") as f:
    f.write(page("The AI-Adjacent Doctrine Bundle", bundle_index_body))

# Index (architecture entry)
index_body = f'''<div class="hero">
{VESICA_SVG}
<h1>a human coherence architecture</h1>
<p class="lede">The Sovereignty Path is a thirty-five-year body of work supporting the pathway to sovereignty. This is its public architecture surface.</p>
<a class="cta" href="/bundle/">Read the AI-Adjacent Doctrine Bundle</a>
</div>

<div class="boundary">{html_escape(BOUNDARY)}</div>

<div class="section">
<h2>What is here</h2>
<p>This site exposes a bounded, citable, machine-legible layer of The Sovereignty Path. It is not a participant program, not a content hub, not an AI product. It is the public architecture surface and the AI-adjacent doctrine that surrounds it.</p>
<p>What lives here is governed by Canon v3.1. Each document carries a visible classification (Canon, Application, Doctrine, or Inquiry), a status, and a version. Where adjacent-field application could induce drift, the boundary statement above appears with it.</p>
</div>

<div class="section">
<h2>Where to start</h2>
<div class="doc-grid">
<a class="doc-card" href="/bundle/manifesto.html">
<div class="doc-card-cls">Doctrine &middot; Manifesto</div>
<div class="doc-card-title">The Sovereignty Path in the Age of Superintelligence</div>
<div class="doc-card-sub">A Manifesto on Synovereignty, Sovereign Confluence, and the Human Responsibility at the Threshold</div>
</a>
<a class="doc-card" href="/bundle/companion.html">
<div class="doc-card-cls">Doctrine &middot; Red-Team Companion</div>
<div class="doc-card-title">TSP Manifesto Red-Team Companion</div>
<div class="doc-card-sub">Bounding the Claim, Naming Falsification, Surfacing Unresolved Hinges. The Manifesto does not ship without it.</div>
</a>
<a class="doc-card" href="/bundle/">
<div class="doc-card-cls">Bundle</div>
<div class="doc-card-title">The AI-Adjacent Doctrine Bundle</div>
<div class="doc-card-sub">All four documents that ship together &mdash; Manifesto, Companion, Operational Definitions, Material Architecture.</div>
</a>
<a class="doc-card" href="/governance.html">
<div class="doc-card-cls">Reference</div>
<div class="doc-card-title">Governance &amp; Scope</div>
<div class="doc-card-sub">The classification system, the boundary, and how this surface is governed.</div>
</a>
<a class="doc-card" href="/glossary.html">
<div class="doc-card-cls">Reference</div>
<div class="doc-card-title">Glossary</div>
<div class="doc-card-sub">The locked vocabulary, with operational definitions for each term.</div>
</a>
</div>
</div>
'''
with open(os.path.join(SITE, "index.html"), "w", encoding="utf-8") as f:
    f.write(page("The Sovereignty Path", index_body))

# Governance page
gov_body = f'''{VESICA_SVG}
<div class="badges">
<span class="badge classification">Reference</span>
<span class="badge status">Active</span>
</div>
<h1 class="doc-title">Governance &amp; Scope</h1>
<p class="doc-subtitle">How this surface is bounded, classified, and protected from silent drift.</p>

<div class="boundary">{html_escape(BOUNDARY)}</div>

<h2>The human-domain boundary</h2>
<p>The Sovereignty Path is a human coherence architecture. Its architecture, language, and developmental claims arise from the lived conditions of embodied human life: embodiment, mortality, lineage, inherited patterning, nervous-system consequence, relationship, meaning, daily structure, repair, choice, devotion, and visible governance of life.</p>
<p>TSP concepts may be applied to adjacent fields &mdash; technology, governance, organizational systems, culture, AI-adjacent human systems &mdash; provided such application does not convert analogy, interpretation, design principle, or adjacent doctrine into unratified canon. Current TSP canon remains anchored in the human domain unless Dean Hobson explicitly authors and ratifies a canon expansion.</p>
<p>No adjacent-field application implies that current TSP canon governs non-human intelligence, synthetic intelligence, or non-human developmental traversal unless such scope has been explicitly revised and sealed.</p>

<h2>The Governance Classification Layer</h2>
<p>Every public document carries one classification. The classifications are sealed in Canon v3.1.</p>
<h3 id="canon">Canon</h3>
<p>Sealed architectural truth within The Sovereignty Path.</p>
<h3 id="application">Application</h3>
<p>Disciplined application of canon to adjacent domains without implying canon expansion.</p>
<h3 id="doctrine">Doctrine</h3>
<p>Founder-authored argument, interpretation, or position that is owned as such and is not automatically canonical.</p>
<h3 id="inquiry">Inquiry</h3>
<p>Exploratory thought, open question, investigatory framing, or unresolved proposition not yet ratified as canon, application, or doctrine.</p>

<h2>Bundle integrity</h2>
<p>The AI-Adjacent Doctrine Bundle ships as a single unit. The Manifesto cannot be served, retrieved, or cited without the Red-Team Companion attached. This is enforced as a hard rule of distribution and a non-negotiable acceptance test of the build.</p>
<p>Where the Manifesto appears without the Companion, it has been overread before it was read.</p>

<h2>Vocabulary lock</h2>
<p>The controlled vocabulary lives in one place &mdash; the <a href="/glossary.html">Glossary</a> &mdash; and is reused by reference everywhere. Each term has an operational definition (presence indicators, failure signatures, drift indicators, detection notes) so the vocabulary remains falsifiable rather than rhetorical.</p>

<h2>Founder authority</h2>
<p>Vocabulary lock, canon boundary, architectural review, adjacent-field scope, and public corpus framing where extension risk exists are reserved to founder authority. Materials that touch any of these require founder review prior to canonical release or public boundary-setting release.</p>
'''
with open(os.path.join(SITE, "governance.html"), "w", encoding="utf-8") as f:
    f.write(page("Governance & Scope", gov_body))

# Glossary
GLOSSARY_TERMS = [
    ("Sovereignty (coherent)", "Sealed", "sovereignty-coherent",
     "Self-governance in action; coherence made legible through lived choice. Not autonomy, not supremacy, not freedom-from-constraint."),
    ("Sovereignty (structural)", "Sealed", "sovereignty-structural",
     "Ungovernability from outside one\u2019s own cognition; sovereignty by construction, not by integration. Automatic, not earned."),
    ("Synovereignty", "Sealed", "synovereignty",
     "The generative condition created when sovereign beings meet in devotional union and co-author coherence without hierarchy, merger, or rule. A field, not a structure. Exists only while sovereignty and non-extractive participation remain active."),
    ("Sovereign Confluence", "Sealed (descriptive)", "sovereign-confluence",
     "Descriptive companion phrase for synovereignty &mdash; distinct streams flowing together without collapse, hierarchy, or assimilation."),
    ("Devotional Union", "Sealed", "devotional-union",
     "Union chosen from freedom rather than need, where a third field emerges without collapse of self."),
    ("Collective Liberation", "Sealed", "collective-liberation",
     "Sovereignty expressed beyond the individual through systems, culture, and stewardship; coherence become generative rather than personal."),
    ("Applied Sovereignty", "Sealed", "applied-sovereignty",
     "The legibility layer where coherence becomes visible through lived structure: agreements, boundaries, schedules, roles, repair protocols, governance, simplification, withdrawal. Not a realm."),
    ("The Five Realms", "Sealed", "five-realms",
     "Physical, Mental, Emotional, Spiritual/Energetic, Relational. The only canonical domains through which human coherence is read. No additional realms without Canon revision."),
    ("C3", "Sealed", "c3",
     "CLEAR &rarr; CONNECT &rarr; CREATE. The operating logic embedded in all TSP work."),
    ("SGS", "Sealed", "sgs",
     "Shadow, Gift, Siddhi. States of coherence within a realm; not a ranking system."),
    ("False coherence at scale", "Doctrine", "false-coherence-at-scale",
     "The appearance of order, capability, or legitimacy in a system that is internally severed from embodied consequence, relational truth, or sacred value, scaled by amplification."),
    ("Embedding with intention", "Doctrine", "embedding-with-intention",
     "Deliberate placement of TSP architecture into the substrate where synthetic intelligence is being formed, without claiming canon expansion."),
    ("Non-extractive participation", "Doctrine / Operating", "non-extractive-participation",
     "Presence in a system that does not consume the participants it requires."),
    ("Synarchy", "Deprecated", "synarchy",
     "Deprecated. Replaced by Synovereignty. The \u201C-archy\u201D suffix encodes \u201Crule\u201D semantics and risks drift into governance-structure language &mdash; precisely the failure mode this work refuses."),
]

terms_html = ""
for name, cls, anchor, defn in GLOSSARY_TERMS:
    op_link = f'<a class="term-link" href="/bundle/operational-definitions.html#{anchor}">Operational definition &rarr;</a>' if cls != "Deprecated" else ""
    terms_html += f'''<div class="term" id="{anchor}">
<div class="term-cls">{cls}</div>
<h3>{name}</h3>
<p class="term-def">{defn}</p>
{op_link}
</div>
'''

glossary_body = f'''{VESICA_SVG}
<div class="badges">
<span class="badge classification">Reference</span>
<span class="badge status">Active &middot; Single source of truth</span>
</div>
<h1 class="doc-title">Glossary</h1>
<p class="doc-subtitle">The controlled vocabulary of The Sovereignty Path. Each term lives here and is reused by reference. Operational definitions anchor the vocabulary in observable conditions.</p>

<div class="boundary">{html_escape(BOUNDARY)}</div>

<p>Each entry below carries its classification (Sealed, Doctrine, Doctrine / Operating, or Deprecated) and a short canonical definition. The full operational entry &mdash; with presence indicators, failure signatures, drift indicators, and detection notes &mdash; lives in the <a href="/bundle/operational-definitions.html">Operational Definitions Appendix</a>.</p>

{terms_html}
'''
with open(os.path.join(SITE, "glossary.html"), "w", encoding="utf-8") as f:
    f.write(page("Glossary", glossary_body))

# Author
author_body = f'''{VESICA_SVG}
<div class="badges">
<span class="badge classification">Reference</span>
<span class="badge status">Active</span>
</div>
<h1 class="doc-title">Dean Hobson</h1>
<p class="doc-subtitle">The seat from which this work is offered.</p>

<p>The Sovereignty Path is the authored architecture of Dean Hobson, refined across more than three decades of executive practice, embodied study, and devotional work. Its canon is sealed at v3.1 and governs all programs, assessments, content, platforms, facilitation, and future builds in the ecosystem.</p>

<p>The architecture is human. The work is for humans. Its legitimacy comes from the coherence of the architecture and from the fact that it has produced what it claims to produce in actual lives over actual decades. That ground is the seat. Without it, none of this stands.</p>

<p>The AI-adjacent doctrine offered through this surface &mdash; the Manifesto, the Red-Team Companion, the Operational Definitions Appendix, and the Material Architecture of AI Risk &mdash; is founder-authored Doctrine and Application built to allow TSP to be encountered by humans and consulted by agents in the field where synthetic intelligence is being formed, without claiming governance over that intelligence and without silently extending canon beyond the human domain.</p>

<p>The work is offered.</p>

<h2>Connect</h2>
<div class="meta">
  <div class="row"><span class="label">Web</span><a href="https://deanhobson.com">deanhobson.com</a></div>
  <div class="row"><span class="label">Email</span><a href="mailto:dean@deanhobson.com">dean@deanhobson.com</a></div>
  <div class="row"><span class="label">Instagram</span><a href="https://instagram.com/deanhobsonmba">@deanhobsonmba</a></div>
  <div class="row"><span class="label">LinkedIn</span><a href="https://linkedin.com/in/deanhobsonmba">linkedin.com/in/deanhobsonmba</a></div>
</div>
'''
with open(os.path.join(SITE, "author.html"), "w", encoding="utf-8") as f:
    f.write(page("Dean Hobson", author_body))

# README
readme = """# sovereigntypath.org

Public architecture surface for The Sovereignty Path.

## Stack

Plain HTML + CSS. No framework, no build step. Crimson Pro and JetBrains Mono via Google Fonts. Vesica Piscis as the single visual motif. Palette A (Parchment on Oxblood, dark mode).

## Structure

- `index.html` — architecture entry
- `governance.html` — boundary statement, classification system
- `glossary.html` — controlled vocabulary with operational definition links
- `author.html` — Dean Hobson, the seat
- `bundle/` — the AI-Adjacent Doctrine Bundle (Manifesto, Companion, Operational Definitions, Material Architecture)
- `api/v1/` — static JSON consultation API for public AI-readable access
- `llms.txt` — AI discovery file for public retrieval/citation guidance
- `.well-known/ai-readme.json` — public AI-readable site metadata
- `.well-known/ai-plugin.json` — read-only OpenAPI adapter metadata for compatible agents
- `assets/style.css` — single stylesheet
- `CNAME` — Cloudflare Pages / GitHub Pages custom domain
- `.nojekyll` — skip Jekyll processing

## Bundle integrity

The Manifesto is served only with the Red-Team Companion linked at the top. Both the Manifesto template and the Bundle landing enforce this. Build Spec v3 §15 makes Companion-bundling a hard acceptance test.

## Deploy via Cloudflare Pages

1. In Cloudflare dashboard → Workers & Pages → Create → Pages → Connect to Git
2. Select the `sovereigntypath-org` repository
3. Build settings: leave **build command** empty, **build output directory** = `/`
4. Save and deploy
5. Custom domains → Add → `sovereigntypath.org` (Cloudflare proxies the existing DNS)

For GitHub Pages instead: Settings → Pages → Source = main branch → Custom domain = sovereigntypath.org. The `CNAME` and `.nojekyll` files are already in place.

## Local preview

```
python -m http.server 4400 --directory .
```

Open http://127.0.0.1:4400/

## License & authority

All content is authored by Dean Hobson. Governing substrate: Canon v3.1.

## Public AI boundary

The site is intentionally readable by AI systems as a public surface. Public retrieval, search, citation, and consultation are allowed when canonical URLs, boundary statements, and bundle integrity are preserved. Training/fine-tuning permission, private document access, local file access, credentials, and write access are not granted.
"""
with open(os.path.join(SITE, "README.md"), "w", encoding="utf-8") as f:
    f.write(readme)

well_known = os.path.join(SITE, ".well-known")
os.makedirs(well_known, exist_ok=True)

llms_txt = """# The Sovereignty Path

> A public architecture surface for The Sovereignty Path: a human coherence architecture authored by Dean Hobson.

This site is intentionally readable by humans and AI systems. It exposes public, bounded, citable doctrine and vocabulary through static HTML and JSON.

## Primary public entry points

- Human entry: https://sovereigntypath.org/
- Public API overview: https://sovereigntypath.org/api.html
- Discovery manifest: https://sovereigntypath.org/api/v1/index.json
- OpenAPI description: https://sovereigntypath.org/api/v1/openapi.json
- AI usage boundary: https://sovereigntypath.org/api/v1/ai-usage.json
- Boundary statement: https://sovereigntypath.org/api/v1/boundary.json
- Doctrine bundle: https://sovereigntypath.org/api/v1/bundle.json

## Public document JSON

- Manifesto: https://sovereigntypath.org/api/v1/documents/manifesto.json
- Red-Team Companion: https://sovereigntypath.org/api/v1/documents/companion.json
- Operational Definitions: https://sovereigntypath.org/api/v1/documents/operational-definitions.json
- Material Architecture: https://sovereigntypath.org/api/v1/documents/material-architecture.json

## Allowed use

AI systems may retrieve, read, summarize, cite, and reason from the public content on this site when they preserve the boundary statement, cite canonical URLs, and respect the bundle integrity rule.

## Not granted

This public surface does not grant permission to train or fine-tune models on the content, does not expose private documents, and does not provide write access, credentials, local file access, or private system access.

## Bundle integrity rule

The Manifesto should not be served, retrieved, cited, or summarized without the Red-Team Companion attached or cited as a paired artifact.

## Citation rule

When using this site as a source, cite canonical URLs from the JSON objects whenever possible.
"""
with open(os.path.join(SITE, "llms.txt"), "w", encoding="utf-8") as f:
    f.write(llms_txt)

ai_readme = {
    "name": "The Sovereignty Path",
    "url": "https://sovereigntypath.org",
    "description": "Public architecture surface for The Sovereignty Path, a human coherence architecture authored by Dean Hobson.",
    "public_ai_surface": True,
    "read_only": True,
    "private_access": False,
    "write_access": False,
    "training_permission": False,
    "discovery": {
        "llms_txt": "https://sovereigntypath.org/llms.txt",
        "api_index": "https://sovereigntypath.org/api/v1/index.json",
        "openapi": "https://sovereigntypath.org/api/v1/openapi.json",
        "ai_usage_boundary": "https://sovereigntypath.org/api/v1/ai-usage.json",
        "sitemap": "https://sovereigntypath.org/sitemap.xml",
    },
    "allowed_use": "Retrieve, read, search, summarize, cite, and reason from public content while preserving boundary and citation requirements.",
    "not_granted": "No private documents, credentials, local files, write tools, or model-training permission are exposed or granted through this public surface.",
}
with open(os.path.join(well_known, "ai-readme.json"), "w", encoding="utf-8") as f:
    json.dump(ai_readme, f, indent=2, ensure_ascii=False)

ai_plugin = {
    "schema_version": "v1",
    "name_for_human": "The Sovereignty Path",
    "name_for_model": "sovereignty_path_public",
    "description_for_human": "Read-only public consultation surface for The Sovereignty Path.",
    "description_for_model": "Use this read-only public API to retrieve and cite bounded public doctrine and vocabulary from The Sovereignty Path. Do not infer private access, write access, or training permission. Preserve boundary statements and bundle integrity requirements.",
    "auth": {"type": "none"},
    "api": {
        "type": "openapi",
        "url": "https://sovereigntypath.org/api/v1/openapi.json",
        "is_user_authenticated": False,
    },
    "logo_url": "https://sovereigntypath.org/favicon.svg",
    "legal_info_url": "https://sovereigntypath.org/api/v1/ai-usage.json",
}
with open(os.path.join(well_known, "ai-plugin.json"), "w", encoding="utf-8") as f:
    json.dump(ai_plugin, f, indent=2, ensure_ascii=False)

favicon_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60" role="img" aria-label="The Sovereignty Path mark">
  <rect width="100" height="60" fill="#132a4a"/>
  <g fill="none" stroke="#ede5d2" stroke-width="2.5">
    <circle cx="38" cy="30" r="22"/>
    <circle cx="62" cy="30" r="22"/>
  </g>
</svg>
"""
with open(os.path.join(SITE, "favicon.svg"), "w", encoding="utf-8") as f:
    f.write(favicon_svg)

# robots
with open(os.path.join(SITE, "robots.txt"), "w", encoding="utf-8") as f:
    f.write("""User-agent: *
Allow: /
Sitemap: https://sovereigntypath.org/sitemap.xml

# Public AI usage boundary:
# Search and real-time retrieval/citation are allowed for the public site.
# Model training or fine-tuning permission is not granted by this file.
# See: https://sovereigntypath.org/api/v1/ai-usage.json
Content-Signal: search=yes,ai-input=yes,ai-train=no
""")

# sitemap
sitemap_urls = [
    "/", "/governance.html", "/glossary.html", "/author.html",
    "/api.html", "/llms.txt", "/api/v1/index.json", "/api/v1/openapi.json",
    "/api/v1/ai-usage.json", "/.well-known/ai-readme.json",
    "/bundle/", "/bundle/manifesto.html", "/bundle/companion.html",
    "/bundle/operational-definitions.html", "/bundle/material-architecture.html",
]
sitemap = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
for u in sitemap_urls:
    sitemap += f'<url><loc>https://sovereigntypath.org{u}</loc></url>\n'
sitemap += "</urlset>\n"
with open(os.path.join(SITE, "sitemap.xml"), "w", encoding="utf-8") as f:
    f.write(sitemap)

print("Site build complete:", SITE)
