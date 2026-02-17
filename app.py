import io
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================
# CONFIG
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
DARK = "#111827"
MUTED = "#6B7280"

CATERING_POSTS = [
    "Accueil café",
    "Pause matin",
    "Déjeuner",
    "Pause après-midi",
    "Cocktail",
    "Boissons (global)",
    "Options",
    "Autres (logistique)",
]

TECH_POSTS = [
    "Périmètre",
    "Équipe",
    "Captation",
    "Régie",
    "Diffusion",
    "Replay",
    "Inclus",
    "Contraintes / options",
    "Conseil",
]


# =========================
# UTIL TEXT
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ")
    s = s.replace("\t", " ").replace("\n", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def fold(s: str) -> str:
    return norm(s).lower()


def extract_pdf_text(file) -> str:
    reader = PdfReader(file)
    parts = []
    for page in reader.pages:
        t = page.extract_text() or ""
        parts.append(t)
    return "\n".join(parts)


def guess_vendor_name(text: str, filename: str) -> str:
    # Heuristique simple : première ligne "forte" sinon nom de fichier
    lines = [norm(x) for x in text.splitlines() if norm(x)]
    for ln in lines[:25]:
        if len(ln) >= 4 and len(ln) <= 60 and not re.search(r"\b(devis|facture|date|total|tva|siret|iban)\b", fold(ln)):
            return ln
    return filename.rsplit(".", 1)[0]


def parse_eur_amount(s: str) -> float | None:
    """
    Parse "10 159,42" or "10159.42" -> float
    """
    s = norm(s)
    s = s.replace("€", "").replace("EUR", "").replace("euros", "")
    # keep digits, space, comma, dot
    s = re.sub(r"[^0-9,.\s-]", "", s).strip()
    if not s:
        return None
    # If comma is decimal separator
    # Remove spaces thousand sep
    s2 = s.replace(" ", "")
    # If both comma and dot exist, assume dot thousand and comma decimal (rare) -> remove dots
    if "," in s2 and "." in s2:
        s2 = s2.replace(".", "")
    # Now comma -> dot
    s2 = s2.replace(",", ".")
    try:
        return float(s2)
    except Exception:
        return None


def find_total_ttc(text: str) -> float | None:
    """
    Finds Total TTC or NET A PAYER patterns.
    """
    t = text
    # Prefer "Total TTC"
    patterns = [
        r"total\s+ttc\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"net\s+a\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+à\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
    ]
    for pat in patterns:
        m = re.search(pat, fold(t), flags=re.IGNORECASE | re.DOTALL)
        if m:
            amt = parse_eur_amount(m.group(1))
            if amt is not None:
                return amt
    return None


def split_lines(text: str) -> List[str]:
    # Keep original-ish lines
    raw = text.splitlines()
    out = []
    for r in raw:
        rr = r.replace("\u00A0", " ")
        rr = re.sub(r"\s+", " ", rr).strip()
        if rr:
            out.append(rr)
    return out


def is_section_header(line: str) -> bool:
    l = fold(line)
    # Common headers in FR catering + technique
    keys = [
        "accueil", "petit-déjeuner", "petit déjeuner",
        "pause", "déjeuner", "dejeuner", "buffet",
        "cocktail", "apéritif", "aperitif",
        "boissons", "soft", "vin", "champagne",
        "livraison", "service", "personnel", "location", "vaisselle", "nappage",
        "captation", "diffusion", "live", "zoom", "replay", "régie", "regie",
        "installation", "ingénieur", "ingenieur", "cadreur", "réalisateur", "realisateur",
        "inclus", "option", "conseil",
    ]
    return any(k in l for k in keys)


def classify_catering_section(title: str) -> str:
    l = fold(title)
    # time hints (08h30 etc.) decide morning/afternoon
    # but we keep it simple: keyword first, time second.
    if "accueil" in l or "petit" in l:
        return "Accueil café"
    if "déjeuner" in l or "dejeuner" in l or "buffet" in l or "déjeunatoire" in l:
        return "Déjeuner"
    if "cocktail" in l or "apéritif" in l or "aperitif" in l:
        # if "déjeunatoire" -> lunch
        if "déjeunatoire" in l:
            return "Déjeuner"
        return "Cocktail"
    if "pause" in l:
        # try by time in title
        if re.search(r"\b(14|15|16)h", l):
            return "Pause après-midi"
        if re.search(r"\b(10|11|12)h", l):
            return "Pause matin"
        # default pause morning
        return "Pause matin"
    if "boisson" in l or "soft" in l or "vin" in l or "champagne" in l:
        return "Boissons (global)"
    if "option" in l:
        return "Options"
    if "livraison" in l or "service" in l or "personnel" in l or "vaisselle" in l or "location" in l or "nappage" in l:
        return "Autres (logistique)"
    return "Autres (logistique)"


def classify_tech_section(title: str) -> str:
    l = fold(title)
    if "conseil" in l:
        return "Conseil"
    if "inclus" in l:
        return "Inclus"
    if "équipe" in l or "ingenieur" in l or "ingénieur" in l or "cadreur" in l or "réalisateur" in l:
        return "Équipe"
    if "captation" in l or "caméra" in l or "camera" in l:
        return "Captation"
    if "régie" in l or "regie" in l:
        return "Régie"
    if "diffusion" in l or "live" in l or "zoom" in l:
        return "Diffusion"
    if "replay" in l or "wetransfer" in l:
        return "Replay"
    if "option" in l or "forfait" in l or "connexion" in l or "contraint" in l:
        return "Contraintes / options"
    if "prestation" in l or "incluant" in l or "installation" in l:
        return "Périmètre"
    return "Périmètre"


def extract_sections(lines: List[str]) -> List[Tuple[str, List[str]]]:
    """
    Returns list of (section_title, section_lines).
    Heuristic: a line that looks like a header starts a new section.
    """
    sections = []
    current_title = "Général"
    current = []

    for ln in lines:
        # If line looks like a header, start new section
        if is_section_header(ln) and (len(ln) <= 90):
            # Avoid splitting too aggressively on bullet lines
            if not ln.startswith(("•", "-", "–")):
                # flush current
                if current:
                    sections.append((current_title, current))
                current_title = ln
                current = []
                continue

        current.append(ln)

    if current:
        sections.append((current_title, current))

    return sections


def bulletize(lines: List[str]) -> List[str]:
    """
    Keep bullets if present; also treat lines starting with • or - as bullets.
    Returns clean bullet items.
    """
    items = []
    for ln in lines:
        l = norm(ln)
        if not l:
            continue
        if l.startswith(("•", "-", "–")):
            l2 = l.lstrip("•-– ").strip()
            if l2:
                items.append(l2)
        else:
            # keep short informative lines, but avoid table headings
            if len(l) <= 140 and not re.fullmatch(r"(désignation|quantité|p\.u\.|montant|tva|total).*", fold(l)):
                items.append(l)
    # de-dup consecutive
    out = []
    for it in items:
        if not out or fold(out[-1]) != fold(it):
            out.append(it)
    return out


def summarize_for_table(items: List[str], max_chars: int = 320) -> str:
    """
    1-2 lines max. Joins with " • " and truncates.
    """
    if not items:
        return ""
    s = " • ".join(items)
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) <= max_chars:
        return s
    return s[: max_chars - 5].rstrip() + " (...)"


# =========================
# DATA MODELS
# =========================
@dataclass
class CateringOffer:
    vendor: str
    total_ttc: float | None
    posts: Dict[str, List[str]]  # full items by post
    options: List[str]
    comment: str


@dataclass
class TechOffer:
    vendor: str
    total_ttc: float | None
    posts: Dict[str, List[str]]
    options: List[str]
    comment: str


# =========================
# PARSERS
# =========================
def parse_catering_offer(text: str, filename: str) -> CateringOffer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    sections = extract_sections(lines)

    posts: Dict[str, List[str]] = {k: [] for k in CATERING_POSTS}
    options: List[str] = []

    for title, body in sections:
        post = classify_catering_section(title)
        items = bulletize(body)

        # Separate "options" if lines contain explicit option markers
        for it in items:
            if "option" in fold(it) or "en option" in fold(it):
                options.append(it)
            else:
                posts[post].append(it)

    # Move explicit options into posts too (kept separate for table + details)
    posts["Options"].extend(options)

    # Light cleanup: keep only meaningful unique items per post
    for k in list(posts.keys()):
        cleaned = []
        seen = set()
        for it in posts[k]:
            key = fold(it)
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(it)
        posts[k] = cleaned

    return CateringOffer(
        vendor=vendor,
        total_ttc=total_ttc,
        posts=posts,
        options=options,
        comment="",
    )


def parse_tech_offer(text: str, filename: str) -> TechOffer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    sections = extract_sections(lines)

    posts: Dict[str, List[str]] = {k: [] for k in TECH_POSTS}
    options: List[str] = []

    for title, body in sections:
        post = classify_tech_section(title)
        items = bulletize(body)

        for it in items:
            if "option" in fold(it) or "forfait" in fold(it) or "connexion" in fold(it):
                options.append(it)
            else:
                posts[post].append(it)

    posts["Contraintes / options"].extend(options)

    # de-dup
    for k in list(posts.keys()):
        cleaned = []
        seen = set()
        for it in posts[k]:
            key = fold(it)
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(it)
        posts[k] = cleaned

    return TechOffer(
        vendor=vendor,
        total_ttc=total_ttc,
        posts=posts,
        options=options,
        comment="",
    )


# =========================
# WORD GENERATION (PREMIUM-LIGHT)
# =========================
def set_cell_shading(cell, fill_hex: str):
    """Set cell background color."""
    fill_hex = fill_hex.replace("#", "")
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, value in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "Calibri"
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_subtitle(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = None
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_small(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(9)
    run.font.color.rgb = None


def add_hr(doc: Document):
    doc.add_paragraph("")


def euro_fmt(x: float | None) -> str:
    if x is None:
        return "—"
    return f"{x:,.2f} €".replace(",", " ").replace(".", ",")


def build_word(
    event_title: str,
    event_date: str,
    guests: int,
    catering: List[CateringOffer],
    tech: List[TechOffer],
) -> bytes:
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    # Header
    add_title(doc, "SYNTHÈSE DEVIS — PRESTATAIRES")
    add_subtitle(doc, f"{event_title} — {event_date} — Sur la base de {guests} convives")
    add_small(doc, f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    add_hr(doc)

    # =========================
    # PAGE 1 — CATERING SUMMARY TABLE
    # =========================
    if catering:
        add_subtitle(doc, "1) PRESTATION TRAITEUR — Comparatif (synthèse)")

        vendors = catering[:3]
        col_count = 1 + len(vendors)  # left label + vendors
        table = doc.add_table(rows=1, cols=col_count)
        table.autofit = True

        # Header row
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Poste"
        set_cell_shading(hdr_cells[0], PRIMARY)
        hdr_cells[0].paragraphs[0].runs[0].font.color.rgb = None
        hdr_cells[0].paragraphs[0].runs[0].font.bold = True

        for i, off in enumerate(vendors, start=1):
            hdr_cells[i].text = off.vendor
            set_cell_shading(hdr_cells[i], PRIMARY)
            for r in hdr_cells[i].paragraphs[0].runs:
                r.font.bold = True
                r.font.size = Pt(10)

        for cell in hdr_cells:
            set_cell_margins(cell)

        # Rows
        summary_rows = [
            ("Total TTC (hors options)", lambda o: [euro_fmt(o.total_ttc)]),
            ("Accueil café", lambda o: [summarize_for_table(o.posts.get("Accueil café", []), 280)]),
            ("Pause matin", lambda o: [summarize_for_table(o.posts.get("Pause matin", []), 280)]),
            ("Déjeuner", lambda o: [summarize_for_table(o.posts.get("Déjeuner", []), 320)]),
            ("Pause après-midi", lambda o: [summarize_for_table(o.posts.get("Pause après-midi", []), 280)]),
            ("Cocktail", lambda o: [summarize_for_table(o.posts.get("Cocktail", []), 320)]),
            ("Boissons (global)", lambda o: [summarize_for_table(o.posts.get("Boissons (global)", []), 220)]),
            ("Options", lambda o: [summarize_for_table(o.posts.get("Options", []), 220)]),
            ("Commentaire", lambda o: [norm(o.comment)[:220] + (" (...)" if len(norm(o.comment)) > 220 else "")]),
        ]

        for label, fn in summary_rows:
            row_cells = table.add_row().cells
            row_cells[0].text = label
            row_cells[0].paragraphs[0].runs[0].font.bold = True
            row_cells[0].paragraphs[0].runs[0].font.size = Pt(9)
            set_cell_shading(row_cells[0], "F3F4F6")
            set_cell_margins(row_cells[0])

            for j, off in enumerate(vendors, start=1):
                val = fn(off)[0]
                row_cells[j].text = val if val else "—"
                # style
                for p in row_cells[j].paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9)
                set_cell_margins(row_cells[j])

        add_hr(doc)

    # =========================
    # PAGE 1 — TECH SUMMARY TABLE
    # =========================
    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        # 2-column table per tech offer (keep it light)
        for idx, off in enumerate(tech[:2], start=1):
            add_small(doc, f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")

            t = doc.add_table(rows=1, cols=2)
            t.autofit = True
            h = t.rows[0].cells
            h[0].text = "Item"
            h[1].text = "Détail"
            set_cell_shading(h[0], PRIMARY)
            set_cell_shading(h[1], PRIMARY)
            for cell in h:
                for r in cell.paragraphs[0].runs:
                    r.font.bold = True
                    r.font.size = Pt(10)
                set_cell_margins(cell)

            rows = [
                ("Périmètre", summarize_for_table(off.posts.get("Périmètre", []), 420)),
                ("Équipe", summarize_for_table(off.posts.get("Équipe", []), 420)),
                ("Captation", summarize_for_table(off.posts.get("Captation", []), 420)),
                ("Régie", summarize_for_table(off.posts.get("Régie", []), 420)),
                ("Diffusion", summarize_for_table(off.posts.get("Diffusion", []), 420)),
                ("Replay", summarize_for_table(off.posts.get("Replay", []), 420)),
                ("Inclus", summarize_for_table(off.posts.get("Inclus", []), 420)),
                ("Contraintes / options", summarize_for_table(off.posts.get("Contraintes / options", []), 420)),
                ("Conseil", norm(off.comment)[:420] + (" (...)" if len(norm(off.comment)) > 420 else "")),
            ]
            for k, v in rows:
                rc = t.add_row().cells
                rc[0].text = k
                rc[1].text = v if v else "—"
                rc[0].paragraphs[0].runs[0].font.bold = True
                rc[0].paragraphs[0].runs[0].font.size = Pt(9)
                set_cell_shading(rc[0], "F3F4F6")
                for cell in rc:
                    set_cell_margins(cell)
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(9)

            add_hr(doc)

    # =========================
    # DETAILS — MENUS (EXHAUSTIVE)
    # =========================
    # Start details section on new page
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (menus exhaustifs)")
    add_small(doc, "Objectif : permettre la vérification fine des pièces / contenus par poste.")

    # Catering details
    if catering:
        add_hr(doc)
        add_subtitle(doc, "A) TRAITEUR — Détail par prestataire")

        for i, off in enumerate(catering[:3], start=1):
            add_hr(doc)
            add_subtitle(doc, f"PROPOSITION {i} — {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")

            for post in ["Accueil café", "Pause matin", "Déjeuner", "Pause après-midi", "Cocktail", "Boissons (global)", "Options", "Autres (logistique)"]:
                items = off.posts.get(post, [])
                p = doc.add_paragraph()
                r = p.add_run(f"{post} :")
                r.bold = True
                r.font.size = Pt(10)

                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it, style=None)
                        para.style = doc.styles["List Bullet"]

            if off.comment.strip():
                p = doc.add_paragraph()
                r = p.add_run("Commentaire : ")
                r.bold = True
                doc.add_paragraph(off.comment)

    # Tech details
    if tech:
        add_hr(doc)
        add_subtitle(doc, "B) TECHNIQUE — Détail par prestataire")
        for i, off in enumerate(tech[:2], start=1):
            add_hr(doc)
            add_subtitle(doc, f"PRESTATION TECHNIQUE {i} — {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")

            for post in TECH_POSTS:
                items = off.posts.get(post, [])
                p = doc.add_paragraph()
                r = p.add_run(f"{post} :")
                r.bold = True
                r.font.size = Pt(10)

                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it)
                        para.style = doc.styles["List Bullet"]

            if off.comment.strip():
                p = doc.add_paragraph()
                r = p.add_run("Conseil (synthèse) : ")
                r.bold = True
                doc.add_paragraph(off.comment)

    # Save to bytes
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title=APP_TITLE, layout="wide")

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap');
html, body, .stApp, [class*="css"] {{
  font-family: 'Montserrat', sans-serif !important;
}}
header[data-testid="stHeader"] {{ display: none; }}
.stApp {{ background: #F6F7FB; }}
[data-testid="stHorizontalBlock"] {{
  background: white;
  border-radius: 16px;
  padding: 0.75rem 0.85rem;
  margin-bottom: 0.65rem;
  box-shadow: 0 1px 12px rgba(0,0,0,0.06);
}}
.stButton > button {{
  background-color: {PRIMARY} !important;
  color: #ffffff !important;
  border: none !important;
  border-radius: 14px !important;
  padding: 0.85rem 1.05rem !important;
  font-weight: 900 !important;
  min-height: 52px !important;
  white-space: nowrap !important;
}}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(f"## {APP_TITLE}")
st.caption("Upload devis PDF → extraction automatique → validation rapide → export Word (synthèse + détails exhaustifs).")
st.divider()

# Event info
c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="center")
with c1:
    event_title = st.text_input("Événement (titre court)", placeholder="Ex : Journée scientifique — Colloque Génétique et Société")
with c2:
    event_date = st.text_input("Date", placeholder="Ex : 19/03/2026")
with c3:
    guests = st.number_input("Nb convives", min_value=1, max_value=5000, value=100, step=10)

st.markdown("### 1) Devis traiteur (max 3)")
catering_files = st.file_uploader("Upload PDF traiteur", type=["pdf"], accept_multiple_files=True)

st.markdown("### 2) Devis technique (max 2)")
tech_files = st.file_uploader("Upload PDF technique", type=["pdf"], accept_multiple_files=True)

if (not catering_files) and (not tech_files):
    st.info("Commence par uploader au moins un devis PDF.")
    st.stop()

# Limit counts
catering_files = (catering_files or [])[:3]
tech_files = (tech_files or [])[:2]

# Parse
st.markdown("### Extraction & validation")
catering_offers: List[CateringOffer] = []
tech_offers: List[TechOffer] = []

with st.spinner("Lecture des PDFs…"):
    for f in catering_files:
        txt = extract_pdf_text(f)
        catering_offers.append(parse_catering_offer(txt, f.name))

    for f in tech_files:
        txt = extract_pdf_text(f)
        tech_offers.append(parse_tech_offer(txt, f.name))

# Validation UI
tab1, tab2 = st.tabs(["Traiteur", "Technique"])

with tab1:
    if not catering_offers:
        st.caption("Aucun devis traiteur uploadé.")
    else:
        st.caption("Corrige rapidement si besoin : nom prestataire, total TTC, contenus par poste, commentaire.")
        for idx, off in enumerate(catering_offers, start=1):
            with st.expander(f"PROPOSITION {idx} — {off.vendor}", expanded=(idx == 1)):
                off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"c_vendor_{idx}")
                ttc_str = st.text_input("Total TTC (hors options) — si vide, saisis-le", value=(""
                                                                                               if off.total_ttc is None
                                                                                               else euro_fmt(off.total_ttc)),
                                        key=f"c_ttc_{idx}")
                # Parse back
                ttc_val = parse_eur_amount(ttc_str)
                off.total_ttc = ttc_val

                colL, colR = st.columns(2)
                with colL:
                    for post in ["Accueil café", "Pause matin", "Déjeuner", "Pause après-midi"]:
                        content = "\n".join(off.posts.get(post, []))
                        edited = st.text_area(post, value=content, height=150, key=f"c_{idx}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                with colR:
                    for post in ["Cocktail", "Boissons (global)", "Options", "Autres (logistique)"]:
                        content = "\n".join(off.posts.get(post, []))
                        edited = st.text_area(post, value=content, height=150, key=f"c_{idx}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]

                off.comment = st.text_area("Commentaire (1–2 phrases max)", value=off.comment, height=80, key=f"c_comment_{idx}")

with tab2:
    if not tech_offers:
        st.caption("Aucun devis technique uploadé.")
    else:
        st.caption("Corrige rapidement : nom prestataire, total TTC, périmètre, inclus, options, conseil.")
        for idx, off in enumerate(tech_offers, start=1):
            with st.expander(f"TECHNIQUE {idx} — {off.vendor}", expanded=(idx == 1)):
                off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"t_vendor_{idx}")
                ttc_str = st.text_input("Total TTC — si vide, saisis-le", value=(""
                                                                                 if off.total_ttc is None
                                                                                 else euro_fmt(off.total_ttc)),
                                        key=f"t_ttc_{idx}")
                off.total_ttc = parse_eur_amount(ttc_str)

                # Compact layout
                col1, col2 = st.columns(2)
                with col1:
                    for post in ["Périmètre", "Équipe", "Captation", "Régie"]:
                        content = "\n".join(off.posts.get(post, []))
                        edited = st.text_area(post, value=content, height=120, key=f"t_{idx}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                with col2:
                    for post in ["Diffusion", "Replay", "Inclus", "Contraintes / options"]:
                        content = "\n".join(off.posts.get(post, []))
                        edited = st.text_area(post, value=content, height=120, key=f"t_{idx}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]

                off.comment = st.text_area("Conseil (2–3 phrases max)", value=off.comment, height=90, key=f"t_comment_{idx}")

st.divider()

# Generate doc
colA, colB = st.columns([2, 1], vertical_alignment="center")
with colA:
    st.caption("Conseil : garde la synthèse très courte, et mets l’exhaustif dans les menus détaillés.")
with colB:
    if st.button("Générer le Word (.docx)", use_container_width=True, type="primary"):
        docx_bytes = build_word(
            event_title=event_title.strip() or "Événement",
            event_date=event_date.strip() or "Date à préciser",
            guests=int(guests),
            catering=catering_offers,
            tech=tech_offers,
        )
        ts = datetime.now().strftime("%Y-%m-%d_%H%M")
        fname = f"synthese_devis_{ts}.docx"
        st.download_button(
            "⬇️ Télécharger la synthèse (Word)",
            data=docx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
