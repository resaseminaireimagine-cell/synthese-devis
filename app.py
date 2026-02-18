# app.py — Synthèse devis prestataires (FINAL FINAL — SIMPLE)
# ----------------------------------------------------------
# Philosophie:
# - Extraction auto pour pré-remplir (vendor + TTC + cocktail count + items tech)
# - Edition MINIMALE (moins usine à gaz):
#   * Vendor (texte)
#   * Total TTC (texte)
#   * Résumé cocktail (texte court) — optionnel
#   * Détail (un seul gros champ par prestataire) — c’est ce qui alimente le Word
#
# Word:
# - 1 tableau comparatif traiteur (court)
# - 1 tableau synthèse tech (court)
# - Section DÉTAILS très séparée + contenu modifiable (le gros champ)
#
# Dépendances: streamlit, pypdf, python-docx

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import streamlit as st
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================
# BRAND
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
BG = "#F6F7FB"
FONT = "Montserrat"


# =========================
# HELPERS — text
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").replace("\t", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fold(s: str) -> str:
    return norm(s).lower()


def extract_pdf_text(uploaded_file) -> str:
    reader = PdfReader(uploaded_file)
    return "\n".join([(p.extract_text() or "") for p in reader.pages])


def parse_eur_amount(s: str) -> Optional[float]:
    s = norm(s)
    if not s:
        return None
    s = s.replace("€", "").replace("EUR", "").replace("euros", "")
    s = re.sub(r"[^0-9,.\s-]", "", s).strip()
    if not s:
        return None
    s2 = s.replace(" ", "")
    if "," in s2 and "." in s2:
        s2 = s2.replace(".", "")
    s2 = s2.replace(",", ".")
    try:
        return float(s2)
    except Exception:
        return None


def euro_fmt(x: Optional[float]) -> str:
    if x is None:
        return "—"
    return f"{x:,.2f} €".replace(",", " ").replace(".", ",")


def find_total_ttc(text: str) -> Optional[float]:
    lt = fold(text)
    patterns = [
        r"total\s+ttc\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+devis\s+t\.t\.c\.\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"net\s+à\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"net\s+a\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+à\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
    ]
    found = []
    for pat in patterns:
        for m in re.finditer(pat, lt, flags=re.I | re.DOTALL):
            amt = parse_eur_amount(m.group(1))
            if amt is not None:
                found.append(amt)
    return found[-1] if found else None


def vendor_from_filename(filename: str) -> str:
    base = filename.rsplit(".", 1)[0]
    base = re.sub(r"[_\-]+", " ", base)
    base = re.sub(r"\s+", " ", base).strip()
    # enlever tokens parasites
    junk = {"v1", "v2", "v3", "devis", "dev", "institut", "imagine", "pax", "100pax"}
    toks = [t for t in base.split() if t.lower() not in junk]
    return " ".join(toks[:6]).strip() or base


def guess_vendor_name(text: str, filename: str) -> str:
    # Minimaliste: on préfère le fichier (stable), puis on tente 2 heuristiques simples.
    fallback = vendor_from_filename(filename)

    lines = [norm(x) for x in text.splitlines() if norm(x)]
    head = lines[:220]

    # 1) ligne juste au-dessus de "SIRET"
    for i, ln in enumerate(head):
        if "siret" in fold(ln) or "rcs" in fold(ln):
            for back in range(1, 10):
                j = i - back
                if j >= 0:
                    cand = head[j]
                    cand = re.split(r"(?i)\bau capital\b", cand)[0].strip(" -–,;:")
                    # coupe avant adresse si elle démarre
                    cand = re.split(r"(?i)\b(rue|avenue|boulevard|quai|impasse|route|chemin)\b", cand)[0].strip(" -–,;:")
                    # tue les fins " 1" de type page
                    cand = re.sub(r"\s+\d{1,2}$", "", cand).strip()
                    if len(cand) >= 4 and not any(w in fold(cand) for w in ["devis", "facture", "total", "tva"]):
                        return cand or fallback
            break

    # 2) sinon fallback filename
    return fallback


def extract_cocktail_summary(text: str) -> str:
    """
    Cherche un pattern cocktail du type:
    - "10 pièces par personne"
    - "5 pièces salées ... 2 pièces sucrées"
    On ne prend que ce qui est proche de 'cocktail' / 'apéritif' pour éviter les viennoiseries.
    """
    t = text
    lt = fold(t)

    # window autour du mot cocktail
    m = re.search(r"(cocktail|apéritif|aperitif)", lt)
    window = t if not m else t[max(0, m.start()-800): m.start()+1200]
    lw = fold(window)

    total = None
    sale = None
    sucre = None

    mt = re.search(r"\b(\d{1,2})\s*pi[eè]ces?\s*(par\s+personne|/pers|par\s+convive)", lw, flags=re.I)
    if mt:
        total = int(mt.group(1))

    ms = re.search(r"\b(\d{1,2})\s*pi[eè]ces?.{0,40}(sal[ée]es?|froides?)", lw, flags=re.I)
    if ms:
        sale = int(ms.group(1))

    mu = re.search(r"\b(\d{1,2})\s*pi[eè]ces?.{0,40}(sucr[ée]es?)", lw, flags=re.I)
    if mu:
        sucre = int(mu.group(1))

    if total or sale or sucre:
        bits = []
        if total:
            bits.append(f"{total} pièces/pers")
        if sale or sucre:
            sub = []
            if sale:
                sub.append(f"{sale} salées")
            if sucre:
                sub.append(f"{sucre} sucrées")
            bits.append(" + ".join(sub))
        opt = "option" in lw and "sucr" in lw
        if opt:
            bits.append("(sucré en option)")
        return " — ".join(bits)

    return "—"


def extract_tech_one_liner(text: str) -> str:
    """
    Un résumé technique simple: capte 6-10 mots clés présents.
    """
    l = fold(text)
    hits = []
    for k in ["captation", "caméra", "4k", "réalisateur", "cadreur", "ingénieur", "son", "régie", "zoom", "live", "replay", "wetransfer"]:
        if k in l:
            hits.append(k)
    if not hits:
        return "—"
    # dédoublonne + limite
    out = []
    for h in hits:
        if h not in out:
            out.append(h)
    return " + ".join(out[:10])


# =========================
# MODEL
# =========================
@dataclass
class Offer:
    kind: str  # "traiteur" | "tech"
    vendor: str
    total_ttc: Optional[float]
    summary: str  # cocktail summary (traiteur) or tech keywords (tech)
    detail: str   # gros texte modifiable (source pour Word)


# =========================
# WORD HELPERS
# =========================
def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.replace("#", "").strip()
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_run(run, bold=False, size=10, color="#111827"):
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = hex_to_rgb(color)


def add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, bold=True, size=16, color=PRIMARY)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_subtitle(doc: Document, text: str):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, bold=True, size=11, color="#111827")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_small(doc: Document, text: str):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, bold=False, size=9, color="#374151")


def set_cell_shading(cell, fill_hex: str):
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


def add_band(doc: Document, label: str, sub: str = ""):
    t = doc.add_table(rows=1, cols=1)
    cell = t.rows[0].cells[0]
    set_cell_shading(cell, PRIMARY)
    set_cell_margins(cell, top=140, bottom=140, start=160, end=160)
    p = cell.paragraphs[0]
    r = p.add_run(label)
    set_run(r, bold=True, size=11, color="#FFFFFF")
    if sub:
        p2 = cell.add_paragraph()
        r2 = p2.add_run(sub)
        set_run(r2, bold=False, size=9, color="#FFFFFF")
    doc.add_paragraph("")


def build_word(event_title: str, event_date: str, guests: int,
               catering: List[Offer], tech: List[Offer]) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    add_title(doc, "SYNTHÈSE DEVIS — PRESTATAIRES")
    add_subtitle(doc, f"{event_title} — {event_date} — Sur la base de {guests} convives")
    add_small(doc, f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("")

    # ---- TRAITEUR: comparatif court
    if catering:
        add_subtitle(doc, "1) PRESTATION TRAITEUR — Comparatif (synthèse)")
        vendors = catering[:3]
        table = doc.add_table(rows=1, cols=1 + len(vendors))

        hdr = table.rows[0].cells
        hdr[0].text = "Poste"
        set_cell_shading(hdr[0], PRIMARY)
        for rr in hdr[0].paragraphs[0].runs:
            set_run(rr, bold=True, size=10, color="#FFFFFF")

        for i, off in enumerate(vendors, start=1):
            hdr[i].text = off.vendor
            set_cell_shading(hdr[i], PRIMARY)
            for rr in hdr[i].paragraphs[0].runs:
                set_run(rr, bold=True, size=10, color="#FFFFFF")

        for c in hdr:
            set_cell_margins(c)

        rows = [
            ("Total TTC", lambda o: euro_fmt(o.total_ttc)),
            ("Accueil café", lambda o: "—"),   # volontairement court (on ne “devine” pas)
            ("Pause matin", lambda o: "—"),
            ("Déjeuner", lambda o: "—"),
            ("Cocktail", lambda o: o.summary or "—"),
            ("Boissons", lambda o: "—"),
            ("Options", lambda o: "—"),
        ]

        for label, fn in rows:
            r = table.add_row().cells
            r[0].text = label
            set_cell_shading(r[0], "F3F4F6")
            for rr in r[0].paragraphs[0].runs:
                set_run(rr, bold=True, size=9, color="#111827")
            set_cell_margins(r[0])

            for j, off in enumerate(vendors, start=1):
                r[j].text = fn(off) or "—"
                for p in r[j].paragraphs:
                    for rr in p.runs:
                        set_run(rr, bold=False, size=9, color="#111827")
                set_cell_margins(r[j])

        doc.add_paragraph("")

    # ---- TECH: comparatif court
    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        for idx, off in enumerate(tech[:2], start=1):
            p = doc.add_paragraph()
            rr = p.add_run(f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            set_run(rr, bold=True, size=9, color="#111827")
            p2 = doc.add_paragraph()
            rr2 = p2.add_run(f"Synthèse : {off.summary or '—'}")
            set_run(rr2, bold=False, size=9, color="#111827")
        doc.add_paragraph("")

    # ---- DETAILS
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (modifiable via l’outil)")
    add_small(doc, "Ci-dessous : contenu complet (champ Détail).")
    doc.add_paragraph("")

    if catering:
        add_title(doc, "DÉTAIL — PRESTATIONS TRAITEUR")
        doc.add_paragraph("")
        for off in catering[:3]:
            add_band(doc, off.vendor, f"Total TTC : {euro_fmt(off.total_ttc)}")
            p = doc.add_paragraph(off.detail.strip() or "—")
            for r in p.runs:
                set_run(r, bold=False, size=9, color="#111827")

    if tech:
        doc.add_page_break()
        add_title(doc, "DÉTAIL — PRESTATIONS TECHNIQUES")
        doc.add_paragraph("")
        for off in tech[:2]:
            add_band(doc, off.vendor, f"Total TTC : {euro_fmt(off.total_ttc)}")
            p = doc.add_paragraph(off.detail.strip() or "—")
            for r in p.runs:
                set_run(r, bold=False, size=9, color="#111827")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# STREAMLIT UI (1 écran)
# =========================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap');
html, body, .stApp, [class*="css"] {{ font-family: 'Montserrat', sans-serif !important; }}
header[data-testid="stHeader"] {{ display: none; }}
.stApp {{ background: {BG}; }}
[data-testid="stHorizontalBlock"] {{
  background: white; border-radius: 16px; padding: 0.75rem 0.85rem;
  margin-bottom: 0.65rem; box-shadow: 0 1px 12px rgba(0,0,0,0.06);
}}
.stButton > button {{
  background-color: {PRIMARY} !important; color: #ffffff !important; border: none !important;
  border-radius: 14px !important; padding: 0.85rem 1.05rem !important; font-weight: 900 !important;
  min-height: 52px !important; white-space: nowrap !important;
}}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(f"## {APP_TITLE}")
st.caption("Simple : tu corriges Nom + TTC + Cocktail (si besoin) + Détail. Le reste est automatique.")
st.divider()

c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="center")
with c1:
    event_title = st.text_input("Événement (titre court)", placeholder="Ex : Journée scientifique — Colloque Génétique et Société")
with c2:
    event_date = st.text_input("Date", placeholder="Ex : 19/03/2026")
with c3:
    guests = st.number_input("Nb convives", min_value=1, max_value=5000, value=100, step=10)

st.markdown("### Devis")
colA, colB = st.columns(2)
with colA:
    catering_files = st.file_uploader("PDF Traiteur (max 3)", type=["pdf"], accept_multiple_files=True)
with colB:
    tech_files = st.file_uploader("PDF Technique (max 2)", type=["pdf"], accept_multiple_files=True)

catering_files = (catering_files or [])[:3]
tech_files = (tech_files or [])[:2]

if not catering_files and not tech_files:
    st.info("Upload au moins un devis PDF.")
    st.stop()

offers_catering: List[Offer] = []
offers_tech: List[Offer] = []

with st.spinner("Lecture et pré-remplissage…"):
    for f in catering_files:
        txt = extract_pdf_text(f)
        vendor = guess_vendor_name(txt, f.name)
        ttc = find_total_ttc(txt)
        cocktail = extract_cocktail_summary(txt)
        # détail par défaut = texte brut (tu modifies)
        detail = norm(txt).replace("  ", " ")
        offers_catering.append(Offer(kind="traiteur", vendor=vendor, total_ttc=ttc, summary=cocktail, detail=detail))

    for f in tech_files:
        txt = extract_pdf_text(f)
        vendor = guess_vendor_name(txt, f.name)
        ttc = find_total_ttc(txt)
        summ = extract_tech_one_liner(txt)
        detail = norm(txt).replace("  ", " ")
        offers_tech.append(Offer(kind="tech", vendor=vendor, total_ttc=ttc, summary=summ, detail=detail))

st.markdown("### Édition (minimum vital)")
if offers_catering:
    st.subheader("Traiteur")
    for i, off in enumerate(offers_catering, start=1):
        with st.expander(f"Traiteur {i} — {off.vendor}", expanded=(i == 1)):
            off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"c_vendor_{i}")
            ttc_in = st.text_input("Total TTC", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"c_ttc_{i}")
            off.total_ttc = parse_eur_amount(ttc_in)

            off.summary = st.text_input(
                "Résumé cocktail (pour le comparatif) — ex : 10 pièces/pers — 6 salées + 4 sucrées",
                value=off.summary if off.summary != "—" else "",
                key=f"c_cocktail_{i}",
            ).strip() or "—"

            off.detail = st.text_area(
                "Détail (texte complet modifiable) — c’est CE champ qui alimente la section DÉTAIL du Word",
                value=off.detail,
                height=260,
                key=f"c_detail_{i}",
            )

if offers_tech:
    st.subheader("Technique")
    for i, off in enumerate(offers_tech, start=1):
        with st.expander(f"Technique {i} — {off.vendor}", expanded=(i == 1 and not offers_catering)):
            off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"t_vendor_{i}")
            ttc_in = st.text_input("Total TTC", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"t_ttc_{i}")
            off.total_ttc = parse_eur_amount(ttc_in)

            off.summary = st.text_input("Synthèse (mots-clés) — tu peux corriger", value=off.summary, key=f"t_sum_{i}")
            off.detail = st.text_area("Détail (texte complet modifiable)", value=off.detail, height=260, key=f"t_detail_{i}")

st.divider()
if st.button("Générer le Word (.docx)", use_container_width=True, type="primary"):
    docx_bytes = build_word(
        event_title=event_title.strip() or "Événement",
        event_date=event_date.strip() or "Date à préciser",
        guests=int(guests),
        catering=offers_catering,
        tech=offers_tech,
    )
    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    st.download_button(
        "⬇️ Télécharger la synthèse (Word)",
        data=docx_bytes,
        file_name=f"synthese_devis_{ts}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
