# app.py — V5 (robuste, pro, “mapping obligatoire”)
# ---------------------------------------------------
# Objectif:
# - Upload PDFs (Traiteur x3 max, Technique x2 max)
# - Extraction en BLOCS (titres/sections) + mapping obligatoire par bloc
# - Dans chaque bloc: sélection des lignes (checkbox) pour éviter bruit/CGV/tables
# - Garde-fous: Vendor non suspect + TTC >= seuil (500€ par défaut)
# - Export Word premium (charte Imagine, Montserrat si dispo)
#
# Points clés:
# - On n’essaie plus de “deviner” parfaitement: on assiste et on FIABILISE.
# - Le mapping par bloc rend l’outil utilisable sur des devis hétérogènes.
# ---------------------------------------------------

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import streamlit as st
from pypdf import PdfReader

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================
# BRAND / UI
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
BG = "#F6F7FB"
FONT = "Montserrat"

# Business rule
DEFAULT_TTC_MIN = 500.0

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

# Generic noise (drop)
NOISE_HINTS = [
    "conditions générales", "cgv", "rgpd", "données personnelles", "donnees personnelles",
    "propriété intellectuelle", "propriete intellectuelle", "droit à l'image", "droit a l'image",
    "siret", "rcs", "tva", "iban", "bic", "rib", "banque", "capital",
    "pénalité", "penalite", "recouvrement", "mise en demeure", "tribunal",
    "responsabilité", "responsabilite", "dommages", "intérêts", "interets",
    "déchéance", "decheance", "résolutoire", "resolutoire", "litige", "contestation", "dédommagement", "dedommagement",
    "adresse", "tél", "tel", "email", "e-mail", "www.", "site internet",
    "référence", "reference", "devis n", "date de devis", "date de validité", "signature",
    "mode de paiement", "facture", "net a payer", "net à payer", "net a payer",
    "base ht", "total ht", "total ttc", "page ",
    "société générale", "societe generale", "caisse d'epargne", "caisse d’épargne", "cic",
    "à moins de 24h", "a moins de 24h", "à moins de 48h", "a moins de 48h", "annulée", "annulee",
    "heure supplémentaire", "heures supplémentaires", "heure supplementaire", "heures supplementaires",
]

# Vendor validation (generic forbidden tokens)
VENDOR_FORBIDDEN = [
    "accueil", "pause", "déjeuner", "dejeuner", "cocktail", "boissons", "options",
    "scénographie", "scenographie", "livraison", "personnel", "installation",
    "déroulé", "deroule",
    "récapitulatif", "recapitulatif", "sur la base", "version", "hors options", "budget",
    "proposition", "détail", "detail",
]


# =========================
# TEXT HELPERS
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").replace("\t", " ").replace("\r", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def fold(s: str) -> str:
    return norm(s).lower()


def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.replace("#", "").strip()
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def extract_pdf_text(uploaded_file) -> str:
    reader = PdfReader(uploaded_file)
    parts = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts)


def split_lines(text: str) -> List[str]:
    # make bullets more separable
    text = text.replace("•", "\n• ")
    text = text.replace(":-", ": -")
    text = re.sub(r"([:;])([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2", text)

    raw = text.splitlines()
    out = []
    for r in raw:
        rr = r.replace("\u00A0", " ")
        rr = re.sub(r"\s+", " ", rr).strip()
        if rr:
            out.append(rr)
    return out


def looks_like_price_table_line(s: str) -> bool:
    core = re.sub(r"[€]", "", norm(s))
    if re.fullmatch(r"[\d\s,\.%]+", core) and sum(ch.isdigit() for ch in core) >= 6:
        return True
    return False


def is_noise_line(s: str) -> bool:
    s = norm(s)
    if not s:
        return True
    l = fold(s)

    if re.fullmatch(r"\d{1,3}", s):
        return True
    if looks_like_price_table_line(s):
        return True
    if any(k in l for k in NOISE_HINTS):
        return True
    if "institut imagine" in l and ("étage" in l or "etage" in l or "sur la base" in l):
        return True
    if (re.search(r"\b\d{1,2}h\d{0,2}\b", l) or re.search(r"\bde\s+\d{1,2}h\d{0,2}\b", l)) and len(l) <= 40:
        return True
    return False


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
    patterns = [
        r"total\s+ttc\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"net\s+a\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"net\s+à\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+à\s+payer.*?([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
    ]
    lt = fold(text)
    for pat in patterns:
        m = re.search(pat, lt, flags=re.IGNORECASE | re.DOTALL)
        if m:
            amt = parse_eur_amount(m.group(1))
            if amt is not None:
                return amt
    return None


def vendor_is_suspicious(v: str) -> bool:
    v = norm(v)
    if len(v) < 3:
        return True
    lv = fold(v)
    if "@" in v or "contact" in lv:
        return True
    if any(w in lv for w in VENDOR_FORBIDDEN):
        return True
    if any(k in lv for k in ["rue", "avenue", "boulevard", "france", "paris"]):
        return True
    return False


def guess_vendor_name(text: str, filename: str) -> str:
    # safe heuristic: never trust it fully; mapping+QC will force correction if wrong
    lines = [norm(x) for x in text.splitlines() if norm(x)]
    bad = re.compile(r"\b(devis|facture|date|total|tva|siret|iban|bic|net a payer|net à payer)\b", re.I)

    def penalty(ln: str) -> int:
        l = fold(ln)
        p = 0
        if "contact" in l:
            p += 10
        if "@" in ln:
            p += 12
        if re.search(r"\b0[1-9](\s?\d{2}){4}\b", ln):
            p += 12
        if any(w in l for w in VENDOR_FORBIDDEN):
            p += 14
        if any(k in l for k in ["rue", "avenue", "boulevard", "france", "paris"]):
            p += 6
        return p

    candidates = []
    for ln in lines[:240]:
        if len(ln) < 4 or len(ln) > 70:
            continue
        if bad.search(ln):
            continue
        if is_noise_line(ln):
            continue
        alpha = sum(ch.isalpha() for ch in ln)
        if alpha < 6:
            continue
        if (ln.upper() == ln) or (alpha / max(len(ln), 1) > 0.55):
            candidates.append(ln)

    if candidates:
        prefer = ["réceptions", "receptions", "traiteur", "production", "sas", "sarl"]
        best, best_score = None, -10**9
        for c in candidates:
            fc = fold(c)
            sc = 0
            for k in prefer:
                if k in fc:
                    sc += 6
            sc -= penalty(c)
            if sc > best_score:
                best_score = sc
                best = c
        if best:
            return best

    return filename.rsplit(".", 1)[0]


def is_section_header(line: str) -> bool:
    l = fold(line)
    keys = [
        "accueil", "petit-déjeuner", "petit déjeuner",
        "pause", "déjeuner", "dejeuner", "buffet",
        "cocktail", "apéritif", "aperitif",
        "boissons", "rafraîchissements", "rafraichissements",
        "livraison", "service", "personnel", "location", "vaisselle", "nappage", "scénographie", "scenographie",
        "captation", "diffusion", "live", "zoom", "replay", "régie", "regie",
        "installation", "ingénieur", "ingenieur", "cadreur", "réalisateur", "realisateur",
        "inclus", "option", "conseil",
        "prestation",
    ]
    return any(k in l for k in keys)


def extract_blocks(lines: List[str]) -> List[Tuple[str, List[str]]]:
    """
    Return blocks as (title, lines).
    Blocks are started by header-like lines.
    Unlike previous versions, we keep blocks fairly “raw” then user maps them.
    """
    blocks: List[Tuple[str, List[str]]] = []
    current_title = "Bloc (non identifié)"
    current: List[str] = []

    for ln in lines:
        # avoid splitting on table header
        if re.fullmatch(r"(RÉF|REF|QTÉ|QTE|PU|PU HT|MONTANT|TVA|TAUX|BASE HT)\b.*", ln, flags=re.I):
            current.append(ln)
            continue

        if is_section_header(ln) and len(ln) <= 110 and not ln.startswith(("•", "-", "–")):
            if current:
                blocks.append((current_title, current))
            current_title = ln
            current = []
            continue

        current.append(ln)

    if current:
        blocks.append((current_title, current))

    # remove blocks that are basically empty after noise filter
    cleaned = []
    for title, body in blocks:
        body2 = [x for x in body if not is_noise_line(x)]
        if body2:
            cleaned.append((title, body2))
    return cleaned


def unglue(s: str) -> str:
    s = norm(s)
    s = re.sub(r"(\d)\s*(pi[eè]ces)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2 • \3", s, flags=re.I)
    s = re.sub(r"(personne|convive|invité)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 • \2", s, flags=re.I)
    s = re.sub(r"([:;])([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2", s)
    return s


def extract_items_from_block(block_lines: List[str]) -> List[str]:
    """
    Extract “presentable” items from a block:
    - keep bullet lines
    - keep short lines that are not noise
    - avoid table-ish lines
    """
    items: List[str] = []
    for ln in block_lines:
        s = norm(ln)
        if not s or is_noise_line(s):
            continue

        if s.startswith(("•", "-", "–")):
            it = s.lstrip("•-– ").strip()
            if it and not is_noise_line(it):
                items.append(unglue(it))
            continue

        # keep short descriptive lines (avoid admin by noise filter already)
        if len(s) <= 220 and not looks_like_price_table_line(s):
            items.append(unglue(s))

    # de-dup
    out, seen = [], set()
    for it in items:
        k = fold(it)
        if k in seen:
            continue
        seen.add(k)
        out.append(it)
    return out


def summarize_for_table(items: List[str], max_chars: int = 380) -> str:
    if not items:
        return ""
    s = " • ".join(items)
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) <= max_chars:
        return s
    return s[: max_chars - 5].rstrip() + " (...)"


# =========================
# MODEL
# =========================
@dataclass
class Offer:
    kind: str  # "catering" or "tech"
    vendor: str
    total_ttc: Optional[float]
    blocks: List[Tuple[str, List[str]]]                 # raw blocks (title + lines)
    block_map: Dict[int, str]                           # block index -> mapped post
    selected_by_post: Dict[str, List[str]]              # final selected items per post
    comment: str


def init_offer(kind: str, text: str, filename: str) -> Offer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    blocks = extract_blocks(filtered)

    posts = CATERING_POSTS if kind == "catering" else TECH_POSTS
    selected_by_post = {p: [] for p in posts}

    return Offer(
        kind=kind,
        vendor=vendor,
        total_ttc=total_ttc,
        blocks=blocks,
        block_map={},  # filled in UI
        selected_by_post=selected_by_post,
        comment="",
    )


# =========================
# WORD GENERATION
# =========================
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


def build_word(event_title: str, event_date: str, guests: int, catering: List[Offer], tech: List[Offer]) -> bytes:
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

    # PAGE 1 — Catering complete table
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

        def cell_for(off: Offer, post: str, max_chars: int) -> str:
            if post == "Total TTC (hors options)":
                return euro_fmt(off.total_ttc)
            return summarize_for_table(off.selected_by_post.get(post, []), max_chars)

        rows = [
            ("Total TTC (hors options)", 0),
            ("Accueil café", 320),
            ("Pause matin", 320),
            ("Déjeuner", 380),
            ("Pause après-midi", 320),
            ("Cocktail", 380),
            ("Boissons (global)", 260),
            ("Options", 260),
            ("Commentaire", 260),
        ]

        for label, maxc in rows:
            r = table.add_row().cells
            r[0].text = label
            set_cell_shading(r[0], "F3F4F6")
            for rr in r[0].paragraphs[0].runs:
                set_run(rr, bold=True, size=9, color="#111827")
            set_cell_margins(r[0])

            for j, off in enumerate(vendors, start=1):
                if label == "Commentaire":
                    val = norm(off.comment)
                    val = val[:maxc] + (" (...)" if len(val) > maxc else "")
                else:
                    val = cell_for(off, label, maxc) if label != "Total TTC (hors options)" else euro_fmt(off.total_ttc)

                r[j].text = val if val else "—"
                for p in r[j].paragraphs:
                    for rr in p.runs:
                        set_run(rr, bold=False, size=9, color="#111827")
                set_cell_margins(r[j])

        doc.add_paragraph("")

    # PAGE 1 — Tech summary tables (table per vendor)
    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        for idx, off in enumerate(tech[:2], start=1):
            p = doc.add_paragraph()
            rr = p.add_run(f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            set_run(rr, bold=True, size=9, color="#111827")

            t = doc.add_table(rows=1, cols=2)
            h = t.rows[0].cells
            h[0].text = "Item"
            h[1].text = "Détail"
            set_cell_shading(h[0], PRIMARY)
            set_cell_shading(h[1], PRIMARY)
            for cell in h:
                for rrr in cell.paragraphs[0].runs:
                    set_run(rrr, bold=True, size=10, color="#FFFFFF")
                set_cell_margins(cell)

            items = [
                ("Périmètre", 520),
                ("Équipe", 520),
                ("Captation", 520),
                ("Régie", 520),
                ("Diffusion", 520),
                ("Replay", 520),
                ("Inclus", 520),
                ("Contraintes / options", 520),
                ("Conseil", 520),
            ]
            for k, maxc in items:
                rrrow = t.add_row().cells
                rrrow[0].text = k
                rrrow[1].text = summarize_for_table(off.selected_by_post.get(k, []), maxc) if k != "Conseil" else (norm(off.comment)[:maxc] + (" (...)" if len(norm(off.comment)) > maxc else ""))
                set_cell_shading(rrrow[0], "F3F4F6")
                for rrr in rrrow[0].paragraphs[0].runs:
                    set_run(rrr, bold=True, size=9, color="#111827")
                for cell in rrrow:
                    set_cell_margins(cell)
                    for p in cell.paragraphs:
                        for rrr in p.runs:
                            set_run(rrr, bold=False, size=9, color="#111827")

            doc.add_paragraph("")

    # DETAILS
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (éléments validés)")
    add_small(doc, "Les listes ci-dessous correspondent aux éléments cochés (contrôle qualité).")

    if catering:
        doc.add_paragraph("")
        add_subtitle(doc, "A) TRAITEUR — Détail par prestataire")
        for off in catering[:3]:
            doc.add_paragraph("")
            add_subtitle(doc, f"{off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            for post in CATERING_POSTS:
                items = off.selected_by_post.get(post, [])
                p = doc.add_paragraph()
                r = p.add_run(f"{post} :")
                set_run(r, bold=True, size=10, color="#111827")
                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it)
                        para.style = doc.styles["List Bullet"]
                        for rr in para.runs:
                            set_run(rr, bold=False, size=9, color="#111827")
            if norm(off.comment):
                p = doc.add_paragraph()
                r = p.add_run("Commentaire : ")
                set_run(r, bold=True, size=10, color="#111827")
                doc.add_paragraph(off.comment)

    if tech:
        doc.add_paragraph("")
        add_subtitle(doc, "B) TECHNIQUE — Détail par prestataire")
        for off in tech[:2]:
            doc.add_paragraph("")
            add_subtitle(doc, f"{off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            for post in TECH_POSTS:
                items = off.selected_by_post.get(post, [])
                p = doc.add_paragraph()
                r = p.add_run(f"{post} :")
                set_run(r, bold=True, size=10, color="#111827")
                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it)
                        para.style = doc.styles["List Bullet"]
                        for rr in para.runs:
                            set_run(rr, bold=False, size=9, color="#111827")
            if norm(off.comment):
                p = doc.add_paragraph()
                r = p.add_run("Conseil (synthèse) : ")
                set_run(r, bold=True, size=10, color="#111827")
                doc.add_paragraph(off.comment)

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
st.caption("V5 : mapping par blocs obligatoire → fiabilité. Puis sélection des lignes → Word premium.")
st.divider()

ttc_min = st.number_input("Seuil TTC minimum (bloquant) — par prestataire", min_value=0, max_value=100000, value=int(DEFAULT_TTC_MIN), step=50)

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

catering_files = (catering_files or [])[:3]
tech_files = (tech_files or [])[:2]

with st.spinner("Lecture des PDFs…"):
    catering_offers: List[Offer] = []
    tech_offers: List[Offer] = []

    for f in catering_files:
        txt = extract_pdf_text(f)
        catering_offers.append(init_offer("catering", txt, f.name))

    for f in tech_files:
        txt = extract_pdf_text(f)
        tech_offers.append(init_offer("tech", txt, f.name))

# UI: mapping + selection
tab1, tab2 = st.tabs(["Traiteur — mapping + sélection", "Technique — mapping + sélection"])

def render_offer_mapping_and_selection(off: Offer, idx: int):
    posts = CATERING_POSTS if off.kind == "catering" else TECH_POSTS

    off.vendor = st.text_input("Nom prestataire (obligatoire)", value=off.vendor, key=f"{off.kind}_vendor_{idx}")
    ttc_in = st.text_input("Total TTC (obligatoire)", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"{off.kind}_ttc_{idx}")
    off.total_ttc = parse_eur_amount(ttc_in)

    # QC
    if vendor_is_suspicious(off.vendor):
        st.error("Nom prestataire invalide/suspect (titre interne, poste, contact ou adresse). Corrige-le.")
    if off.total_ttc is None:
        st.warning("Total TTC non détecté — saisis-le.")
    elif off.total_ttc < float(ttc_min):
        st.error(f"TTC < seuil ({ttc_min}€). Corrige le TTC (détection erronée) ou ajuste le seuil si nécessaire.")

    st.markdown("#### Étape 1 — Mapper les blocs (obligatoire)")
    st.caption("Pour chaque bloc détecté, choisis le poste correspondant. Ensuite tu sélectionnes les lignes utiles.")

    # mapping per block
    for b_idx, (title, body) in enumerate(off.blocks):
        key = f"{off.kind}_map_{idx}_{b_idx}"
        default = off.block_map.get(b_idx, "")
        choice = st.selectbox(
            f"Bloc {b_idx+1} — {title}",
            options=["(ignorer)"] + posts,
            index=0 if default == "" else (posts.index(default) + 1),
            key=key,
        )
        if choice == "(ignorer)":
            off.block_map[b_idx] = ""
        else:
            off.block_map[b_idx] = choice

    # enforce mapping: at least one mapped block
    mapped_any = any(v for v in off.block_map.values())
    if not mapped_any:
        st.error("Aucun bloc mappé. Mappe au moins un bloc pour continuer.")
        return False

    st.markdown("#### Étape 2 — Sélectionner les lignes (checkbox)")
    # build candidates per post from mapped blocks
    candidates_by_post: Dict[str, List[str]] = {p: [] for p in posts}
    for b_idx, (title, body) in enumerate(off.blocks):
        target = off.block_map.get(b_idx, "")
        if not target:
            continue
        items = extract_items_from_block(body)
        candidates_by_post[target].extend(items)

    # de-dup candidates per post
    for p in posts:
        out, seen = [], set()
        for it in candidates_by_post[p]:
            k = fold(it)
            if k in seen:
                continue
            seen.add(k)
            out.append(it)
        candidates_by_post[p] = out

    # selection UI
    for p in posts:
        cand = candidates_by_post[p]
        if not cand:
            continue
        st.markdown(f"**{p}**")
        sel = []
        default_n = 10 if p in ["Déjeuner", "Cocktail", "Régie", "Captation", "Diffusion"] else 6
        for j, line in enumerate(cand):
            checked = st.checkbox(line, value=(j < default_n), key=f"{off.kind}_sel_{idx}_{p}_{j}")
            if checked:
                sel.append(line)
        off.selected_by_post[p] = sel

    off.comment = st.text_area("Commentaire / Conseil", value=off.comment, height=90, key=f"{off.kind}_comment_{idx}")

    return True

with tab1:
    if not catering_offers:
        st.caption("Aucun devis traiteur.")
    else:
        for i, off in enumerate(catering_offers, start=1):
            with st.expander(f"Traiteur {i} — {off.vendor}", expanded=(i == 1)):
                render_offer_mapping_and_selection(off, i)

with tab2:
    if not tech_offers:
        st.caption("Aucun devis technique.")
    else:
        for i, off in enumerate(tech_offers, start=1):
            with st.expander(f"Technique {i} — {off.vendor}", expanded=(i == 1)):
                render_offer_mapping_and_selection(off, i)

# Final gate
blocked = False
for off in catering_offers:
    if vendor_is_suspicious(off.vendor) or off.total_ttc is None or off.total_ttc < float(ttc_min):
        blocked = True
    if not any(v for v in off.block_map.values()):
        blocked = True
for off in tech_offers:
    if not norm(off.vendor) or off.total_ttc is None or off.total_ttc < float(ttc_min):
        blocked = True
    if not any(v for v in off.block_map.values()):
        blocked = True

st.divider()
colA, colB = st.columns([2, 1], vertical_alignment="center")
with colA:
    st.caption("Génération bloquée si: vendor suspect, TTC manquant ou < seuil, ou aucun bloc mappé.")
with colB:
    if st.button("Générer le Word premium (.docx)", use_container_width=True, type="primary", disabled=blocked):
        docx_bytes = build_word(
            event_title=event_title.strip() or "Événement",
            event_date=event_date.strip() or "Date à préciser",
            guests=int(guests),
            catering=catering_offers,
            tech=tech_offers,
        )
        ts = datetime.now().strftime("%Y-%m-%d_%H%M")
        st.download_button(
            "⬇️ Télécharger la synthèse (Word)",
            data=docx_bytes,
            file_name=f"synthese_devis_{ts}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
