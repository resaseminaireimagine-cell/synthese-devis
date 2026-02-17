# app.py
# --------------------------------------------------------------------
# Synthèse devis prestataires — Institut Imagine (V4)
# Objectif:
# - Upload PDFs (traiteur + technique)
# - Extraction robuste (anti bruit: CGV/IBAN/tableaux/pages)
# - Sélection assistée (checkbox) => fiabilité
# - Contrôle qualité bloquant (vendor + total TTC requis, vendor non "titre")
# - Word premium (charte: rose Imagine, police Montserrat si disponible)
#
# Notes importantes:
# - Les PDFs "marketing" (avec photos) sont très bruités => on privilégie
#   une extraction de candidats + validation par sélection.
# - Montserrat dans Word dépend de l'installation sur le poste lecteur.
# --------------------------------------------------------------------

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Tuple

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

# Generic “menu” signal (kept)
MENU_HINTS = [
    "café", "cafe", "thé", "the", "soft", "jus", "eau", "viennoiser",
    "accueil", "petit", "déjeuner", "dejeuner", "pause", "buffet", "cocktail", "apéritif", "aperitif",
    "pièce", "pieces", "pièces", "/pers", "par personne", "convive", "invité",
    "salée", "sucrée", "dessert", "mignard", "gourmand",
    "sandwich", "wrap", "salade", "fromage", "fruit", "tartelette", "cannel", "financier",
    "vin", "champagne",
]

# Generic “tech” signal (kept)
TECH_HINTS = [
    "captation", "diffusion", "live", "zoom", "replay", "wetransfer", "stream",
    "réalisateur", "realisateur", "cadreur", "ingénieur", "ingenieur", "son",
    "caméra", "camera", "régie", "regie", "installation", "direct",
    "duplex", "plateforme", "tv", "écran", "ecran", "écrans", "ecrans",
    "pavlov", "zapette", "retour", "we transfer",
]

# Generic noise (dropped)
NOISE_HINTS = [
    # legal/admin/payment
    "conditions générales", "cgv", "rgpd", "données personnelles", "donnees personnelles",
    "propriété intellectuelle", "propriete intellectuelle", "droit à l'image", "droit a l'image",
    "siret", "rcs", "tva", "iban", "bic", "rib", "banque", "capital",
    "pénalité", "penalite", "recouvrement", "mise en demeure", "tribunal",
    "responsabilité", "responsabilite", "dommages", "intérêts", "interets",
    "déchéance", "decheance", "résolutoire", "resolutoire", "litige", "contestation", "dédommagement", "dedommagement",
    "adresse", "tél", "tel", "email", "e-mail", "www.", "site internet",
    "référence", "reference", "devis n", "date de devis", "date de validité", "signature",
    "mode de paiement", "facture", "net a payer", "net à payer", "net à payer",
    "base ht", "total ht", "total ttc", "page ",
    "société générale", "societe generale", "caisse d'epargne", "caisse d’épargne", "cic",
    # cancellation boilerplate
    "à moins de 24h", "a moins de 24h", "à moins de 48h", "a moins de 48h", "annulée", "annulee",
    # extra “logistic pricing clause” noise
    "heure supplémentaire", "heures supplémentaires", "heure supplementaire", "heures supplementaires",
]

# Vendor validation (generic forbidden tokens)
VENDOR_FORBIDDEN = [
    # meal posts / internal titles
    "accueil", "pause", "déjeuner", "dejeuner", "cocktail", "boissons", "options",
    "scénographie", "scenographie", "livraison", "personnel", "installation", "déroulé", "deroule",
    # report/recap titles
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
    # improve segmentation of some PDFs
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


def parse_eur_amount(s: str) -> float | None:
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


def euro_fmt(x: float | None) -> str:
    if x is None:
        return "—"
    return f"{x:,.2f} €".replace(",", " ").replace(".", ",")


def find_total_ttc(text: str) -> float | None:
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

    # boilerplate for marketing PDFs
    if "institut imagine" in l and ("étage" in l or "etage" in l or "sur la base" in l):
        return True

    # short time-only lines
    if (re.search(r"\b\d{1,2}h\d{0,2}\b", l) or re.search(r"\bde\s+\d{1,2}h\d{0,2}\b", l)) and len(l) <= 40:
        return True

    return False


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


def extract_sections(lines: List[str]) -> List[Tuple[str, List[str]]]:
    sections: List[Tuple[str, List[str]]] = []
    current_title = "Général"
    current: List[str] = []

    for ln in lines:
        # avoid section split on table headers
        if re.fullmatch(r"(RÉF|REF|QTÉ|QTE|PU|PU HT|MONTANT|TVA|TAUX|BASE HT)\b.*", ln, flags=re.I):
            current.append(ln)
            continue

        if is_section_header(ln) and len(ln) <= 95 and not ln.startswith(("•", "-", "–")):
            if current:
                sections.append((current_title, current))
            current_title = ln
            current = []
            continue

        current.append(ln)

    if current:
        sections.append((current_title, current))
    return sections


def unglue(s: str) -> str:
    s = norm(s)
    s = re.sub(r"(\d)\s*(pi[eè]ces)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2 • \3", s, flags=re.I)
    s = re.sub(r"(personne|convive|invité)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 • \2", s, flags=re.I)
    s = s.replace(":-", ": -")
    s = re.sub(r"([:;])([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2", s)
    return s


def bullet_candidates(lines: List[str], mode: str) -> List[str]:
    """
    mode:
      - 'catering' => keep MENU_HINTS
      - 'tech'     => keep TECH_HINTS
    """
    items: List[str] = []
    for ln in lines:
        s = norm(ln)
        if not s or is_noise_line(s):
            continue

        if s.startswith(("•", "-", "–")):
            s2 = s.lstrip("•-– ").strip()
            if s2 and not is_noise_line(s2):
                items.append(unglue(s2))
            continue

        l = fold(s)
        if len(s) <= 200:
            if mode == "catering" and any(k in l for k in MENU_HINTS):
                items.append(unglue(s))
            elif mode == "tech" and any(k in l for k in TECH_HINTS):
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


def vendor_is_suspicious(v: str) -> bool:
    v = norm(v)
    if len(v) < 3:
        return True
    lv = fold(v)
    if "@" in v or "contact" in lv:
        return True
    if any(w in lv for w in VENDOR_FORBIDDEN):
        return True
    # address-like lines
    if any(k in lv for k in ["rue", "avenue", "boulevard", "france", "paris"]):
        # not always invalid, but suspicious as vendor name
        return True
    return False


def guess_vendor_name(text: str, filename: str) -> str:
    """
    Generic vendor guess (safe):
    - find “brand-like” lines
    - penalize contacts, addresses, forbidden words
    """
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


def classify_catering_section(title: str) -> str:
    l = fold(title)
    if "accueil" in l or "petit" in l:
        return "Accueil café"
    if "déjeuner" in l or "dejeuner" in l or "buffet" in l or "déjeunatoire" in l:
        return "Déjeuner"
    if "cocktail" in l or "apéritif" in l or "aperitif" in l:
        if "déjeunatoire" in l:
            return "Déjeuner"
        return "Cocktail"
    if "pause" in l:
        if re.search(r"\b(14|15|16)h", l):
            return "Pause après-midi"
        if re.search(r"\b(10|11|12)h", l):
            return "Pause matin"
        return "Pause matin"
    if "boisson" in l or "soft" in l or "vin" in l or "rafraîch" in l or "rafraich" in l or "champagne" in l:
        return "Boissons (global)"
    if "option" in l:
        return "Options"
    if any(k in l for k in ["livraison", "service", "personnel", "vaisselle", "location", "nappage", "scénographie", "scenographie"]):
        return "Autres (logistique)"
    return "Autres (logistique)"


def classify_tech_section(title: str) -> str:
    l = fold(title)
    if "conseil" in l:
        return "Conseil"
    if "inclus" in l:
        return "Inclus"
    if any(k in l for k in ["équipe", "ingenieur", "ingénieur", "cadreur", "réalisateur", "realisateur", "son"]):
        return "Équipe"
    if "captation" in l or "caméra" in l or "camera" in l:
        return "Captation"
    if "régie" in l or "regie" in l:
        return "Régie"
    if "diffusion" in l or "live" in l or "zoom" in l or "stream" in l:
        return "Diffusion"
    if "replay" in l or "wetransfer" in l:
        return "Replay"
    if any(k in l for k in ["option", "forfait", "connexion", "contraint"]):
        return "Contraintes / options"
    if any(k in l for k in ["prestation", "incluant", "installation", "direct"]):
        return "Périmètre"
    return "Périmètre"


# =========================
# MODEL
# =========================
@dataclass
class Offer:
    vendor: str
    total_ttc: float | None
    candidates: Dict[str, List[str]]
    selected: Dict[str, List[str]]
    comment: str


def parse_catering(text: str, filename: str) -> Offer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    sections = extract_sections(filtered)

    candidates = {k: [] for k in CATERING_POSTS}
    for title, body in sections:
        post = classify_catering_section(title)
        items = bullet_candidates(body, mode="catering")
        for it in items:
            if "en option" in fold(it) or fold(it).startswith("option"):
                candidates["Options"].append(it)
            else:
                candidates[post].append(it)

    # de-dup
    for k in candidates:
        out, seen = [], set()
        for it in candidates[k]:
            kk = fold(it)
            if kk in seen:
                continue
            seen.add(kk)
            out.append(it)
        candidates[k] = out

    selected = {k: [] for k in CATERING_POSTS}
    return Offer(vendor, total_ttc, candidates, selected, "")


def parse_tech(text: str, filename: str) -> Offer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    sections = extract_sections(filtered)

    candidates = {k: [] for k in TECH_POSTS}
    for title, body in sections:
        post = classify_tech_section(title)
        items = bullet_candidates(body, mode="tech")

        # Additionally, keep short lines with TECH_HINTS even if not bulletized
        # (already handled by bullet_candidates), but we also want to keep some structured phrases
        # that may not match hints perfectly (e.g., "2 caméras 4K").
        for ln in body:
            s = norm(ln)
            if not s or is_noise_line(s):
                continue
            l = fold(s)
            if len(s) <= 200 and (("4k" in l) or ("cam" in l) or ("zoom" in l) or ("régi" in l) or ("regi" in l)):
                items.append(unglue(s))

        # de-dup
        out, seen = [], set()
        for it in items:
            kk = fold(it)
            if kk in seen:
                continue
            seen.add(kk)
            out.append(it)

        candidates[post].extend(out)

    # final de-dup per post
    for k in candidates:
        out, seen = [], set()
        for it in candidates[k]:
            kk = fold(it)
            if kk in seen:
                continue
            seen.add(kk)
            out.append(it)
        candidates[k] = out

    selected = {k: [] for k in TECH_POSTS}
    return Offer(vendor, total_ttc, candidates, selected, "")


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

    if catering:
        add_subtitle(doc, "1) PRESTATION TRAITEUR — Comparatif (synthèse)")
        vendors = catering[:3]
        table = doc.add_table(rows=1, cols=1 + len(vendors))

        hdr = table.rows[0].cells
        hdr[0].text = "Poste"
        set_cell_shading(hdr[0], PRIMARY)
        for r in hdr[0].paragraphs[0].runs:
            set_run(r, bold=True, size=10, color="#FFFFFF")

        for i, off in enumerate(vendors, start=1):
            hdr[i].text = off.vendor
            set_cell_shading(hdr[i], PRIMARY)
            for r in hdr[i].paragraphs[0].runs:
                set_run(r, bold=True, size=10, color="#FFFFFF")

        for c in hdr:
            set_cell_margins(c)

        rows = [
            ("Total TTC (hors options)", lambda o: euro_fmt(o.total_ttc)),
            ("Accueil café", lambda o: summarize_for_table(o.selected.get("Accueil café", []), 320)),
            ("Pause matin", lambda o: summarize_for_table(o.selected.get("Pause matin", []), 320)),
            ("Déjeuner", lambda o: summarize_for_table(o.selected.get("Déjeuner", []), 380)),
            ("Pause après-midi", lambda o: summarize_for_table(o.selected.get("Pause après-midi", []), 320)),
            ("Cocktail", lambda o: summarize_for_table(o.selected.get("Cocktail", []), 380)),
            ("Boissons (global)", lambda o: summarize_for_table(o.selected.get("Boissons (global)", []), 260)),
            ("Options", lambda o: summarize_for_table(o.selected.get("Options", []), 260)),
            ("Commentaire", lambda o: norm(o.comment)[:260] + (" (...)" if len(norm(o.comment)) > 260 else "")),
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

    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        for idx, off in enumerate(tech[:2], start=1):
            p = doc.add_paragraph()
            r = p.add_run(f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            set_run(r, bold=True, size=9, color="#111827")

            t = doc.add_table(rows=1, cols=2)
            h = t.rows[0].cells
            h[0].text = "Item"
            h[1].text = "Détail"
            set_cell_shading(h[0], PRIMARY)
            set_cell_shading(h[1], PRIMARY)
            for cell in h:
                for rr in cell.paragraphs[0].runs:
                    set_run(rr, bold=True, size=10, color="#FFFFFF")
                set_cell_margins(cell)

            items = [
                ("Périmètre", summarize_for_table(off.selected.get("Périmètre", []), 520)),
                ("Équipe", summarize_for_table(off.selected.get("Équipe", []), 520)),
                ("Captation", summarize_for_table(off.selected.get("Captation", []), 520)),
                ("Régie", summarize_for_table(off.selected.get("Régie", []), 520)),
                ("Diffusion", summarize_for_table(off.selected.get("Diffusion", []), 520)),
                ("Replay", summarize_for_table(off.selected.get("Replay", []), 520)),
                ("Inclus", summarize_for_table(off.selected.get("Inclus", []), 520)),
                ("Contraintes / options", summarize_for_table(off.selected.get("Contraintes / options", []), 520)),
                ("Conseil", norm(off.comment)[:520] + (" (...)" if len(norm(off.comment)) > 520 else "")),
            ]
            for k, v in items:
                rr = t.add_row().cells
                rr[0].text = k
                rr[1].text = v if v else "—"
                set_cell_shading(rr[0], "F3F4F6")
                for rrr in rr[0].paragraphs[0].runs:
                    set_run(rrr, bold=True, size=9, color="#111827")
                for cell in rr:
                    set_cell_margins(cell)
                    for p in cell.paragraphs:
                        for rrr in p.runs:
                            set_run(rrr, bold=False, size=9, color="#111827")
            doc.add_paragraph("")

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
                items = off.selected.get(post, [])
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
            if off.comment.strip():
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
                items = off.selected.get(post, [])
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
            if off.comment.strip():
                p = doc.add_paragraph()
                r = p.add_run("Conseil (synthèse) : ")
                set_run(r, bold=True, size=10, color="#111827")
                doc.add_paragraph(off.comment)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# STREAMLIT APP
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
st.caption("Upload → extraction → sélection (checkbox) → Word premium (charte Imagine).")
st.divider()

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
        catering_offers.append(parse_catering(txt, f.name))

    for f in tech_files:
        txt = extract_pdf_text(f)
        tech_offers.append(parse_tech(txt, f.name))

tab1, tab2 = st.tabs(["Traiteur (sélection)", "Technique (sélection)"])

with tab1:
    if not catering_offers:
        st.caption("Aucun devis traiteur.")
    else:
        st.caption("Coche uniquement ce qui doit apparaître (page 1 + détails).")
        for i, off in enumerate(catering_offers, start=1):
            with st.expander(f"{off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire (obligatoire)", value=off.vendor, key=f"c_vendor_{i}")
                ttc_in = st.text_input("Total TTC (hors options) — obligatoire", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"c_ttc_{i}")
                off.total_ttc = parse_eur_amount(ttc_in)

                if vendor_is_suspicious(off.vendor):
                    st.error("Nom prestataire invalide/suspect (titre interne, poste, contact ou adresse). Corrige-le.")
                if off.total_ttc is None:
                    st.warning("Total TTC non détecté — saisis-le.")

                for post in CATERING_POSTS:
                    cand = off.candidates.get(post, [])
                    if not cand:
                        continue
                    st.markdown(f"**{post}**")
                    sel = []
                    default_n = 8 if post in ["Déjeuner", "Cocktail"] else 5
                    for idx, line in enumerate(cand):
                        checked = st.checkbox(line, value=(idx < default_n), key=f"c_{i}_{post}_{idx}")
                        if checked:
                            sel.append(line)
                    off.selected[post] = sel

                off.comment = st.text_area("Commentaire (1–2 phrases max)", value=off.comment, height=80, key=f"c_comment_{i}")

with tab2:
    if not tech_offers:
        st.caption("Aucun devis technique.")
    else:
        for i, off in enumerate(tech_offers, start=1):
            with st.expander(f"{off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire (obligatoire)", value=off.vendor, key=f"t_vendor_{i}")
                ttc_in = st.text_input("Total TTC — obligatoire", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"t_ttc_{i}")
                off.total_ttc = parse_eur_amount(ttc_in)

                if vendor_is_suspicious(off.vendor):
                    st.warning("Nom prestataire à vérifier.")
                if off.total_ttc is None:
                    st.warning("Total TTC non détecté — saisis-le.")

                for post in TECH_POSTS:
                    cand = off.candidates.get(post, [])
                    if not cand:
                        continue
                    st.markdown(f"**{post}**")
                    sel = []
                    default_n = 10 if post in ["Périmètre", "Captation", "Régie", "Diffusion"] else 8
                    for idx, line in enumerate(cand):
                        checked = st.checkbox(line, value=(idx < default_n), key=f"t_{i}_{post}_{idx}")
                        if checked:
                            sel.append(line)
                    off.selected[post] = sel

                off.comment = st.text_area("Conseil (2–3 phrases max)", value=off.comment, height=90, key=f"t_comment_{i}")

# Gate generation
blocked = False
for off in catering_offers:
    if vendor_is_suspicious(off.vendor) or (off.total_ttc is None):
        blocked = True
for off in tech_offers:
    if (not norm(off.vendor)) or (off.total_ttc is None):
        blocked = True

st.divider()
colA, colB = st.columns([2, 1], vertical_alignment="center")
with colA:
    st.caption("Contrôle qualité : génération bloquée si vendor/Total TTC manquants. Le Word contient uniquement les éléments cochés.")
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
