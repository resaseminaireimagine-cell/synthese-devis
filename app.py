# app.py — Synthèse devis prestataires (ROBUSTE / STABLE)
# ------------------------------------------------------
# Objectif: ne JAMAIS planter + sortir un Word propre.
# - Vendor: scoring robuste + fallback filename
# - Décollage texte PDF (Les boissons froidesJus...) + segmentation
# - Filtrage bruit (CGV / RCS / TVA / IBAN / clauses / horaires)
# - Extraction par items -> classification par postes
# - UI "light": vendor + TTC + synthèse (optionnelle) + détail (modifiable)
#
# Dépendances:
#   pip install streamlit pypdf python-docx

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import streamlit as st
from pypdf import PdfReader

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================
# BRAND / CONFIG
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
BG = "#F6F7FB"
FONT = "Montserrat"

MAX_CATERING = 3
MAX_TECH = 2


# =========================
# POSTS
# =========================
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
# LEXICON
# =========================
MENU_KEEP_HINTS = [
    "accueil", "petit", "déjeuner", "dejeuner", "pause", "buffet", "cocktail", "apéritif", "aperitif",
    "déjeunatoire", "dejeunatoire",
    "café", "cafe", "thé", "the", "soft", "jus", "eau", "limonade", "citronnade",
    "viennoiser", "gourmand", "mignard", "financier", "cannel", "tartelette", "cheesecake", "brochette",
    "pièce", "pieces", "pièces", "/pers", "par personne", "convive", "invité", "invite",
    "salée", "salées", "sucrée", "sucrées", "sucree", "sucrees",
    "sandwich", "wrap", "salade", "fromage", "fruit", "charcut", "planche", "tapenade", "houmous",
    "vin", "champagne", "bouteille",
    "verrerie", "flûte", "flutes", "assiette", "nappage", "mobilier", "mange-debout",
    "service", "maître d’hôtel", "maitre d'hotel", "barman", "personnel",
    "livraison", "reprise", "mise en place", "debarrassage", "débarrassage",
]

TECH_KEEP_HINTS = [
    "captation", "caméra", "camera", "4k", "cadreur", "réalisateur", "realisateur",
    "ingénieur", "ingenieur", "son", "audio",
    "régie", "regie", "diffusion", "live", "zoom", "duplex", "plateforme",
    "replay", "wetransfer", "we transfer", "enregistrement",
    "pavlov", "zapette", "écran", "ecran", "tv", "moniteur",
    "micro", "hf", "retour", "mixette",
    "installation", "démontage", "demontage",
]

# Hard noise
NOISE_HINTS = [
    "conditions générales", "conditions generales", "cgv",
    "rgpd", "données personnelles", "donnees personnelles",
    "iban", "bic", "rib", "banque",
    "tva intracommunautaire", "tva :", "total tva", "montant tva", "base ht", "total ht",
    "siret", "rcs", "au capital", "capital social",
    "tribunal", "mise en demeure", "pénalité", "penalite", "recouvrement",
    "adresse", "email", "e-mail", "www.", "site internet",
    "référence", "reference", "date de devis", "date de validité", "signature", "bon pour accord",
    "mode de paiement", "net à payer", "net a payer",
    "désignation", "designation", "quantité", "quantite", "montant", "p.u", "pu ht",
    "indemnité forfaitaire", "indemnite forfaitaire",
    "clause", "résolutoire", "resolutoire", "déchéance", "decheance",
    "le client", "le vendeur", "l’acheteur", "l'acheteur",
    "diminution du nombre", "augmentation du nombre",
    "se dégage", "se degage", "dédommagement", "dedommagement",
    "page ",
]

VENDOR_FORBIDDEN = [
    "si vous", "merci de", "nous retourner", "en accord", "proposition",
    "r.c.s", "rcs", "siret", "tva", "iban", "bic", "rib",
    "prix par convive", "en euros", "appliqué", "applique",
    "devis", "facture", "total", "montant", "désignation", "designation",
]

COMPANY_MARKERS = [
    "sas", "sarl", "sa", "eurl", "sasu", "association", "groupe", "production", "traiteur", "réceptions", "receptions"
]


# =========================
# TEXT UTILS
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").replace("\t", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fold(s: str) -> str:
    return norm(s).lower()


def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.replace("#", "").strip()
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def safe_extract_pdf_text(uploaded_file) -> str:
    """Never crash: return '' on any PDF weirdness."""
    try:
        reader = PdfReader(uploaded_file)
        chunks = []
        for p in reader.pages:
            try:
                chunks.append(p.extract_text() or "")
            except Exception:
                chunks.append("")
        return "\n".join(chunks)
    except Exception:
        return ""


def decollage(text: str) -> str:
    """
    PDF parfois colle des phrases sans espaces:
    'Les boissons froidesJus d’orangeEvian...'
    On insère des séparateurs sur transitions probables.
    """
    if not text:
        return ""
    t = text

    # Insert newline after punctuation when missing space
    t = re.sub(r"([:;,.!?])([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1\n\2", t)

    # Insert newline between lowercase/accent and Uppercase (incl. É, À...)
    t = re.sub(r"([a-zà-öø-ÿ])([A-ZÀ-ÖØ-Þ])", r"\1\n\2", t)

    # Insert newline between digit and Uppercase (e.g., "6piècesCocktail")
    t = re.sub(r"(\d)([A-ZÀ-ÖØ-Þ])", r"\1\n\2", t)

    # Bullets
    t = t.replace("•", "\n• ")

    # Common glued patterns (boissons/chaudes/froides etc.)
    t = re.sub(r"(boissons?\s+(?:chaudes|froides))([A-ZÀ-ÖØ-Þ])", r"\1\n\2", t, flags=re.I)
    t = re.sub(r"(cocktail|déjeuner|dejeuner|pause)([A-ZÀ-ÖØ-Þ])", r"\1\n\2", t, flags=re.I)

    return t


def split_lines(text: str) -> List[str]:
    text = decollage(text)
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
    digits = sum(ch.isdigit() for ch in core)
    letters = sum(ch.isalpha() for ch in core)
    # many digits, few letters -> table/prices
    if digits >= 10 and digits > letters:
        return True
    if re.fullmatch(r"[\d\s,\.%\-\/]+", core) and digits >= 6:
        return True
    return False


def looks_like_schedule_line(s: str) -> bool:
    l = fold(s)
    # "06h00 à 08h30 ..." or "8h 30 / 09h 00"
    if re.search(r"\b0?\d{1,2}h\s?\d{0,2}\b", l) and ((" à " in l) or (" a " in l) or ("/" in l)):
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
    if looks_like_schedule_line(s) and len(s) <= 140:
        return True
    if any(k in l for k in NOISE_HINTS):
        return True
    # pure address-ish
    if re.search(r"\b(rue|avenue|boulevard|quai|impasse|route|cedex)\b", l) and sum(ch.isdigit() for ch in s) >= 2:
        return True
    # phone-like
    if re.search(r"\b0[1-9](\s?\d{2}){4}\b", s):
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
    lt = fold(text)
    patterns = [
        r"total\s+ttc\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+devis\s+t\.t\.c\.\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
        r"total\s+devis\s+ttc\s*[:\-]?\s*([0-9][0-9 \u00A0]*[,\.][0-9]{2})",
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
    junk = {"v1", "v2", "v3", "devis", "dev", "institut", "imagine", "pax", "100pax", "pdf"}
    toks = [t for t in base.split() if t.lower() not in junk]
    out = " ".join(toks[:6]).strip()
    return out or base


def vendor_score(line: str) -> int:
    """Higher is better."""
    s = norm(line)
    l = fold(s)
    score = 0

    # hard rejects
    if not s or len(s) < 4:
        return -10_000
    if any(b in l for b in VENDOR_FORBIDDEN):
        score -= 200
    if re.search(r"\bfr\d{5,}\b", l):  # VAT number like FR...
        score -= 250
    if re.search(r"\br\.?c\.?s\.?\b", l) or "rcs" in l:
        score -= 200
    if looks_like_price_table_line(s):
        score -= 200
    if looks_like_schedule_line(s):
        score -= 120
    if is_noise_line(s):
        score -= 250

    # positives
    if any(m in l for m in COMPANY_MARKERS):
        score += 120
    # looks like name (many letters)
    alpha = sum(ch.isalpha() for ch in s)
    if alpha >= 8:
        score += 30
    if alpha / max(len(s), 1) > 0.6:
        score += 30
    # uppercased brand
    if s.upper() == s and alpha >= 6:
        score += 60

    # negatives: sentences
    if re.search(r"\b(merci|veuillez|si vous|nous vous|proposition|accord)\b", l):
        score -= 180
    if len(s) > 75:
        score -= 30

    return score


def guess_vendor_name(text: str, filename: str) -> str:
    fallback = vendor_from_filename(filename)
    lines = [norm(x) for x in split_lines(text) if norm(x)]
    head = lines[:260]

    # candidates from top
    best = None
    best_sc = -10_000
    for ln in head:
        sc = vendor_score(ln)
        if sc > best_sc:
            best_sc = sc
            best = ln

    # sanitize
    if best and best_sc >= 40:
        cand = best
    else:
        cand = fallback

    cand = re.split(r"(?i)\bau\s+capital\b", cand)[0].strip(" -–,;:")
    cand = re.sub(r"\s+\d{1,2}$", "", cand).strip()
    if len(cand) > 60:
        cand = cand[:60].rstrip() + "…"
    return cand or fallback


# =========================
# EXTRACTION
# =========================
def extract_items(lines: List[str], keep_hints: List[str], max_len: int) -> List[str]:
    items = []
    seen = set()
    for ln in lines:
        s = norm(ln)
        if not s or is_noise_line(s):
            continue

        # bullets
        if s.startswith(("•", "-", "–")):
            it = s.lstrip("•-– ").strip()
            it = norm(it)
            if it and not is_noise_line(it):
                k = fold(it)
                if k not in seen:
                    seen.add(k)
                    items.append(it)
            continue

        l = fold(s)
        if len(s) <= max_len and any(k in l for k in keep_hints):
            if not looks_like_price_table_line(s):
                k = fold(s)
                if k not in seen:
                    seen.add(k)
                    items.append(s)

    return items


def classify_catering_item(s: str) -> str:
    l = fold(s)

    if "cocktail" in l or "apéritif" in l or "aperitif" in l or "pièce cocktail" in l or "pieces cocktail" in l:
        return "Cocktail"

    if any(k in l for k in ["déjeuner", "dejeuner", "buffet", "wrap", "sandwich", "déjeunatoire", "dejeunatoire"]):
        return "Déjeuner"

    if any(k in l for k in ["pause", "viennoiser", "mignard", "financier", "cannel", "tartelette", "cheesecake", "brochette de fruits"]):
        # si ça mentionne clairement l’après-midi
        if re.search(r"\b(14|15|16|17)h", l):
            return "Pause après-midi"
        return "Pause matin"

    if any(k in l for k in ["accueil", "thermos", "café", "cafe", "thé", "the"]):
        return "Accueil café"

    if any(k in l for k in ["vin", "champagne", "soft", "eau", "jus", "perrier", "evian"]):
        return "Boissons (global)"

    if "option" in l or "supplément" in l or "supplement" in l or "heure supplémentaire" in l:
        return "Options"

    if any(k in l for k in ["vaisselle", "verrerie", "nappage", "mobilier", "mange-debout", "livraison", "reprise", "service", "maître", "barman", "personnel"]):
        return "Autres (logistique)"

    return "Autres (logistique)"


def classify_tech_item(s: str) -> str:
    l = fold(s)

    # kill TVA ghosts early
    if "tva" in l or "total tva" in l:
        return "Périmètre"

    if "replay" in l or "wetransfer" in l or "we transfer" in l or "enregistrement" in l:
        return "Replay"
    if "zoom" in l or "live" in l or "diffusion" in l or "duplex" in l or "plateforme" in l:
        return "Diffusion"
    if "régie" in l or "regie" in l or "pavlov" in l or "zapette" in l or "écran" in l or "ecran" in l or "tv" in l:
        return "Régie"
    if "caméra" in l or "camera" in l or "captation" in l or "4k" in l:
        return "Captation"
    if any(k in l for k in ["cadreur", "réalisateur", "realisateur", "ingénieur", "ingenieur", "son", "audio"]):
        return "Équipe"
    if any(k in l for k in ["option", "forfait", "connexion", "contraint", "supplément"]):
        return "Contraintes / options"
    return "Périmètre"


def cocktail_summary(items: List[str]) -> str:
    """
    Résumé cocktail: "10 pièces/pers — 6 salées + 4 sucrées".
    Si pas trouvable, "Cocktail (voir détail)".
    """
    if not items:
        return "—"
    txt = " ".join(items)
    l = fold(txt)

    total = None
    sale = None
    sucre = None

    mt = re.search(r"\b(\d{1,2})\s*pi[eè]ces?\s*(par\s+personne|/pers|par\s+convive)", l)
    if mt:
        total = int(mt.group(1))

    ms = re.search(r"\b(\d{1,2})\s*pi[eè]ces?.{0,40}(sal[ée]es?)", l)
    if ms:
        sale = int(ms.group(1))

    mu = re.search(r"\b(\d{1,2})\s*pi[eè]ces?.{0,40}(sucr[ée]es?)", l)
    if mu:
        sucre = int(mu.group(1))

    # Some docs have "12 salées + 12 sucrées" (without "pièces")
    if sale is None:
        ms2 = re.search(r"\b(\d{1,2})\s*(?:sal[ée]es?)\b", l)
        if ms2:
            sale = int(ms2.group(1))
    if sucre is None:
        mu2 = re.search(r"\b(\d{1,2})\s*(?:sucr[ée]es?)\b", l)
        if mu2:
            sucre = int(mu2.group(1))

    if total or sale or sucre:
        bits = []
        if total:
            bits.append(f"{total} pièces/pers")
        sub = []
        if sale:
            sub.append(f"{sale} salées")
        if sucre:
            sub.append(f"{sucre} sucrées")
        if sub:
            bits.append(" + ".join(sub))
        if "option" in l and "sucr" in l:
            bits.append("(sucré en option)")
        return " — ".join(bits)

    return "Cocktail (voir détail)"


def short_summary(items: List[str], max_chars: int = 90) -> str:
    if not items:
        return "—"
    s = " / ".join(items[:2])
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > max_chars:
        s = s[: max_chars - 1].rstrip() + "…"
    return s or "—"


# =========================
# MODEL
# =========================
@dataclass
class CateringOffer:
    vendor: str
    total_ttc: Optional[float]
    posts: Dict[str, List[str]]
    summary_by_post: Dict[str, str]
    detail_text: str  # modifiable


@dataclass
class TechOffer:
    vendor: str
    total_ttc: Optional[float]
    posts: Dict[str, List[str]]
    summary_by_post: Dict[str, str]
    detail_text: str  # modifiable


def parse_catering_offer(text: str, filename: str) -> CateringOffer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    items = extract_items(filtered, MENU_KEEP_HINTS, max_len=260)

    posts: Dict[str, List[str]] = {p: [] for p in CATERING_POSTS}
    for it in items:
        posts[classify_catering_item(it)].append(it)

    # dedupe
    for p in posts:
        seen = set()
        uniq = []
        for it in posts[p]:
            k = fold(it)
            if k not in seen:
                seen.add(k)
                uniq.append(it)
        posts[p] = uniq

    summary: Dict[str, str] = {}
    for p in CATERING_POSTS:
        if p == "Cocktail":
            summary[p] = cocktail_summary(posts[p])
        else:
            summary[p] = short_summary(posts[p], 90)

    # default detail: structured per post
    detail_lines: List[str] = []
    for p in CATERING_POSTS:
        if posts[p]:
            detail_lines.append(f"{p}:")
            for it in posts[p]:
                detail_lines.append(f"• {it}")
            detail_lines.append("")
    detail_text = "\n".join(detail_lines).strip() or "—"

    return CateringOffer(vendor, total_ttc, posts, summary, detail_text)


def parse_tech_offer(text: str, filename: str) -> TechOffer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    items = extract_items(filtered, TECH_KEEP_HINTS, max_len=320)

    # Kill TVA-only noise inside extracted items
    items = [it for it in items if "tva" not in fold(it)]

    posts: Dict[str, List[str]] = {p: [] for p in TECH_POSTS}
    for it in items:
        posts[classify_tech_item(it)].append(it)

    for p in posts:
        seen = set()
        uniq = []
        for it in posts[p]:
            k = fold(it)
            if k not in seen:
                seen.add(k)
                uniq.append(it)
        posts[p] = uniq

    summary: Dict[str, str] = {}
    for p in TECH_POSTS:
        summary[p] = short_summary(posts[p], 95)

    detail_lines: List[str] = []
    for p in TECH_POSTS:
        if posts[p]:
            detail_lines.append(f"{p}:")
            for it in posts[p]:
                detail_lines.append(f"• {it}")
            detail_lines.append("")
    detail_text = "\n".join(detail_lines).strip() or "—"

    return TechOffer(vendor, total_ttc, posts, summary, detail_text)


# =========================
# WORD HELPERS
# =========================
def set_run(run, bold=False, size=10, color="#111827"):
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = hex_to_rgb(color)


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
               catering: List[CateringOffer], tech: List[TechOffer]) -> bytes:
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

    # ---- CATERING SUMMARY TABLE
    if catering:
        add_subtitle(doc, "1) PRESTATION TRAITEUR — Comparatif (synthèse)")
        vendors = catering[:MAX_CATERING]
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

        rows = ["Total TTC"] + [p for p in CATERING_POSTS if p != "Autres (logistique)"] + ["Commentaire"]
        for label in rows:
            r = table.add_row().cells
            r[0].text = label
            set_cell_shading(r[0], "F3F4F6")
            for rr in r[0].paragraphs[0].runs:
                set_run(rr, bold=True, size=9, color="#111827")
            set_cell_margins(r[0])

            for j, off in enumerate(vendors, start=1):
                if label == "Total TTC":
                    val = euro_fmt(off.total_ttc)
                elif label == "Commentaire":
                    val = "—"
                else:
                    val = off.summary_by_post.get(label, "—")
                r[j].text = val if val else "—"
                for pp in r[j].paragraphs:
                    for rr in pp.runs:
                        set_run(rr, bold=False, size=9, color="#111827")
                set_cell_margins(r[j])

        doc.add_paragraph("")

    # ---- TECH SUMMARY
    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        for idx, off in enumerate(tech[:MAX_TECH], start=1):
            p = doc.add_paragraph()
            rr = p.add_run(f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            set_run(rr, bold=True, size=9, color="#111827")

            t = doc.add_table(rows=1, cols=2)
            h = t.rows[0].cells
            h[0].text = "Item"
            h[1].text = "Synthèse"
            set_cell_shading(h[0], PRIMARY)
            set_cell_shading(h[1], PRIMARY)
            for cell in h:
                for rrr in cell.paragraphs[0].runs:
                    set_run(rrr, bold=True, size=10, color="#FFFFFF")
                set_cell_margins(cell)

            for item in TECH_POSTS:
                rrrow = t.add_row().cells
                rrrow[0].text = item
                rrrow[1].text = off.summary_by_post.get(item, "—")
                set_cell_shading(rrrow[0], "F3F4F6")
                for rrr in rrrow[0].paragraphs[0].runs:
                    set_run(rrr, bold=True, size=9, color="#111827")
                for cell in rrrow:
                    set_cell_margins(cell)
                    for pp in cell.paragraphs:
                        for rrr in pp.runs:
                            set_run(rrr, bold=False, size=9, color="#111827")
            doc.add_paragraph("")

    # ---- DETAILS
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (modifiable via l’outil)")
    add_small(doc, "Le comparatif ci-dessus est une synthèse ; ci-dessous la liste des contenus.")
    doc.add_paragraph("")  # no extra blank page; just spacing

    if catering:
        add_title(doc, "DÉTAIL — PRESTATIONS TRAITEUR")
        doc.add_paragraph("")
        for off in catering[:MAX_CATERING]:
            add_band(doc, off.vendor, f"Total TTC : {euro_fmt(off.total_ttc)}")
            doc.add_paragraph(off.detail_text.strip() or "—")

    if tech:
        doc.add_page_break()
        add_title(doc, "DÉTAIL — PRESTATIONS TECHNIQUES")
        doc.add_paragraph("")
        for off in tech[:MAX_TECH]:
            add_band(doc, off.vendor, f"Total TTC : {euro_fmt(off.total_ttc)}")
            doc.add_paragraph(off.detail_text.strip() or "—")

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
st.caption("Robuste : extraction auto + correction manuelle possible + Word toujours générable.")
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

catering_files = (catering_files or [])[:MAX_CATERING]
tech_files = (tech_files or [])[:MAX_TECH]

with st.spinner("Lecture des PDFs…"):
    catering_offers: List[CateringOffer] = []
    tech_offers: List[TechOffer] = []

    for f in catering_files:
        txt = safe_extract_pdf_text(f)
        catering_offers.append(parse_catering_offer(txt, f.name))

    for f in tech_files:
        txt = safe_extract_pdf_text(f)
        tech_offers.append(parse_tech_offer(txt, f.name))

st.divider()
st.subheader("Édition (light)")

if catering_offers:
    st.markdown("#### Traiteur")
    for i, off in enumerate(catering_offers, start=1):
        with st.expander(f"Traiteur {i} — {off.vendor}", expanded=(i == 1)):
            off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"c_vendor_{i}")

            ttc_in = st.text_input(
                "Total TTC",
                value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)),
                key=f"c_ttc_{i}",
            )
            off.total_ttc = parse_eur_amount(ttc_in)

            synth_default = "\n".join(
                [f"{p}: {off.summary_by_post.get(p,'—')}" for p in ["Cocktail", "Déjeuner", "Pause matin", "Accueil café", "Boissons (global)", "Options"]]
            )
            synth = st.text_area("Synthèse (par postes) — optionnel", value=synth_default, height=140, key=f"c_synth_{i}")

            if norm(synth) != norm(synth_default):
                for line in synth.splitlines():
                    if ":" in line:
                        k, v = line.split(":", 1)
                        k = norm(k)
                        v = norm(v) or "—"
                        if k in off.summary_by_post:
                            off.summary_by_post[k] = v

            off.detail_text = st.text_area("Détail (sert pour la section DÉTAIL du Word)", value=off.detail_text, height=260, key=f"c_detail_{i}")

if tech_offers:
    st.markdown("#### Technique")
    for i, off in enumerate(tech_offers, start=1):
        with st.expander(f"Technique {i} — {off.vendor}", expanded=(i == 1 and not catering_offers)):
            off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"t_vendor_{i}")

            ttc_in = st.text_input(
                "Total TTC",
                value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)),
                key=f"t_ttc_{i}",
            )
            off.total_ttc = parse_eur_amount(ttc_in)

            synth_default = "\n".join([f"{p}: {off.summary_by_post.get(p,'—')}" for p in TECH_POSTS])
            synth = st.text_area("Synthèse (par postes) — optionnel", value=synth_default, height=160, key=f"t_synth_{i}")

            if norm(synth) != norm(synth_default):
                for line in synth.splitlines():
                    if ":" in line:
                        k, v = line.split(":", 1)
                        k = norm(k)
                        v = norm(v) or "—"
                        if k in off.summary_by_post:
                            off.summary_by_post[k] = v

            off.detail_text = st.text_area("Détail (sert pour la section DÉTAIL du Word)", value=off.detail_text, height=260, key=f"t_detail_{i}")

st.divider()
if st.button("Générer le Word (.docx)", use_container_width=True, type="primary"):
    try:
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
    except Exception as e:
        st.error("Erreur lors de la génération Word (mais l’app n’a pas planté).")
        st.code(str(e))
