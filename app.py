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
# BRAND
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
BG = "#F6F7FB"
FONT = "Montserrat"  # Word substitue si non installée
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
    "café", "cafe", "thé", "the", "soft", "jus", "eau",
    "viennoiser", "gourmand", "mignard",
    "pièce", "pieces", "pièces", "/pers", "par personne", "convive", "invité", "invite",
    "salée", "sucrée", "dessert",
    "sandwich", "wrap", "salade", "fromage", "fruit",
    "vin", "champagne", "bière", "biere",
    "thermos", "gobelet", "tasse", "serviette", "plateau",
]
TECH_KEEP_HINTS = [
    "captation", "caméra", "camera", "4k", "cadreur", "réalisateur", "realisateur",
    "ingénieur", "ingenieur", "son", "audio",
    "régie", "regie", "diffusion", "live", "zoom", "duplex", "plateforme",
    "replay", "wetransfer", "we transfer", "enregistrement",
    "pavlov", "zapette", "tv", "écran", "ecran", "écrans", "ecrans",
    "micro", "hf", "console", "mélangeur", "melangeur", "obs", "vmix",
]
VENDOR_FORBIDDEN = [
    "accueil", "pause", "déjeuner", "dejeuner", "buffet", "cocktail", "boissons", "options",
    "personnel", "service", "scénographie", "scenographie",
    "récapitulatif", "recapitulatif", "sur la base", "hors options", "budget", "déroulé", "deroule",
    "détail", "detail", "proposition", "prestation", "inclus", "option",
    "en euros", "devis", "facture", "total",
]
VENDOR_HARD_FORBIDDEN = [
    "déroulé", "deroule", "rangement", "départ", "depart", "fin de la prestation",
    "reprise", "arrivée", "arrivee", "livraison", "horaire", "planning",
    "installation", "désinstallation", "desinstallation",
    "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche",
]
ADDRESS_HINTS = [
    "rue", "avenue", "boulevard", "allée", "allee", "bp", "cedex",
    "france", "paris", "lyon", "marseille", "clichy", "nanterre", "saint",
    "750", "751", "752", "753", "754", "755", "756", "757", "758", "759",
]
NOISE_HINTS = [
    "conditions générales", "conditions generales", "cgv", "rgpd", "données personnelles", "donnees personnelles",
    "propriété intellectuelle", "propriete intellectuelle", "droit à l'image", "droit a l'image",
    "iban", "bic", "rib", "banque", "capital", "tva intracommunautaire",
    "pénalité", "penalite", "recouvrement", "mise en demeure", "tribunal",
    "responsabilité", "responsabilite", "dommages", "intérêts", "interets",
    "déchéance", "decheance", "résolutoire", "resolutoire", "litige", "contestation",
    "adresse", "tél", "tel", "email", "e-mail", "www.", "site internet",
    "référence", "reference", "devis n", "date de devis", "date de validité", "signature",
    "mode de paiement", "facture",
    "net a payer", "net à payer",
    "désignation", "designation", "quantité", "quantite", "p.u", "pu ht", "montant", "remise", "taux", "qté", "qte", "réf", "ref",
    "base ht", "total ht", "total tva", "tva :", "tva", "page ",
    "le client", "le vendeur", "le preneur",
    "annulation de la commande", "augmentation du nombre", "diminution du nombre",
    "sans que", "dans les délais", "dans les delais",
]
# =========================
# TEXT UTILS
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
    return "\n".join([(p.extract_text() or "") for p in reader.pages])
def split_lines(text: str) -> List[str]:
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
def cut_at_cgv(lines: List[str]) -> List[str]:
    out = []
    for ln in lines:
        l = fold(ln)
        if any(k in l for k in ["conditions générales", "conditions generales", "cgv"]):
            break
        out.append(ln)
    return out
def looks_like_schedule_line(s: str) -> bool:
    l = fold(s)
    if re.search(r"\b0?\d{1,2}h\d{2}\b", l) and ((" à " in l) or (" a " in l) or ("-" in l)):
        return len(l) <= 170
    return False
def looks_like_placeholder_line(s: str) -> bool:
    ss = norm(s)
    if re.match(r"^\s*:\s*[A-Za-zÀ-ÖØ-öø-ÿ].{0,120}:\s*", ss):
        return True
    if ss.strip().startswith(":") and re.search(r"\b(pax|nb pax|nombre de pax)\b", fold(ss)):
        return True
    return False
def looks_like_price_table_line(s: str) -> bool:
    core = re.sub(r"[€]", "", norm(s))
    digits = sum(ch.isdigit() for ch in core)
    letters = sum(ch.isalpha() for ch in core)
    if digits >= 10 and digits > letters:
        return True
    if re.fullmatch(r"[\d\s,\.%\-\/]+", core) and digits >= 6:
        return True
    return False
def is_noise_line(s: str) -> bool:
    s = norm(s)
    if not s:
        return True
    l = fold(s)
    if looks_like_placeholder_line(s):
        return True
    if re.fullmatch(r"\d{1,5}", s):
        return True
    if looks_like_schedule_line(s):
        return True
    if any(k in l for k in NOISE_HINTS):
        return True
    if "institut imagine" in l and ("étage" in l or "etage" in l or "sur la base" in l):
        return True
    if re.fullmatch(r"(de\s+)?\d{1,2}h(\d{2})?", l):
        return True
    return False
def unglue(s: str) -> str:
    s = norm(s)
    s = re.sub(r"(personne|convive|invité|invite)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 • \2", s, flags=re.I)
    s = re.sub(r"([:;])([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2", s)
    return s
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
    found: List[float] = []
    for pat in patterns:
        for m in re.finditer(pat, lt, flags=re.IGNORECASE | re.DOTALL):
            amt = parse_eur_amount(m.group(1))
            if amt is not None:
                found.append(amt)
    return found[-1] if found else None
# =========================
# TABULAR LABEL EXTRACTION
# =========================
def extract_left_label_from_tabular(line: str) -> Optional[str]:
    s = norm(line)
    if len(s) < 6:
        return None
    if looks_like_placeholder_line(s):
        return None
    parts = s.split()
    cut = None
    for i, tok in enumerate(parts):
        if any(ch.isdigit() for ch in tok):
            cut = i
            break
    if cut is None or cut == 0:
        return None
    left = " ".join(parts[:cut]).strip(" -–:;")
    if len(left) < 4:
        return None
    ll = fold(left)
    if any(k in ll for k in ADDRESS_HINTS):
        return None
    if any(k in ll for k in ["total", "tva", "montant", "remise", "désignation", "designation", "quantité", "quantite", "page"]):
        return None
    return left
def looks_like_menu_item_line(s: str) -> bool:
    s = norm(s)
    l = fold(s)
    if len(s) < 4 or len(s) > 200:
        return False
    if looks_like_price_table_line(s):
        return bool(extract_left_label_from_tabular(s))
    if any(k in l for k in ["total", "tva", "iban", "siret", "cgv", "référence", "reference"]):
        return False
    if re.search(r"\b(g|gr|kg|ml|cl|l)\b", l):
        return True
    if re.search(r"\b(filet|tarte|salade|quiche|risotto|fromage|dessert|mignard|verrine|canapé|canape|wrap|sandwich)\b", l):
        return True
    words = [w for w in re.split(r"\s+", s) if w]
    if 2 <= len(words) <= 12 and sum(w[:1].isupper() for w in words if w[:1].isalpha()) >= 2:
        return True
    return False
def extract_items(lines: List[str], keep_hints: List[str], max_len: int, relax: bool = False) -> List[str]:
    items: List[str] = []
    for ln in lines:
        s = norm(ln)
        if not s or is_noise_line(s):
            continue
        if s.startswith(("•", "-", "–")):
            it = s.lstrip("•-– ").strip()
            if it and not is_noise_line(it):
                items.append(unglue(it))
            continue
        if looks_like_price_table_line(s):
            lab = extract_left_label_from_tabular(s)
            if lab:
                items.append(unglue(lab))
            continue
        l = fold(s)
        if len(s) <= max_len:
            if any(k in l for k in keep_hints):
                items.append(unglue(s))
                continue
            if relax and looks_like_menu_item_line(s):
                items.append(unglue(s))
                continue
    out, seen = [], set()
    for it in items:
        k = fold(it)
        if k in seen:
            continue
        seen.add(k)
        out.append(it)
    return out
# =========================
# VENDOR (amélioré)
# =========================
def vendor_is_suspicious(v: str) -> bool:
    v = norm(v)
    lv = fold(v)
    # identifiants admin
    if re.search(r"\b(r\.?c\.?s\.?|siret|tva|tva intracommunautaire)\b", lv):
        return True
    compact = re.sub(r"\s+", "", v)
    if re.fullmatch(r"FR\d{8,14}", compact):
        return True
    if re.fullmatch(r"\d{8,18}", compact):
        return True
    if any(k in lv for k in VENDOR_HARD_FORBIDDEN):
        return True
    if any(k in lv for k in VENDOR_FORBIDDEN):
        return True
    # adresse probable: trop de signaux
    if sum(k in lv for k in ["rue", "avenue", "boulevard", "cedex", "france", "paris"]) >= 2:
        return True
    # commence par ":"
    if v.strip().startswith(":"):
        return True
    if "@" in v or "contact" in lv:
        return True
    # trop court
    if len(v) < 3:
        return True
    return False
def guess_vendor_name(text: str, filename: str) -> str:
    lines = [norm(x) for x in text.splitlines() if norm(x)]
    top = lines[:240]
    # 1) ancrage: au-dessus de SIRET/RCS/TVA intracom
    for idx, ln in enumerate(top):
        l = fold(ln)
        if "siret" in l or "r.c.s" in l or "rcs" in l or "tva intracommunautaire" in l:
            for back in range(1, 13):
                j = idx - back
                if j >= 0:
                    cand = top[j].strip()
                    if not cand or is_noise_line(cand) or vendor_is_suspicious(cand):
                        continue
                    # signal entreprise
                    if re.search(r"\b(sas|sarl|sa|groupe|concept|production|traiteur)\b", fold(cand)) or sum(ch.isalpha() for ch in cand) >= 8:
                        return cand
            break
    # 2) fallback: meilleure ligne "raison sociale-ish"
    bad = re.compile(r"\b(devis|facture|date|total|tva|iban|bic|net a payer|net à payer)\b", re.I)
    candidates = []
    for ln in top:
        if len(ln) < 4 or len(ln) > 70:
            continue
        if bad.search(ln):
            continue
        if is_noise_line(ln):
            continue
        if vendor_is_suspicious(ln):
            continue
        alpha = sum(ch.isalpha() for ch in ln)
        if alpha < 6:
            continue
        upper_ratio = sum(ch.isupper() for ch in ln if ch.isalpha()) / max(1, alpha)
        if ln.upper() == ln or upper_ratio > 0.60:
            candidates.append(ln)
    if candidates:
        prefer = ["traiteur", "réceptions", "receptions", "production", "concept", "unik", "cercle", "cadet", "exupery", "ll"]
        best, best_score = None, -10**9
        for c in candidates:
            fc = fold(c)
            score = 0
            for k in prefer:
                if k in fc:
                    score += 6
            if sum(k in fc for k in ["rue", "avenue", "boulevard", "cedex"]) >= 1:
                score -= 6
            if score > best_score:
                best_score, best = score, c
        if best and not vendor_is_suspicious(best):
            return best
    # 3) filename fallback nettoyé
    base = filename.rsplit(".", 1)[0]
    base = re.sub(r"\s+", " ", base).strip()
    return base
# =========================
# ROUTING
# =========================
def route_catering_item(item: str) -> str:
    l = fold(item)
    if re.search(r"\b(14h|15h|16h|17h|18h)\b", l):
        if "cocktail" in l:
            return "Cocktail"
        if "pause" in l or "gourmand" in l:
            return "Pause après-midi"
        return "Pause après-midi"
    if re.search(r"\b(8h|9h|10h|11h|12h)\b", l):
        if "déjeuner" in l or "dejeuner" in l or "buffet" in l:
            return "Déjeuner"
        if "pause" in l:
            return "Pause matin"
        return "Accueil café"
    if any(k in l for k in ["accueil café", "accueil cafe", "café", "cafe", "thé", "the", "thermos", "jus", "lait", "sucre", "tasse", "gobelet"]):
        return "Accueil café"
    if any(k in l for k in ["pause", "viennoiser", "gourmand", "mignard", "financier", "cannel", "cookie"]):
        return "Pause matin"
    if any(k in l for k in ["déjeuner", "dejeuner", "buffet", "sandwich", "wrap", "salade", "plat", "entrée", "entree", "dessert"]):
        return "Déjeuner"
    if any(k in l for k in ["cocktail", "apéritif", "aperitif", "pièce", "pieces", "canapé", "canape", "verrine", "mini"]):
        return "Cocktail"
    if any(k in l for k in ["vin", "champagne", "bière", "biere", "soft", "jus", "eau"]):
        return "Boissons (global)"
    if any(k in l for k in ["option", "en option", "supplément", "supplement"]):
        return "Options"
    if any(k in l for k in ["livraison", "mise en place", "service", "personnel", "vaisselle", "nappage", "location", "frais"]):
        return "Autres (logistique)"
    return "Autres (logistique)"
def route_tech_item(item: str) -> str:
    l = fold(item)
    if any(k in l for k in ["réalisateur", "realisateur", "cadreur", "ingénieur", "ingenieur", "son", "technicien"]):
        return "Équipe"
    if any(k in l for k in ["caméra", "camera", "4k", "objectif", "pied", "captation"]):
        return "Captation"
    if any(k in l for k in ["régie", "regie", "mélangeur", "melangeur", "obs", "vmix", "console", "écran", "ecran", "tv", "pavlov", "zapette"]):
        return "Régie"
    if any(k in l for k in ["zoom", "live", "duplex", "diffusion", "plateforme", "stream"]):
        return "Diffusion"
    if any(k in l for k in ["replay", "wetransfer", "we transfer", "enregistrement", "export"]):
        return "Replay"
    if any(k in l for k in ["inclus", "comprend", "incluant"]):
        return "Inclus"
    if any(k in l for k in ["option", "supplément", "supplement", "contrainte"]):
        return "Contraintes / options"
    return "Périmètre"
# =========================
# MODEL
# =========================
@dataclass
class CateringOffer:
    vendor: str
    total_ttc: Optional[float]
    posts: Dict[str, List[str]]
    comment: str
@dataclass
class TechOffer:
    vendor: str
    total_ttc: Optional[float]
    posts: Dict[str, List[str]]
    comment: str
def parse_catering_offer(text: str, filename: str) -> Tuple[CateringOffer, Dict]:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)
    lines = cut_at_cgv(split_lines(text))
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    # collecte globale (plus fiable que sections dans tes devis)
    collected = extract_items(filtered, MENU_KEEP_HINTS, max_len=260, relax=True)
    posts: Dict[str, List[str]] = {p: [] for p in CATERING_POSTS}
    for it in collected:
        posts[route_catering_item(it)].append(it)
    # dedup
    for p in posts:
        out, seen = [], set()
        for it in posts[p]:
            k = fold(it)
            if k in seen:
                continue
            seen.add(k)
            out.append(it)
        posts[p] = out
    debug = {"vendor": vendor, "total_ttc": total_ttc, "n_items": len(collected)}
    return CateringOffer(vendor=vendor, total_ttc=total_ttc, posts=posts, comment=""), debug
def parse_tech_offer(text: str, filename: str) -> Tuple[TechOffer, Dict]:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)
    lines = cut_at_cgv(split_lines(text))
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    collected = extract_items(filtered, TECH_KEEP_HINTS, max_len=360, relax=False)
    posts: Dict[str, List[str]] = {p: [] for p in TECH_POSTS}
    for it in collected:
        posts[route_tech_item(it)].append(it)
    # dedup
    for p in posts:
        out, seen = [], set()
        for it in posts[p]:
            k = fold(it)
            if k in seen:
                continue
            seen.add(k)
            out.append(it)
        posts[p] = out
    debug = {"vendor": vendor, "total_ttc": total_ttc, "n_items": len(collected)}
    return TechOffer(vendor=vendor, total_ttc=total_ttc, posts=posts, comment=""), debug
# =========================
# WORD HELPERS
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
def add_band(doc: Document, label: str, sub: str = ""):
    # bandeau pleine largeur via tableau 1x1
    t = doc.add_table(rows=1, cols=1)
    cell = t.rows[0].cells[0]
    set_cell_shading(cell, PRIMARY)
    set_cell_margins(cell, top=120, bottom=120, start=140, end=140)
    p = cell.paragraphs[0]
    r = p.add_run(label)
    set_run(r, bold=True, size=11, color="#FFFFFF")
    if sub:
        p2 = cell.add_paragraph()
        r2 = p2.add_run(sub)
        set_run(r2, bold=False, size=9, color="#FFFFFF")
    doc.add_paragraph("")
def summarize_for_table(items: List[str], max_chars: int) -> str:
    if not items:
        return ""
    s = " • ".join(items)
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) <= max_chars:
        return s
    return s[: max_chars - 5].rstrip() + " (...)"
def add_offer_detail_block(doc: Document, title: str, total: str, posts: Dict[str, List[str]], order: List[str]):
    add_band(doc, title, f"Total TTC : {total}")
    # table poste -> items
    t = doc.add_table(rows=1, cols=2)
    h = t.rows[0].cells
    h[0].text = "Poste"
    h[1].text = "Détail"
    set_cell_shading(h[0], PRIMARY)
    set_cell_shading(h[1], PRIMARY)
    for cell in h:
        for rr in cell.paragraphs[0].runs:
            set_run(rr, bold=True, size=10, color="#FFFFFF")
        set_cell_margins(cell)
    for post in order:
        rr = t.add_row().cells
        rr[0].text = post
        set_cell_shading(rr[0], "F3F4F6")
        for run in rr[0].paragraphs[0].runs:
            set_run(run, bold=True, size=9, color="#111827")
        set_cell_margins(rr[0])
        items = posts.get(post, [])
        rr[1].text = "—" if not items else "\n".join(f"• {x}" for x in items)
        for p in rr[1].paragraphs:
            for run in p.runs:
                set_run(run, bold=False, size=9, color="#111827")
        set_cell_margins(rr[1])
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
    # -------- Synthèse Traiteur --------
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
            ("Total TTC", 0),
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
                if label == "Total TTC":
                    val = euro_fmt(off.total_ttc)
                elif label == "Commentaire":
                    val = norm(off.comment) or "—"
                else:
                    val = summarize_for_table(off.posts.get(label, []), maxc) or "—"
                r[j].text = val
                for p in r[j].paragraphs:
                    for rr in p.runs:
                        set_run(rr, bold=False, size=9, color="#111827")
                set_cell_margins(r[j])
        doc.add_paragraph("")
    # -------- Synthèse Technique --------
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
            for item in TECH_POSTS:
                rrrow = t.add_row().cells
                rrrow[0].text = item
                set_cell_shading(rrrow[0], "F3F4F6")
                for rrr in rrrow[0].paragraphs[0].runs:
                    set_run(rrr, bold=True, size=9, color="#111827")
                set_cell_margins(rrrow[0])
                val = norm(off.comment) if item == "Conseil" else "\n".join(off.posts.get(item, []))
                rrrow[1].text = val if val else "—"
                set_cell_margins(rrrow[1])
                for p in rrrow[1].paragraphs:
                    for rrr in p.runs:
                        set_run(rrr, bold=False, size=9, color="#111827")
            doc.add_paragraph("")
    # =========================
    # DÉTAILS (séparations + édition)
    # =========================
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (modifiable via l’outil)")
    add_small(doc, "Contenu issu de l’extraction + corrections manuelles réalisées dans l’interface.")
    # -- Détail traiteur sur sa propre page --
    if catering:
        doc.add_page_break()
        add_title(doc, "DÉTAIL — PRESTATIONS TRAITEUR")
        add_small(doc, "Séparation par prestataire.")
        doc.add_paragraph("")
        for off in catering[:3]:
            add_offer_detail_block(
                doc,
                title=off.vendor,
                total=euro_fmt(off.total_ttc),
                posts=off.posts,
                order=CATERING_POSTS,
            )
    # -- Détail technique sur sa propre page --
    if tech:
        doc.add_page_break()
        add_title(doc, "DÉTAIL — PRESTATIONS TECHNIQUES")
        add_small(doc, "Séparation par prestataire.")
        doc.add_paragraph("")
        for off in tech[:2]:
            add_offer_detail_block(
                doc,
                title=off.vendor,
                total=euro_fmt(off.total_ttc),
                posts=off.posts,
                order=TECH_POSTS,
            )
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
st.caption("Pré-rempli automatique (traiteur + technique) → tu corriges (y compris les détails) → export Word.")
st.divider()
debug_mode = st.checkbox("Mode debug (voir vendor/ttc/items)", value=False)
ttc_min = st.number_input("Seuil TTC minimum (alerte) — pas bloquant", min_value=0, max_value=200000, value=500, step=50)
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
    catering_offers: List[CateringOffer] = []
    tech_offers: List[TechOffer] = []
    debug_dump = {"catering": [], "tech": []}
    for f in catering_files:
        txt = extract_pdf_text(f)
        offer, dbg = parse_catering_offer(txt, f.name)
        catering_offers.append(offer)
        debug_dump["catering"].append({"file": f.name, **dbg})
    for f in tech_files:
        txt = extract_pdf_text(f)
        offer, dbg = parse_tech_offer(txt, f.name)
        tech_offers.append(offer)
        debug_dump["tech"].append({"file": f.name, **dbg})
if debug_mode:
    st.subheader("DEBUG")
    st.json(debug_dump)
tab1, tab2, tab3 = st.tabs(["Synthèse (édition)", "Détail (édition)", "Technique (édition)"])
# -------- Synthèse traiteur (édition) --------
with tab1:
    if not catering_offers:
        st.caption("Aucun devis traiteur.")
    else:
        for i, off in enumerate(catering_offers, start=1):
            with st.expander(f"Traiteur {i} — {off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"c_vendor_{i}_s")
                if vendor_is_suspicious(off.vendor):
                    st.warning("Nom prestataire probablement faux (adresse / en-tête / RCS/TVA). Corrige-le.")
                ttc_in = st.text_input("Total TTC", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"c_ttc_{i}_s")
                off.total_ttc = parse_eur_amount(ttc_in)
                if off.total_ttc is not None and off.total_ttc < float(ttc_min):
                    st.warning(f"TTC < {ttc_min}€ : probable mauvaise détection (à vérifier).")
                colL, colR = st.columns(2)
                with colL:
                    for post in ["Accueil café", "Pause matin", "Déjeuner", "Pause après-midi"]:
                        edited = st.text_area(post, value="\n".join(off.posts.get(post, [])), height=150, key=f"c_sum_{i}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                with colR:
                    for post in ["Cocktail", "Boissons (global)", "Options", "Autres (logistique)"]:
                        edited = st.text_area(post, value="\n".join(off.posts.get(post, [])), height=150, key=f"c_sum_{i}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                off.comment = st.text_area("Commentaire", value=off.comment, height=80, key=f"c_comment_{i}")
# -------- Détail traiteur (édition dédiée) --------
with tab2:
    if not catering_offers:
        st.caption("Aucun devis traiteur.")
    else:
        st.info("Ici tu modifies ce qui alimentera la section “DÉTAIL DES OFFRES” du Word.")
        for i, off in enumerate(catering_offers, start=1):
            with st.expander(f"Détail — Traiteur {i} : {off.vendor}", expanded=False):
                # NOTE: Avoid key collision by NOT repeating the vendor input with same key
                # Just show the name or allow edit if needed, but watch out for session state issues
                for post in CATERING_POSTS:
                    edited = st.text_area(
                        f"{post} (détail)",
                        value="\n".join(off.posts.get(post, [])),
                        height=160,
                        key=f"c_det_{i}_{post}",
                    )
                    off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
# -------- Technique (édition) --------
with tab3:
    if not tech_offers:
        st.caption("Aucun devis technique.")
    else:
        for i, off in enumerate(tech_offers, start=1):
            with st.expander(f"Technique {i} — {off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire", value=off.vendor, key=f"t_vendor_{i}")
                if vendor_is_suspicious(off.vendor):
                    st.warning("Nom prestataire probablement faux (adresse / en-tête / TVA/RCS). Corrige-le.")
                ttc_in = st.text_input("Total TTC", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"t_ttc_{i}")
                off.total_ttc = parse_eur_amount(ttc_in)
                if off.total_ttc is not None and off.total_ttc < float(ttc_min):
                    st.warning(f"TTC < {ttc_min}€ : probable mauvaise détection (à vérifier).")
                col1, col2 = st.columns(2)
                with col1:
                    for post in ["Périmètre", "Équipe", "Captation", "Régie"]:
                        edited = st.text_area(post, value="\n".join(off.posts.get(post, [])), height=130, key=f"t_{i}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                with col2:
                    for post in ["Diffusion", "Replay", "Inclus", "Contraintes / options"]:
                        edited = st.text_area(post, value="\n".join(off.posts.get(post, [])), height=130, key=f"t_{i}_{post}")
                        off.posts[post] = [norm(x) for x in edited.splitlines() if norm(x)]
                off.comment = st.text_area("Conseil", value=off.comment, height=90, key=f"t_comment_{i}")
st.divider()
if "docx_file" not in st.session_state:
    st.session_state["docx_file"] = None
if st.button("Générer le document Word", use_container_width=True, type="primary"):
    data = build_word(
        event_title=event_title.strip() or "Événement",
        event_date=event_date.strip() or "Date à préciser",
        guests=int(guests),
        catering=catering_offers,
        tech=tech_offers,
    )
    st.session_state["docx_file"] = data
    st.rerun()
if st.session_state["docx_file"]:
    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    st.download_button(
        label="⬇️ Télécharger la synthèse (Word)",
        data=st.session_state["docx_file"],
        file_name=f"synthese_devis_{ts}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
