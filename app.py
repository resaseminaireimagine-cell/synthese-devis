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
# BRAND
# =========================
APP_TITLE = "Synthèse devis prestataires — Institut Imagine"
PRIMARY = "#AF0073"
BG = "#F6F7FB"
FONT = "Montserrat"  # Word utilisera Montserrat si dispo côté machine; sinon substitution.

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

MENU_HINTS = [
    "café", "cafe", "thé", "the", "soft", "jus", "eau", "viennoiser",
    "pause", "déjeuner", "dejeuner", "buffet", "cocktail", "apéritif", "aperitif",
    "pièce", "pieces", "pièces", "par personne", "/pers", "convive", "invité",
    "salée", "sucrée", "dessert", "sandwich", "wrap", "buns", "cake", "salade",
    "fromage", "fruit", "mignard", "cannel", "financier", "tartelette", "sphère",
    "saumon", "gambas", "volaille", "risotto", "briochin", "bresaola", "chèvre",
    "vin", "champagne",
]

LEGAL_KILL = [
    "clause", "mise en demeure", "tribunal", "responsabilité", "dommages", "intérêts",
    "déchéance", "résolutoire", "infructueuse", "litige", "contestat", "dédommagement",
    "rgpd", "données personnelles", "propriété intellectuelle", "droit à l'image",
]

ADMIN_NOISE_HINTS = [
    "conditions générales", "cgv", "siret", "rcs", "tva", "iban", "bic", "rib", "banque",
    "capital", "pénalité", "recouvrement", "force majeure",
    "adresse", "tél", "tel", "email", "e-mail", "www.",
    "référence", "devis n", "date de devis", "date de validité", "signature",
    "mode de paiement", "facture", "net a payer", "net à payer",
    "désignation", "quantité", "p.u", "pu ht", "montant", "remise", "taux", "qté", "réf", "ref",
    "base ht", "total ht", "total ttc", "page ",
    "société générale", "caisse d'epargne", "caisse d’épargne", "cic",
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
    # try to unglue some common patterns
    text = text.replace(":-", ": -").replace("•", "\n• ")
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
    l = fold(s)
    if re.fullmatch(r"\d{1,3}", s.strip()):
        return True
    if looks_like_price_table_line(s):
        return True
    if any(k in l for k in ADMIN_NOISE_HINTS):
        return True
    if any(k in l for k in LEGAL_KILL):
        return True
    # remove boilerplate location/time in marketing PDFs
    if "institut imagine" in l and ("etage" in l or "étage" in l or "sur la base" in l):
        return True
    if ("heure supplémentaire" in l) or ("heures supplémentaires" in l):
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


def guess_vendor_name(text: str, filename: str) -> str:
    lines = [norm(x) for x in text.splitlines() if norm(x)]
    bad = re.compile(r"\b(devis|facture|date|total|tva|siret|iban|bic|net a payer|net à payer|contact)\b", re.I)

    def penalty(ln: str) -> int:
        l = fold(ln)
        p = 0
        if "contact" in l:
            p += 7
        if "@" in ln or re.search(r"\b0[1-9](\s?\d{2}){4}\b", ln):
            p += 7
        if any(k in l for k in ["boulevard", "rue", "avenue", "paris", "bezons", "aulnay"]):
            p += 4
        # avoid section titles
        if any(k in l for k in ["accueil", "pause", "déjeuner", "cocktail"]):
            p += 8
        return p

    candidates = []
    for ln in lines[:180]:
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
        prefer = ["réceptions", "receptions", "traiteur", "production", "sas", "sarl", "inedit", "exupery", "cadet", "unik"]
        best, best_score = None, -10**9
        for c in candidates:
            fc = fold(c)
            sc = 0
            for k in prefer:
                if k in fc:
                    sc += 6
            sc -= penalty(c)
            if "institut imagine" in fc:
                sc -= 10
            if sc > best_score:
                best_score = sc
                best = c
        if best:
            return best

    return filename.rsplit(".", 1)[0]


def extract_sections(lines: List[str]) -> List[Tuple[str, List[str]]]:
    sections: List[Tuple[str, List[str]]] = []
    current_title = "Général"
    current: List[str] = []

    for ln in lines:
        if re.fullmatch(r"(RÉF|REF|QTÉ|QTE|PU|PU HT|MONTANT|TVA|TAUX|BASE HT)\b.*", ln, flags=re.I):
            current.append(ln)
            continue

        if is_section_header(ln) and (len(ln) <= 95) and (not ln.startswith(("•", "-", "–"))):
            if current:
                sections.append((current_title, current))
            current_title = ln
            current = []
            continue

        current.append(ln)

    if current:
        sections.append((current_title, current))
    return sections


def bullet_candidates(lines: List[str]) -> List[str]:
    """
    Return candidate lines that might be part of menu (not final yet).
    Keeps bullets + short non-bullets with menu hints.
    """
    items: List[str] = []
    for ln in lines:
        s = norm(ln)
        if not s or is_noise_line(s):
            continue

        if s.startswith(("•", "-", "–")):
            s2 = s.lstrip("•-– ").strip()
            if s2 and not is_noise_line(s2):
                items.append(s2)
            continue

        l = fold(s)
        if len(s) <= 180 and any(k in l for k in MENU_HINTS) and not is_noise_line(s):
            items.append(s)

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
    if any(k in l for k in ["équipe", "ingenieur", "ingénieur", "cadreur", "réalisateur", "realisateur"]):
        return "Équipe"
    if "captation" in l or "caméra" in l or "camera" in l:
        return "Captation"
    if "régie" in l or "regie" in l:
        return "Régie"
    if "diffusion" in l or "live" in l or "zoom" in l:
        return "Diffusion"
    if "replay" in l or "wetransfer" in l:
        return "Replay"
    if any(k in l for k in ["option", "forfait", "connexion", "contraint"]):
        return "Contraintes / options"
    if any(k in l for k in ["prestation", "incluant", "installation", "direct"]):
        return "Périmètre"
    return "Périmètre"


# =========================
# MODELS
# =========================
@dataclass
class Offer:
    vendor: str
    total_ttc: float | None
    candidates: Dict[str, List[str]]   # candidates per post
    selected: Dict[str, List[str]]     # selected per post
    comment: str


# =========================
# PARSERS (candidates only)
# =========================
def parse_catering(text: str, filename: str) -> Offer:
    vendor = guess_vendor_name(text, filename)
    total_ttc = find_total_ttc(text)

    lines = split_lines(text)
    filtered = [ln for ln in lines if not is_noise_line(ln)]
    sections = extract_sections(filtered)

    candidates: Dict[str, List[str]] = {k: [] for k in CATERING_POSTS}

    for title, body in sections:
        post = classify_catering_section(title)
        items = bullet_candidates(body)
        for it in items:
            if "en option" in fold(it) or fold(it).startswith("option"):
                candidates["Options"].append(it)
            else:
                candidates[post].append(it)

    # de-dup
    for k in candidates:
        out, seen = [], set()
        for it in candidates[k]:
            if is_noise_line(it):
                continue
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

    candidates: Dict[str, List[str]] = {k: [] for k in TECH_POSTS}

    for title, body in sections:
        post = classify_tech_section(title)
        items = bullet_candidates(body)

        # keep some key non-bullet lines if contain tech words
        for ln in body:
            s = norm(ln)
            if not s or is_noise_line(s):
                continue
            l = fold(s)
            if any(k in l for k in ["captation", "diffusion", "zoom", "replay", "wetransfer", "caméra", "camera", "régie", "regie", "ingénieur", "cadreur", "réalisateur", "installation", "direct"]):
                if len(s) <= 200:
                    items.append(s)

        # de-dup
        out, seen = [], set()
        for it in items:
            it = norm(it)
            if not it or is_noise_line(it):
                continue
            kk = fold(it)
            if kk in seen:
                continue
            seen.add(kk)
            out.append(it)

        candidates[post].extend(out)

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
# WORD (premium)
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


def set_run_brand(run, bold=False, size=10, color=None):
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = hex_to_rgb(color)


def add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_brand(run, bold=True, size=16, color=PRIMARY)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_subtitle(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_brand(run, bold=True, size=11, color="#111827")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_small(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_brand(run, bold=False, size=9, color="#374151")


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

    # PAGE 1 — Catering full table
    if catering:
        add_subtitle(doc, "1) PRESTATION TRAITEUR — Comparatif (synthèse)")
        vendors = catering[:3]
        table = doc.add_table(rows=1, cols=1 + len(vendors))

        hdr = table.rows[0].cells
        hdr[0].text = "Poste"
        set_cell_shading(hdr[0], PRIMARY)
        for run in hdr[0].paragraphs[0].runs:
            set_run_brand(run, bold=True, size=10, color="#FFFFFF")

        for i, off in enumerate(vendors, start=1):
            hdr[i].text = off.vendor
            set_cell_shading(hdr[i], PRIMARY)
            for run in hdr[i].paragraphs[0].runs:
                set_run_brand(run, bold=True, size=10, color="#FFFFFF")

        for cell in hdr:
            set_cell_margins(cell)

        rows = [
            ("Total TTC (hors options)", lambda o: euro_fmt(o.total_ttc)),
            ("Accueil café", lambda o: summarize_for_table(o.selected.get("Accueil café", []), 320)),
            ("Pause matin", lambda o: summarize_for_table(o.selected.get("Pause matin", []), 320)),
            ("Déjeuner", lambda o: summarize_for_table(o.selected.get("Déjeuner", []), 380)),
            ("Pause après-midi", lambda o: summarize_for_table(o.selected.get("Pause après-midi", []), 320)),
            ("Cocktail", lambda o: summarize_for_table(o.selected.get("Cocktail", []), 380)),
            ("Boissons (global)", lambda o: summarize_for_table(o.selected.get("Boissons (global)", []), 260)),
            ("Options", lambda o: summarize_for_table(o.selected.get("Options", []), 260)),
            ("Commentaire", lambda o: (norm(o.comment)[:260] + (" (...)" if len(norm(o.comment)) > 260 else ""))),
        ]

        for label, fn in rows:
            r = table.add_row().cells
            r[0].text = label
            set_cell_shading(r[0], "F3F4F6")
            for run in r[0].paragraphs[0].runs:
                set_run_brand(run, bold=True, size=9, color="#111827")
            set_cell_margins(r[0])

            for j, off in enumerate(vendors, start=1):
                r[j].text = fn(off) or "—"
                for para in r[j].paragraphs:
                    for run in para.runs:
                        set_run_brand(run, bold=False, size=9, color="#111827")
                set_cell_margins(r[j])

        doc.add_paragraph("")

    # PAGE 1 — Tech summary (1 table per vendor)
    if tech:
        add_subtitle(doc, "2) PRESTATION TECHNIQUE — Synthèse")
        for idx, off in enumerate(tech[:2], start=1):
            p = doc.add_paragraph()
            run = p.add_run(f"Prestataire technique {idx} : {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            set_run_brand(run, bold=True, size=9, color="#111827")

            t = doc.add_table(rows=1, cols=2)
            h = t.rows[0].cells
            h[0].text = "Item"
            h[1].text = "Détail"
            set_cell_shading(h[0], PRIMARY)
            set_cell_shading(h[1], PRIMARY)
            for cell in h:
                for run in cell.paragraphs[0].runs:
                    set_run_brand(run, bold=True, size=10, color="#FFFFFF")
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
                for run in rr[0].paragraphs[0].runs:
                    set_run_brand(run, bold=True, size=9, color="#111827")
                for cell in rr:
                    set_cell_margins(cell)
                    for para in cell.paragraphs:
                        for run in para.runs:
                            set_run_brand(run, bold=False, size=9, color="#111827")

            doc.add_paragraph("")

    # DETAILS
    doc.add_page_break()
    add_title(doc, "DÉTAIL DES OFFRES (éléments validés)")
    add_small(doc, "Les listes ci-dessous correspondent aux éléments cochés dans l’outil (contrôle qualité).")

    if catering:
        doc.add_paragraph("")
        add_subtitle(doc, "A) TRAITEUR — Détail par prestataire")
        for i, off in enumerate(catering[:3], start=1):
            doc.add_paragraph("")
            add_subtitle(doc, f"PROPOSITION {i} — {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            for post in ["Accueil café", "Pause matin", "Déjeuner", "Pause après-midi", "Cocktail", "Boissons (global)", "Options", "Autres (logistique)"]:
                items = off.selected.get(post, [])
                p = doc.add_paragraph()
                run = p.add_run(f"{post} :")
                set_run_brand(run, bold=True, size=10, color="#111827")
                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it)
                        para.style = doc.styles["List Bullet"]
                        for r in para.runs:
                            set_run_brand(r, bold=False, size=9, color="#111827")
            if off.comment.strip():
                p = doc.add_paragraph()
                run = p.add_run("Commentaire : ")
                set_run_brand(run, bold=True, size=10, color="#111827")
                doc.add_paragraph(off.comment)

    if tech:
        doc.add_paragraph("")
        add_subtitle(doc, "B) TECHNIQUE — Détail par prestataire")
        for i, off in enumerate(tech[:2], start=1):
            doc.add_paragraph("")
            add_subtitle(doc, f"PRESTATION TECHNIQUE {i} — {off.vendor} — Total TTC : {euro_fmt(off.total_ttc)}")
            for post in TECH_POSTS:
                items = off.selected.get(post, [])
                p = doc.add_paragraph()
                run = p.add_run(f"{post} :")
                set_run_brand(run, bold=True, size=10, color="#111827")
                if not items:
                    doc.add_paragraph("—")
                else:
                    for it in items:
                        para = doc.add_paragraph(it)
                        para.style = doc.styles["List Bullet"]
                        for r in para.runs:
                            set_run_brand(r, bold=False, size=9, color="#111827")
            if off.comment.strip():
                p = doc.add_paragraph()
                run = p.add_run("Conseil (synthèse) : ")
                set_run_brand(run, bold=True, size=10, color="#111827")
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
st.caption("Upload → extraction → sélection (checkbox) → Word premium charté Imagine.")
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

# Parse into offers with candidates
with st.spinner("Lecture des PDFs…"):
    catering_offers: List[Offer] = []
    tech_offers: List[Offer] = []

    for f in catering_files:
        txt = extract_pdf_text(f)
        catering_offers.append(parse_catering(txt, f.name))

    for f in tech_files:
        txt = extract_pdf_text(f)
        tech_offers.append(parse_tech(txt, f.name))

# Validation UI
tab1, tab2 = st.tabs(["Traiteur (sélection)", "Technique (sélection)"])

def vendor_quality_warning(v: str) -> bool:
    lv = fold(v)
    return any(k in lv for k in ["accueil", "pause", "déjeuner", "cocktail"]) or ("@" in v) or ("contact" in lv)

with tab1:
    if not catering_offers:
        st.caption("Aucun devis traiteur.")
    else:
        st.caption("Coche uniquement ce que tu veux voir dans la synthèse direction + détails.")
        for i, off in enumerate(catering_offers, start=1):
            with st.expander(f"PROPOSITION {i} — {off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire (obligatoire)", value=off.vendor, key=f"c_vendor_{i}")
                if vendor_quality_warning(off.vendor):
                    st.error("Nom prestataire suspect (ressemble à un poste / contact). Corrige-le.")

                ttc_str = st.text_input("Total TTC (hors options)", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"c_ttc_{i}")
                off.total_ttc = parse_eur_amount(ttc_str)
                if off.total_ttc is None:
                    st.warning("Total TTC non détecté — saisis-le pour une synthèse propre.")

                # selection per post
                for post in CATERING_POSTS:
                    cand = off.candidates.get(post, [])
                    if not cand:
                        continue
                    st.markdown(f"**{post}**")
                    sel = []
                    for idx, line in enumerate(cand):
                        key = f"c_{i}_{post}_{idx}"
                        checked = st.checkbox(line, value=(idx < 6 if post in ["Déjeuner", "Cocktail"] else idx < 4), key=key)
                        if checked:
                            sel.append(line)
                    off.selected[post] = sel

                off.comment = st.text_area("Commentaire (1–2 phrases max)", value=off.comment, height=80, key=f"c_comment_{i}")

with tab2:
    if not tech_offers:
        st.caption("Aucun devis technique.")
    else:
        for i, off in enumerate(tech_offers, start=1):
            with st.expander(f"TECHNIQUE {i} — {off.vendor}", expanded=(i == 1)):
                off.vendor = st.text_input("Nom prestataire (obligatoire)", value=off.vendor, key=f"t_vendor_{i}")
                if vendor_quality_warning(off.vendor):
                    st.warning("Nom prestataire à vérifier.")
                ttc_str = st.text_input("Total TTC", value=("" if off.total_ttc is None else euro_fmt(off.total_ttc)), key=f"t_ttc_{i}")
                off.total_ttc = parse_eur_amount(ttc_str)
                if off.total_ttc is None:
                    st.warning("Total TTC non détecté — saisis-le.")

                for post in TECH_POSTS:
                    cand = off.candidates.get(post, [])
                    if not cand:
                        continue
                    st.markdown(f"**{post}**")
                    sel = []
                    for idx, line in enumerate(cand):
                        key = f"t_{i}_{post}_{idx}"
                        checked = st.checkbox(line, value=(idx < 6), key=key)
                        if checked:
                            sel.append(line)
                    off.selected[post] = sel

                off.comment = st.text_area("Conseil (2–3 phrases max)", value=off.comment, height=90, key=f"t_comment_{i}")

st.divider()
colA, colB = st.columns([2, 1], vertical_alignment="center")
with colA:
    st.caption("Cette V3 privilégie la fiabilité : le Word contient uniquement les lignes validées (cochées).")
with colB:
    if st.button("Générer le Word premium (.docx)", use_container_width=True, type="primary"):
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
