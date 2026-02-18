# app.py — Synthèse devis prestataires (FINAL++++)
# ---------------------------------------------
# Fixes vs version précédente:
# - Vendor: supprime suffixes " 1" / " - 1" / " page 1", coupe "SARL ... 1"
# - Traiteur: routage "viennoiseries/mini pains/croissant" -> pauses (pas Cocktail)
# - Cocktail: compte pièces uniquement sur items "cocktail" (évite mini viennoiseries)
# - Tech: supprime "TV" parasite (résidu TVA) en noise+filters
# - Intro détail: pas de page blanche (déjà OK)

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

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
MENU_HINTS = [
    "accueil", "petit", "déjeuner", "dejeuner", "pause", "buffet", "cocktail", "apéritif", "aperitif",
    "déjeunatoire", "dejeunatoire",
    "café", "cafe", "thé", "the", "soft", "jus", "eau",
    "viennoiser", "gourmand", "mignard", "financier", "cannel", "cookie", "brownie", "madeleine",
    "pièce", "pieces", "pièces", "/pers", "par personne", "convive", "invité", "invite",
    "salée", "sucrée", "sucree",
    "sandwich", "wrap", "salade", "fromage", "fruit",
    "vin", "champagne", "bière", "biere",
    "thermos", "gobelet", "tasse", "serviette", "plateau",
    "verrerie", "flûtes", "assiettes", "nappage", "mobilier", "mange-debout",
    "livraison", "reprise", "mise en place", "personnel", "service",
]

TECH_HINTS = [
    "captation", "caméra", "camera", "4k", "cadreur", "réalisateur", "realisateur",
    "ingénieur", "ingenieur", "son", "audio",
    "régie", "regie", "diffusion", "live", "zoom", "duplex", "plateforme",
    "replay", "wetransfer", "we transfer", "enregistrement",
    "pavlov", "zapette", "tv", "écran", "ecran",
    "micro", "hf", "console", "mélangeur", "melangeur", "obs", "vmix",
]

LEGAL_FORMS = ["sas", "sarl", "sa", "eurl", "sasu", "association", "scop", "groupe"]

ADDRESS_HINTS = [
    "rue", "avenue", "boulevard", "allée", "allee", "bp", "cedex",
    "france", "paris", "clichy", "nanterre", "saint",
    "quai", "impasse", "route", "chemin",
]

VENDOR_FORBIDDEN = [
    "en euros", "appliqué", "applique", "devis", "facture",
    "désignation", "designation", "quantité", "quantite",
    "montant", "tva", "base ht", "total", "ttc", "ht",
    "récapitulatif", "recapitulatif", "proposition", "prestation",
    "prix par convive",
    "au capital",
]

ADMIN_HINTS = [
    "conditions générales", "conditions generales", "cgv",
    "rgpd", "données personnelles", "donnees personnelles",
    "iban", "bic", "rib", "banque", "tva intracommunautaire",
    "siret", "rcs", "capital",
    "adresse", "tél", "tel", "email", "e-mail", "site internet", "www.",
    "référence", "reference", "date de devis", "date de validité", "signature",
    "mode de paiement", "net à payer", "net a payer",
    "tribunal", "mise en demeure", "penalite", "pénalité",
    "indemnité forfaitaire", "indemnite forfaitaire",
    "page ",
    "code description", "montant ht", "montant tva", "base ht",
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
    # décolle lettres/chiffres (Réceptions1 -> Réceptions 1)
    text = re.sub(r"([A-Za-zÀ-ÖØ-öø-ÿ])(\d)", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-zÀ-ÖØ-öø-ÿ])", r"\1 \2", text)

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
        if "conditions générales" in l or "conditions generales" in l or re.search(r"\bcgv\b", l):
            break
        out.append(ln)
    return out


def looks_like_schedule_line(s: str) -> bool:
    l = fold(s)
    return bool(re.search(r"\b0?\d{1,2}h\d{2}\b", l) and ((" à " in l) or (" a " in l) or ("-" in l)) and len(l) <= 260)


def looks_like_price_table_line(s: str) -> bool:
    core = re.sub(r"[€]", "", norm(s))
    digits = sum(ch.isdigit() for ch in core)
    letters = sum(ch.isalpha() for ch in core)
    if digits >= 10 and digits > letters:
        return True
    if re.fullmatch(r"[\d\s,\.%\-\/]+", core) and digits >= 6:
        return True
    return False


def is_addressy(s: str) -> bool:
    l = fold(s)
    has_cp = bool(re.search(r"\b\d{5}\b", l))
    strong = bool(re.search(r"^\s*\d{1,4}\s*(bis|ter)?\s*(rue|avenue|boulevard|quai|impasse|route|chemin)\b", l))
    hits = sum(k in l for k in ADDRESS_HINTS)
    return strong or (has_cp and hits >= 1) or (hits >= 2)


def is_admin_line(s: str) -> bool:
    l = fold(s)
    if any(k in l for k in ADMIN_HINTS):
        return True
    if re.search(r"\bpage\s+\d+\s+sur\s+\d+\b", l):
        return True
    compact = re.sub(r"\s+", "", s)
    if re.search(r"\bFR\d{2}[0-9A-Z]{10,30}\b", compact):
        return True
    if "@" in s:
        return True
    if re.search(r"\b0[1-9](\s?\d{2}){4}\b", s):
        return True
    return False


def is_noise_line(s: str) -> bool:
    s = norm(s)
    if not s:
        return True

    # IMPORTANT: tue "TV" (résidu TVA)
    if fold(s) in
