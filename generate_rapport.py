"""
Backend de génération de rapport ITV - Immeau
Reçoit les données JSON de l'app Flutter, génère le .docx et l'envoie par mail.
"""

from flask import Flask, request, jsonify
import re
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from docx import Document
from docx.shared import Pt
from copy import deepcopy
from lxml import etree
import copy

app = Flask(__name__)

# ─────────────────────────────────────────────
# Configuration email (à adapter selon le serveur SMTP)
# ─────────────────────────────────────────────
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "rapport@app.immeau.fr")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
MAIL_FROM = os.environ.get("MAIL_FROM", "rapport@app.immeau.fr")
MAIL_TO   = os.environ.get("MAIL_TO",   "ivt@immeau.fr")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.docx")


# ─────────────────────────────────────────────
# Helpers python-docx
# ─────────────────────────────────────────────

def replace_text_in_paragraph(paragraph, old: str, new: str):
    """Remplace old par new dans un paragraphe en préservant le formatage du premier run."""
    # Reconstitue le texte complet du paragraphe
    full = "".join(run.text for run in paragraph.runs)
    if old not in full:
        return False
    # Remplace dans le texte reconstitué
    new_full = full.replace(old, new)
    # Vide tous les runs sauf le premier, qui reçoit tout le texte
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = new_full
        else:
            run.text = ""
    return True


def replace_in_doc(doc: Document, replacements: dict):
    """Applique tous les remplacements textuels dans le document entier."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for old, new in replacements.items():
                        replace_text_in_paragraph(para, old, str(new))
    for para in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_in_paragraph(para, old, str(new))


def find_paragraph_index(doc: Document, text_contains: str):
    """Retourne l'index du paragraphe contenant text_contains, ou -1."""
    for i, para in enumerate(doc.paragraphs):
        if text_contains in para.text:
            return i
    return -1


def delete_paragraphs_between(doc: Document, start_marker: str, end_marker: str):
    """Supprime tous les paragraphes entre start_marker et end_marker (inclus)."""
    body = doc.element.body
    paras = list(body)

    in_section = False
    to_remove = []

    for elem in paras:
        if elem.tag.endswith("}p"):
            text = "".join(t.text or "" for t in elem.iter() if t.tag.endswith("}t"))
            if start_marker in text:
                in_section = True
            if in_section:
                to_remove.append(elem)
            if in_section and end_marker in text and text != start_marker:
                break

    for elem in to_remove:
        body.remove(elem)


def _has_numpr(elem) -> bool:
    """Vérifie si un élément <w:p> a une propriété numPr (élément de liste)."""
    return any(e.tag.endswith("}numPr") for e in elem.iter())


def _has_underline(elem) -> bool:
    """Vérifie si un élément <w:p> contient du texte souligné."""
    return any(e.tag.endswith("}u") for e in elem.iter())


def delete_section_by_title(doc: Document, section_title: str):
    """
    Supprime une section de conclusion identifiée par son titre.
    Structure : chaque section = paragraphe numPr+underline (titre),
    suivi de paragraphes de contenu, jusqu'au prochain numPr+underline.
    """
    body = doc.element.body
    paras = list(body)

    # Trouve le paragraphe titre (numPr=True, underline=True, contient section_title)
    target_idx = None
    for i, elem in enumerate(paras):
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            text = "".join(t.text or "" for t in elem.iter() if t.tag.endswith("}t"))
            if section_title in text:
                target_idx = i
                break

    if target_idx is None:
        return  # Section non trouvée, rien à faire

    # Collecte tous les éléments à supprimer : du titre jusqu'au
    # prochain titre de section (numPr + underline) ou fin du document
    to_remove = [paras[target_idx]]

    for i in range(target_idx + 1, len(paras)):
        elem = paras[i]
        # S'arrête au prochain titre de section de conclusions
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            break
        # Supprime également les éléments non-paragraphes (tables, etc.) entre les sections
        to_remove.append(elem)

    for elem in to_remove:
        try:
            body.remove(elem)
        except ValueError:
            pass


# ─────────────────────────────────────────────
# Logique principale de remplissage du template
# ─────────────────────────────────────────────

def build_rapport(data: dict) -> bytes:
    """
    Prend les données de l'app Flutter (dict JSON) et retourne
    le contenu binaire du .docx généré.
    """
    doc = Document(TEMPLATE_PATH)

    # ── 1. Données générales (page de garde + tableau) ──────────────────
    adresse        = data.get("adresseProjet", "")
    ville          = data.get("villeProjet", "")
    cp             = data.get("cpProjet", "")
    client         = data.get("client", "")
    mo_delegue     = data.get("moDelegue", "")
    adresse_mo     = data.get("adresseMoDelegue", "")
    ville_mo       = data.get("villeMoDelegue", "")
    cp_mo          = data.get("cpMoDelegue", "")
    devis          = data.get("devis", "")
    date_rapport   = data.get("dateRapport", "")
    redacteur      = data.get("redacteur", "")
    verificateur   = data.get("verificateur", "")
    titre_etude    = data.get("titreEtude", "")
    reglement      = data.get("reglementApplicable", "Ville de Paris")

    adresse_full = f"{adresse}, {cp} {ville}".strip(", ")
    adresse_mo_full = f"{adresse_mo}, {cp_mo} {ville_mo}".strip(", ")

    # ── 2. Description du site (généré par Claude) ──────────────────────
    desc_site      = data.get("descriptionSite", "")
    parcelle       = data.get("parcelleCadastre", "")
    section_cad    = data.get("sectionCadastre", "")
    objet_mission  = data.get("objetMission", "")

    # ── 3. Canalisations ────────────────────────────────────────────────
    mat_apparents  = data.get("materiauxApparents", "")   # ex: "Fonte, PVC"
    etat_apparents = data.get("etatApparents", "")        # ex: "mauvais"
    mat_enterres   = data.get("materiauxEnterres", "")
    etat_enterres  = data.get("etatEnterres", "")

    # ── 4. Paragraphes sélectionnés (liste de clés) ──────────────────────
    paragraphes    = data.get("paragraphesSelectionnes", [])

    # ── 5. Remplacements simples (page de garde + champs fixes) ──────────
    replacements = {
        # Page de garde
        "00 RUE DU XXX, PARIS 00":    adresse_full.upper(),
        "SDC DU XX RUE DU XXX":       client.upper(),
        "Foncia Paris Est Nation":     mo_delegue,
        "74 Boulevard de Reuilly":     adresse_mo,
        "75012 Paris":                 f"{cp_mo} {ville_mo}",
        "00000":                       devis,
        "Octobre 2024":                date_rapport,
        "Quentin Zezuka":              redacteur,
        "Jérémy Hahn":                 verificateur,
        # Canalisations
        "Fonte.":                      f"{mat_enterres}.",
        "Fonte, PVC .":                f"{mat_apparents}.",
    }
    replace_in_doc(doc, replacements)

    # ── 6. Remplace les paragraphes longs (description site, objet mission) ──
    # On remplace paragraphe entier pour éviter les doublons de texte

    for para in doc.paragraphs:
        t = para.text

        # Objet de la mission : remplace le client dans la première phrase
        if "SDC du 200 Boulevard de Charonne." in t and "présente étude" in t:
            replace_text_in_paragraph(
                para, "SDC du 200 Boulevard de Charonne.",
                client + "."
            )

        # Objet de la mission : remplace la phrase d'objectif
        elif "inspection télévisée des réseaux" in t and "objectif" in t and objet_mission:
            replace_text_in_paragraph(para, t, objet_mission)

        # Description du site : adresse principale
        elif "200 Boulevard de Charonne" in t and "parcelle" in t:
            replace_text_in_paragraph(
                para, "200 Boulevard de Charonne", adresse
            )
            replace_text_in_paragraph(
                para, "(parcelle cadastrale , section ).",
                f"(parcelle cadastrale {parcelle}, section {section_cad})."
            )

        # Description du site : description bâtiments
        elif "bâtiment sur rue, avec 6 étage(s)" in t and desc_site:
            replace_text_in_paragraph(para, t, desc_site)

    # ── 7. Supprime les sections de conclusions non sélectionnées ────────
    # Mapping : clé Flutter → titre dans le document
    SECTION_MAP = {
        "colonne_ep":            "Colonne d\u2019eaux pluviales de fa\u00e7ade",
        "regard_limite":         "Regard de limite de propriété",
        "sanitaires_sous_sol":   "Installations sanitaires en sous-sol",
        "ancienne_fosse":        "Ancienne fosse d\u2019aisance",
        "regards_non_etanches":  "Regards de visite non étanches",
        "reseau_separatif":      "Réseau séparatif",
        "ventilation":           "Ventilations des réseaux",
        "cas_restaurants":       "Cas des eaux usées provenant des restaurants",
        "cas_garages":           "Cas des eaux provenant des garages",
    }

    for key, title in SECTION_MAP.items():
        if key not in paragraphes:
            delete_section_by_title(doc, title)

    # ── 8. Règlement d'assainissement ────────────────────────────────────
    replace_in_doc(doc, {
        "règlement d'assainissement de la ville de Paris": f"règlement d'assainissement {reglement}",
        "Le règlement d'assainissement x": f"Le règlement d'assainissement {reglement}",
    })

    # ── 9. Sérialise en mémoire ──────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# Envoi du mail
# ─────────────────────────────────────────────

def send_email(docx_bytes: bytes, filename: str, devis: str, adresse: str):
    """Envoie le .docx par mail à MAIL_TO."""
    msg = MIMEMultipart()
    msg["From"]    = MAIL_FROM
    msg["To"]      = MAIL_TO
    msg["Subject"] = f"Création du rapport [{filename.replace('.docx', '')}]"

    body = f"""Bonjour,

Le rapport [{filename}] a été correctement généré et peut être téléchargé en pièce jointe.

Devis : {devis}
Adresse : {adresse}

Application Immeau.
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    part = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
    part.set_payload(docx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(MAIL_FROM, MAIL_TO, msg.as_string())


# ─────────────────────────────────────────────
# Routes Flask
# ─────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generer_rapport", methods=["POST"])
def generer_rapport():
    """
    Endpoint principal.
    Corps : JSON avec toutes les données du projet.
    Retourne : { "success": true } ou { "error": "..." }
    """
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "Corps JSON manquant"}), 400

        # Génère le document
        docx_bytes = build_rapport(data)

        # Nom du fichier
        devis   = data.get("devis", "00000")
        adresse = data.get("adresseProjet", "")
        filename = f"Rapport d'investigations {devis}.docx"

        # Envoie le mail
        send_email(docx_bytes, filename, devis, adresse)

        return jsonify({"success": True, "message": f"Rapport {filename} envoyé à {MAIL_TO}"})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/telecharger_rapport", methods=["POST"])
def telecharger_rapport():
    """
    Variante : retourne le .docx directement (pour test ou téléchargement direct).
    """
    from flask import send_file
    try:
        data = request.get_json(force=True)
        docx_bytes = build_rapport(data)
        devis    = data.get("devis", "00000")
        filename = f"Rapport d'investigations {devis}.docx"
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
