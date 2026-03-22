"""
Backend de génération de rapport ITV - Immeau
Reçoit les données JSON de l'app Flutter, génère le .docx et l'envoie par mail.
"""

from flask import Flask, request, jsonify
import re
import os
import io
import base64
import zipfile
import requests as req_lib
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
from lxml import etree
import copy

app = Flask(__name__)
# Augmente la limite de taille des requêtes pour accepter les photos base64
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024  # 80 MB

# ─────────────────────────────────────────────
# Configuration email via Brevo HTTP API
# (Render.com bloque le SMTP ; l'API HTTP fonctionne toujours)
# Env var à renseigner dans Render.com :
#   BREVO_API_KEY = xkeysib-...  (Brevo → SMTP & API → Clés API)
#   MAIL_FROM     = ivt@immeau.fr
#   MAIL_TO       = ivt@immeau.fr
# ─────────────────────────────────────────────
BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")
MAIL_FROM = os.environ.get("MAIL_FROM", "ivt@immeau.fr")
MAIL_TO   = os.environ.get("MAIL_TO",   "ivt@immeau.fr")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.docx")


# ─────────────────────────────────────────────
# Helpers python-docx
# ─────────────────────────────────────────────

def replace_text_in_paragraph(paragraph, old: str, new: str):
    """Remplace old par new dans un paragraphe en préservant le formatage du premier run."""
    full = "".join(run.text for run in paragraph.runs)
    if old not in full:
        return False
    new_full = full.replace(old, new)
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
    return any(e.tag.endswith("}numPr") for e in elem.iter())


def _has_underline(elem) -> bool:
    return any(e.tag.endswith("}u") for e in elem.iter())


def delete_section_by_title(doc: Document, section_title: str):
    body = doc.element.body
    paras = list(body)
    target_idx = None
    for i, elem in enumerate(paras):
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            text = "".join(t.text or "" for t in elem.iter() if t.tag.endswith("}t"))
            if section_title in text:
                target_idx = i
                break
    if target_idx is None:
        return
    to_remove = [paras[target_idx]]
    for i in range(target_idx + 1, len(paras)):
        elem = paras[i]
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            break
        to_remove.append(elem)
    for elem in to_remove:
        try:
            body.remove(elem)
        except ValueError:
            pass


# ─────────────────────────────────────────────
# Gestion des photos
# ─────────────────────────────────────────────

def _replace_facade_photo(docx_bytes: bytes, new_photo_bytes: bytes) -> bytes:
    """
    Remplace la photo de façade (rId8 -> media/image2.jpeg) dans le DOCX.
    Manipule directement le zip pour echanger le fichier image sans toucher au XML.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as z:
            rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        match = re.search(r'Id="rId8"[^>]+Target="([^"]+)"', rels_xml)
        if not match:
            return docx_bytes
        target = f"word/{match.group(1)}"
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zin:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == target:
                        zout.writestr(item, new_photo_bytes)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        return output.getvalue()
    except Exception as e:
        print(f"[WARN] Impossible de remplacer la photo de facade : {e}")
        return docx_bytes


def _insert_photos_complement(doc: Document, photos_commentees: list):
    """
    Insere les photos avec commentaires dans la section VI - COMPLEMENT PHOTOGRAPHIQUE.
    Chaque entree de photos_commentees : { 'image_base64': str, 'commentaire': str }
    """
    if not photos_commentees:
        return
    body = doc.element.body
    vi_elem = None
    vii_elem = None
    for elem in body:
        if not elem.tag.endswith('}p'):
            continue
        text = ''.join(t.text or '' for t in elem.iter() if t.tag.endswith('}t'))
        if 'VI' in text and 'COMPL' in text and vi_elem is None:
            vi_elem = elem
        elif 'VII' in text and vi_elem is not None and vii_elem is None:
            vii_elem = elem
            break
    if vi_elem is None or vii_elem is None:
        print("[WARN] Section VI-VII introuvable, photos non inserees.")
        return
    to_remove = []
    elem = vi_elem.getnext()
    while elem is not None and elem is not vii_elem:
        text = ''.join(t.text or '' for t in elem.iter() if t.tag.endswith('}t'))
        if not text.strip():
            to_remove.append(elem)
        elem = elem.getnext()
    for e in to_remove:
        body.remove(e)
    for photo_data in photos_commentees:
        image_b64 = photo_data.get('image_base64', '')
        commentaire = photo_data.get('commentaire', '')
        try:
            image_bytes = base64.b64decode(image_b64)
        except Exception:
            continue
        img_para = doc.add_paragraph()
        img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = img_para.add_run()
        run.add_picture(io.BytesIO(image_bytes), width=Inches(5))
        img_elem = img_para._element
        body.remove(img_elem)
        vii_elem.addprevious(img_elem)
        if commentaire:
            comment_para = doc.add_paragraph()
            comment_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cr = comment_para.add_run(commentaire)
            cr.italic = True
            comment_elem = comment_para._element
            body.remove(comment_elem)
            vii_elem.addprevious(comment_elem)
        sep_para = doc.add_paragraph()
        sep_elem = sep_para._element
        body.remove(sep_elem)
        vii_elem.addprevious(sep_elem)


# ─────────────────────────────────────────────
# Logique principale de remplissage du template
# ─────────────────────────────────────────────

def build_rapport(data: dict) -> bytes:
    """
    Prend les donnees de l'app Flutter (dict JSON) et retourne
    le contenu binaire du .docx genere.
    """
    doc = Document(TEMPLATE_PATH)

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

    desc_site      = data.get("descriptionSite", "")
    parcelle       = data.get("parcelleCadastre", "")
    section_cad    = data.get("sectionCadastre", "")
    objet_mission  = data.get("objetMission", "")

    mat_apparents  = data.get("materiauxApparents", "")
    etat_apparents = data.get("etatApparents", "")
    mat_enterres   = data.get("materiauxEnterres", "")
    etat_enterres  = data.get("etatEnterres", "")

    paragraphes    = data.get("paragraphesSelectionnes", [])

    replacements = {
        "00 RUE DU XXX, PARIS 00":    adresse_full.upper(),
        "SDC DU XX RUE DU XXX":       client.upper(),
        "Foncia Paris Est Nation":     mo_delegue,
        "74 Boulevard de Reuilly":     adresse_mo,
        "75012 Paris":                 f"{cp_mo} {ville_mo}",
        "00000":                       devis,
        "Octobre 2024":                date_rapport,
        "Quentin Zezuka":              redacteur,
        "J\u00e9r\u00e9my Hahn":                 verificateur,
        "Fonte.":                      f"{mat_enterres}.",
        "Fonte, PVC .":                f"{mat_apparents}.",
    }
    replace_in_doc(doc, replacements)

    for para in doc.paragraphs:
        t = para.text
        if "SDC du 200 Boulevard de Charonne." in t and "pr\u00e9sente \u00e9tude" in t:
            replace_text_in_paragraph(para, "SDC du 200 Boulevard de Charonne.", client + ".")
        elif "inspection t\u00e9l\u00e9vis\u00e9e des r\u00e9seaux" in t and "objectif" in t and objet_mission:
            replace_text_in_paragraph(para, t, objet_mission)
        elif "200 Boulevard de Charonne" in t and "parcelle" in t:
            replace_text_in_paragraph(para, "200 Boulevard de Charonne", adresse)
            replace_text_in_paragraph(para, "(parcelle cadastrale , section ).",
                f"(parcelle cadastrale {parcelle}, section {section_cad}).")
        elif "b\u00e2timent sur rue, avec 6 \u00e9tage(s)" in t and desc_site:
            replace_text_in_paragraph(para, t, desc_site)

    SECTION_MAP = {
        "colonne_ep":            "Colonne d\u2019eaux pluviales de fa\u00e7ade",
        "regard_limite":         "Regard de limite de propri\u00e9t\u00e9",
        "sanitaires_sous_sol":   "Installations sanitaires en sous-sol",
        "ancienne_fosse":        "Ancienne fosse d\u2019aisance",
        "regards_non_etanches":  "Regards de visite non \u00e9tanches",
        "reseau_separatif":      "R\u00e9seau s\u00e9paratif",
        "ventilation":           "Ventilations des r\u00e9seaux",
        "cas_restaurants":       "Cas des eaux us\u00e9es provenant des restaurants",
        "cas_garages":           "Cas des eaux provenant des garages",
    }

    for key, title in SECTION_MAP.items():
        if key not in paragraphes:
            delete_section_by_title(doc, title)

    replace_in_doc(doc, {
        "r\u00e8glement d'assainissement de la ville de Paris": f"r\u00e8glement d'assainissement {reglement}",
        "Le r\u00e8glement d'assainissement x": f"Le r\u00e8glement d'assainissement {reglement}",
    })

    photos_commentees = data.get("photosCommentees", [])
    if photos_commentees:
        _insert_photos_complement(doc, photos_commentees)

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    photo_facade_b64 = data.get("photoFacade")
    if photo_facade_b64:
        try:
            facade_bytes = base64.b64decode(photo_facade_b64)
            docx_bytes = _replace_facade_photo(docx_bytes, facade_bytes)
        except Exception as e:
            print(f"[WARN] Erreur decodage photo facade : {e}")

    return docx_bytes


# ─────────────────────────────────────────────
# Envoi du mail
# ─────────────────────────────────────────────

def send_email(docx_bytes: bytes, filename: str, devis: str, adresse: str):
    """Envoie le .docx via l'API HTTP Brevo (pas de SMTP, pas de blocage reseau)."""
    if not BREVO_API_KEY:
        raise Exception("BREVO_API_KEY manquante dans les variables d'environnement")
    body_text = (
        f"Bonjour,\n\n"
        f"Le rapport [{filename}] a ete correctement genere et peut etre telecharge en piece jointe.\n\n"
        f"Devis : {devis}\n"
        f"Adresse : {adresse}\n\n"
        f"Application Immeau."
    )
    payload = {
        "sender":      {"name": "Immeau", "email": MAIL_FROM},
        "to":          [{"email": MAIL_TO}],
        "subject":     f"Creation du rapport [{filename.replace('.docx', '')}]",
        "textContent": body_text,
        "attachment":  [{"name": filename, "content": base64.b64encode(docx_bytes).decode()}],
    }
    resp = req_lib.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"api-key": BREVO_API_KEY, "Content-Type": "application/json"},
        json=payload,
        timeout=30,
    )
    if resp.status_code not in (200, 201):
        raise Exception(f"Brevo API erreur {resp.status_code}: {resp.text}")


# ─────────────────────────────────────────────
# Routes Flask
# ─────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generer_rapport", methods=["POST"])
def generer_rapport():
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "Corps JSON manquant"}), 400
        docx_bytes = build_rapport(data)
        devis   = data.get("devis", "00000")
        adresse = data.get("adresseProjet", "")
        filename = f"Rapport d'investigations {devis}.docx"
        send_email(docx_bytes, filename, devis, adresse)
        return jsonify({"success": True, "message": f"Rapport {filename} envoye a {MAIL_TO}"})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/telecharger_rapport", methods=["POST"])
def telecharger_rapport():
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
