"""
Backend de génération de rapport ITV - Immeau
Reçoit les données JSON de l'app Flutter, génère le .docx et l'envoie par mail.
"""

from flask import Flask, request, jsonify
import re
import os
import io
import base64
import requests as req_lib
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
import copy

# Regex pour détecter les en-têtes majeurs de section (I –, II –, III –, ..., VI –, VII –, etc.)
_MAJOR_SECTION_RE = re.compile(r'^[IVX]+\s+\u2013\s+')

app = Flask(__name__)
# Augmente la limite de taille des requêtes pour accepter les photos base64
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024  # 80 MB

# ─────────────────────────────────────────────
# Configuration email via Brevo HTTP API
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

def replace_text_in_paragraph(paragraph, old: str, new: str) -> bool:
    """Remplace old par new dans un paragraphe en préservant le formatage du premier run."""
    full = "".join(run.text for run in paragraph.runs)
    if old not in full:
        return False
    new_full = full.replace(old, new)
    for i, run in enumerate(paragraph.runs):
        run.text = new_full if i == 0 else ""
    return True


def replace_in_doc(doc: Document, replacements: dict):
    """Applique tous les remplacements textuels dans le document entier (paragraphes + tables)."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for old, new in replacements.items():
                        replace_text_in_paragraph(para, old, str(new))
    for para in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_in_paragraph(para, old, str(new))


def _get_para_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter() if t.tag.endswith("}t"))


def _has_numpr(elem) -> bool:
    return any(e.tag.endswith("}numPr") for e in elem.iter())


def _has_underline(elem) -> bool:
    return any(e.tag.endswith("}u") for e in elem.iter())


def _remove_elements(body, elements):
    """Supprime une liste d'éléments du body du document."""
    for elem in elements:
        try:
            body.remove(elem)
        except ValueError:
            pass


def delete_elements_by_text_range(doc: Document, start_contains: str, end_contains: str,
                                   include_start=True, include_end=True):
    """
    Supprime tous les éléments entre start_contains et end_contains dans le body.
    Paramètres include_start/include_end contrôlent si les marqueurs eux-mêmes sont inclus.
    """
    body = doc.element.body
    body_children = list(body)

    in_range = False
    to_remove = []
    found_start = False

    for elem in body_children:
        text = _get_para_text(elem) if elem.tag.endswith("}p") else ""

        if not found_start and start_contains in text:
            found_start = True
            in_range = True
            if include_start:
                to_remove.append(elem)
            continue  # Don't process end check on same element as start

        if in_range:
            if end_contains and end_contains in text:
                if include_end:
                    to_remove.append(elem)
                break
            else:
                to_remove.append(elem)

    _remove_elements(body, to_remove)
    return len(to_remove)


def delete_single_paragraph(doc: Document, contains: str):
    """Supprime le premier paragraphe contenant contains."""
    body = doc.element.body
    for elem in list(body):
        if elem.tag.endswith("}p"):
            if contains in _get_para_text(elem):
                try:
                    body.remove(elem)
                except ValueError:
                    pass
                return True
    return False


def delete_section_by_title(doc: Document, section_title: str):
    """
    Supprime une section de conclusion identifiée par son titre.
    Structure : chaque section = paragraphe numPr+underline (titre),
    suivi de paragraphes de contenu, jusqu'au prochain numPr+underline ou table.
    """
    body = doc.element.body
    paras = list(body)

    target_idx = None
    for i, elem in enumerate(paras):
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            text = _get_para_text(elem)
            if section_title in text:
                target_idx = i
                break
    if target_idx is None:
        return

    to_remove = [paras[target_idx]]
    for i in range(target_idx + 1, len(paras)):
        elem = paras[i]
        # Arrêt impératif aux en-têtes majeurs (VI –, VII –, VIII –, etc.)
        if elem.tag.endswith("}p"):
            text = _get_para_text(elem)
            if _MAJOR_SECTION_RE.match(text):
                break
        # Arrêt au prochain titre de sous-section (numPr + underline)
        if elem.tag.endswith("}p") and _has_numpr(elem) and _has_underline(elem):
            break
        to_remove.append(elem)

    _remove_elements(body, to_remove)


def insert_paragraphs_before(doc: Document, anchor_contains: str, paragraphs_data: list):
    """
    Insère des paragraphes avant l'élément contenant anchor_contains.
    paragraphs_data : liste de dicts { 'text': str, 'bold': bool }
    L'ordre final est le même que l'ordre dans paragraphs_data.
    """
    body = doc.element.body
    anchor = None
    for elem in body:
        if elem.tag.endswith("}p") and anchor_contains in _get_para_text(elem):
            anchor = elem
            break
    if anchor is None:
        return

    last_inserted = None
    for pdata in paragraphs_data:
        new_para = doc.add_paragraph()
        run = new_para.add_run(pdata.get('text', ''))
        if pdata.get('bold'):
            run.bold = True
        if pdata.get('underline'):
            run.underline = True
        font_size = pdata.get('font_size')
        if font_size:
            run.font.size = Pt(font_size)
        new_elem = new_para._element
        body.remove(new_elem)
        if last_inserted is None:
            anchor.addprevious(new_elem)
        else:
            last_inserted.addnext(new_elem)
        last_inserted = new_elem


# ─────────────────────────────────────────────
# Gestion des photos
# ─────────────────────────────────────────────

def _insert_facade_photo(doc: Document, facade_bytes: bytes):
    """
    Remplace le placeholder texte de la photo de façade (P6) par l'image fournie.
    Le nouveau template utilise un placeholder texte, pas une image embarquée.
    """
    FACADE_PH = "\u2018ici photo de la fa\u00e7ade en mode portrait, 9,87cm de hauteur\u2019"
    for para in doc.paragraphs:
        if FACADE_PH in para.text:
            # Vider tous les runs du paragraphe
            for run in para.runs:
                run.text = ""
            # Insérer l'image dans un nouveau run
            try:
                run = para.add_run()
                run.add_picture(io.BytesIO(facade_bytes), height=Cm(9.87))
            except Exception as e:
                print(f"[WARN] Erreur insertion photo façade : {e}")
            return
    # Si placeholder non trouvé, on ignore silencieusement


def _fill_photo_tables(doc: Document, photos_commentees: list):
    """
    Remplit les tables de photos (Tables 5-9) avec les photos fournies.
    Chaque table a N lignes × 2 colonnes.
    Vide les cellules sans photo et supprime les lignes entièrement vides.
    """
    photo_idx = 0
    total_photos = len(photos_commentees)

    for tbl_idx in range(5, 10):
        if tbl_idx >= len(doc.tables):
            break
        tbl = doc.tables[tbl_idx]
        rows_to_remove = []

        for row_idx, row in enumerate(tbl.rows):
            row_has_photo = False
            # Chaque ligne a 2 colonnes
            n_cols = min(2, len(row.cells))
            for col_idx in range(n_cols):
                # Éviter les cellules fusionnées (même _tc)
                cell = row.cells[col_idx]
                if col_idx > 0 and row.cells[col_idx]._tc is row.cells[0]._tc:
                    continue  # cellule fusionnée, déjà traitée
                if photo_idx < total_photos:
                    _fill_photo_cell(doc, cell, photos_commentees[photo_idx])
                    photo_idx += 1
                    row_has_photo = True
                else:
                    _clear_cell(cell)

            if not row_has_photo:
                rows_to_remove.append(row._tr)

        tbl_elem = tbl._tbl
        for tr in rows_to_remove:
            try:
                tbl_elem.remove(tr)
            except Exception:
                pass

        if photo_idx >= total_photos:
            # Vider les tables restantes entièrement
            for remaining_idx in range(tbl_idx + 1, 10):
                if remaining_idx < len(doc.tables):
                    _clear_table(doc.tables[remaining_idx])
            break


def _fill_photo_cell(doc: Document, cell, photo_data: dict):
    """Remplace le contenu d'une cellule par une photo + commentaire."""
    image_b64 = photo_data.get('image_base64', '')
    commentaire = photo_data.get('commentaire', '')

    # Vider la cellule
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""

    try:
        image_bytes = base64.b64decode(image_b64)
        # Utiliser le premier paragraphe pour l'image
        para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(io.BytesIO(image_bytes), height=Cm(9.87))

        # Ajouter le commentaire dans un nouveau paragraphe
        if commentaire:
            comment_para = cell.add_paragraph(commentaire)
            comment_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in comment_para.runs:
                run.italic = True
    except Exception as e:
        print(f"[WARN] Erreur insertion photo : {e}")
        para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        para.add_run(commentaire or "Photo non disponible")


def _clear_cell(cell):
    """Vide le contenu d'une cellule."""
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""


def _clear_table(tbl):
    """Vide toutes les cellules d'une table."""
    for row in tbl.rows:
        for cell in row.cells:
            _clear_cell(cell)




# ─────────────────────────────────────────────
# Logique principale de remplissage du template
# ─────────────────────────────────────────────

def build_rapport(data: dict) -> bytes:
    """
    Prend les donnees de l'app Flutter (dict JSON) et retourne
    le contenu binaire du .docx genere.
    """
    doc = Document(TEMPLATE_PATH)

    # ── Extraction des données ────────────────────────────────────────────
    adresse        = data.get("adresseProjet", "")
    ville          = data.get("villeProjet", "")
    cp             = data.get("cpProjet", "")
    client         = data.get("client", "")
    mo_delegue     = data.get("moDelegue", "")
    adresse_mo     = data.get("adresseMoDelegue", "")
    ville_mo       = data.get("villeMoDelegue", "")
    cp_mo          = data.get("cpMoDelegue", "")
    moe            = data.get("moe", "")
    adresse_moe    = data.get("adresseMoe", "")
    ville_moe      = data.get("villeMoe", "")
    cp_moe         = data.get("cpMoe", "")
    devis          = data.get("devis", "")
    date_rapport   = data.get("dateRapport", "")
    redacteur      = data.get("redacteur", "")
    verificateur   = data.get("verificateur", "")
    titre_etude    = data.get("titreEtude", "")
    reglement      = data.get("reglementApplicable", "Ville de Paris")

    adresse_full     = f"{adresse}, {cp} {ville}".strip(", ")
    adresse_mo_full  = f"{adresse_mo}, {cp_mo} {ville_mo}".strip(", ")
    adresse_moe_full = f"{adresse_moe}, {cp_moe} {ville_moe}".strip(", ")

    desc_site      = data.get("descriptionSite", "")
    parcelle       = data.get("parcelleCadastre", "")
    section_cad    = data.get("sectionCadastre", "")
    objet_mission  = data.get("objetMission", "")

    paragraphes            = data.get("paragraphesSelectionnes", [])
    reglementations        = data.get("reglementationsSelectionnees", [])

    # BPE
    bpe_present            = data.get("bpePresent", False)
    bpe_type               = data.get("bpeTypeBranchement", "")       # 'ouvert', 'ferme', 'canalise'
    bpe_phrase             = data.get("bpePhraseGeneree", "")

    # Zones canalisations
    batiment_zones         = data.get("batimentZones", [])   # [{zoneName, appPresentes, appPhrase, entPresentes, entPhrase}]
    cour_zones             = data.get("courZones", [])        # [{zoneName, entPresentes, entPhrase}]

    is_paris = cp.startswith("75")

    # ── 1. Remplacements dans les tables (page de garde) ──────────────────

    # Table 0 : en-tête avec l'adresse
    tbl0 = doc.tables[0]
    for row in tbl0.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_text_in_paragraph(para,
                    "\u2018ADRESSE DU PROJET, CODE POSTAL VILLE\u2018",
                    adresse_full.upper())

    # Table 1 : acteurs, devis, date, rédacteur, vérificateur
    if len(doc.tables) > 1:
        tbl1 = doc.tables[1]
        # Cellule [1,0] : MOD (cells 0 et 1 sont fusionnées → même XML)
        # Les paragraphes sont : P0=vide, P1=nom MOD, P2=adresse MOD, P3=CP+ville MOD
        cell_mod = tbl1.rows[1].cells[0]
        for para in cell_mod.paragraphs:
            replace_text_in_paragraph(para,
                "\u2018ici Maitre d\u2019ouvrage d\u00e9l\u00e9gu\u00e9\u2019", mo_delegue)
            replace_text_in_paragraph(para,
                "\u2018ici adresse du maitre d\u2019ouvrage d\u00e9l\u00e9gu\u00e9\u2018", adresse_mo)
            replace_text_in_paragraph(para,
                "\u2018ici code postal du MOD\u2019 \u2018ici ville MOD\u2019",
                f"{cp_mo} {ville_mo}")

        # Cellule [1,2] : MO (maître d'œuvre)
        # Les paragraphes sont : P0=espace, P1=nom MO, P2=adresse MO, P3=CP+ville MO, P4=vide
        cell_moe = tbl1.rows[1].cells[2]
        for para in cell_moe.paragraphs:
            replace_text_in_paragraph(para,
                "\u2018ici Maitre d\u2019oeuvre\u2019", moe)
            replace_text_in_paragraph(para,
                "\u2018ici adresse du maitre d\u2019oeuvre\u2018", adresse_moe)
            replace_text_in_paragraph(para,
                "\u2018ici code postal du MO\u2019 \u2018ici ville MO\u2019",
                f"{cp_moe} {ville_moe}")

        # Ligne 3 : devis, date, rédacteur, vérificateur
        row3 = tbl1.rows[3]
        replace_text_in_paragraph(row3.cells[0].paragraphs[0],
            "\u2018ici num\u00e9ro de devis\u2019", devis)
        replace_text_in_paragraph(row3.cells[1].paragraphs[0],
            "\u2018ici mois de g\u00e9n\u00e9ration du rapport + ann\u00e9e\u2019", date_rapport)
        replace_text_in_paragraph(row3.cells[2].paragraphs[0],
            "\u2018ici r\u00e9dacteur du rapport\u2019", redacteur)
        replace_text_in_paragraph(row3.cells[3].paragraphs[0],
            "\u2018ici v\u00e9rifi\u00e9 par\u2026\u2019", verificateur)

    # ── 2. Remplacements dans les paragraphes fixes ───────────────────────
    replacements = {
        # Page de garde – SDC (le "SDC DU" est statique dans le template, on remplace uniquement le placeholder)
        "\u2018Client/maitre d\u2019ouvrage\u2019": client.upper(),

        # Titre d'étude (P9)
        "\u2018ici titre d\u2019\u00e9tude\u2019": titre_etude,

        # Règlement (P52)
        "\u2018remplir ici selon le r\u00e8glement choisi\u2019": reglement,

        # Section I – Objet de la mission
        "\u2018adresse, code postal ville\u2019": adresse_full,

        # Section II – Description du site (parcelle)
        "\u2018num\u00e9ro de parcelle cadastrale\u2019": parcelle,
        "\u2018section cadastrale\u2019": section_cad,

        # Objet de la mission (P68)
        "\u2018ce qui est coch\u00e9 ou tap\u00e9 dans la page objet de la mission\u2019": objet_mission,

        # Description du site – phrase Claude IA (P72)
        "\u2018phrase g\u00e9n\u00e9r\u00e9e avec Claude IA sur la page description du site\u2019": desc_site,
    }
    replace_in_doc(doc, replacements)

    # ── 3. Section IV – Paris vs Hors-Paris ───────────────────────────────
    PARIS_MARKER    = "\u2018si rapport dans paris mettre ce paragraphe\u00a0:\u2019"
    HORS_PARIS_MARKER = "\u2018si rapport en dehors de paris"
    V_CONCLUSIONS   = "V \u2013 CONCLUSIONS"

    if is_paris:
        # Garder le bloc Paris, supprimer le bloc Hors-Paris
        delete_single_paragraph(doc, PARIS_MARKER)
        # Supprimer de hors-Paris marker jusqu'à (mais non inclus) "V – CONCLUSIONS"
        delete_elements_by_text_range(doc,
            start_contains=HORS_PARIS_MARKER,
            end_contains=V_CONCLUSIONS,
            include_start=True,
            include_end=False)
    else:
        # Supprimer le bloc Paris (de son marker jusqu'à hors-Paris marker exclus)
        delete_elements_by_text_range(doc,
            start_contains=PARIS_MARKER,
            end_contains=HORS_PARIS_MARKER,
            include_start=True,
            include_end=False)
        # Supprimer le marker hors-Paris uniquement
        delete_single_paragraph(doc, HORS_PARIS_MARKER)

    # ── 4. Section V – BPE (Branchement Particulier à l'Égout) ───────────
    BPE_MARKER    = "\u2018si dans la page de l\u2019application Branchement"
    BPE_PHRASE_PH = "\u2018phrase de conclusion en fonction des cases \u00e0 cocher"
    CANALAPP_PH   = "Canalisations apparentes en caves b\u00e2timent"

    if not bpe_present:
        # Supprimer tout le bloc BPE (marker + contenu jusqu'au paragraphe suivant après BPE_PHRASE_PH)
        delete_elements_by_text_range(doc,
            start_contains=BPE_MARKER,
            end_contains=CANALAPP_PH,
            include_start=True,
            include_end=False)
    else:
        # BPE présent : supprimer uniquement le marker (garder le titre inclus dans P113)
        # P113 contient : marker + titre "Branchement particulier à l'égout (BPE)"
        # On remplace le marker (+ l'espace qui suit) par vide dans ce paragraphe
        for para in doc.paragraphs:
            if "si dans la page de l" in para.text and "BPE" in para.text:
                replace_text_in_paragraph(para,
                    "\u2018si dans la page de l\u2019application Branchement du Particulier \u00e0 l\u2019\u00e9gout on coche pr\u00e9sence d\u2019un BPE (Paris)\u00a0: oui faire apparaitre ce titre et le texte ci dessous\u00a0:\u2019 ",
                    "")
                # Si le marqueur existe sans espace après :
                replace_text_in_paragraph(para,
                    "\u2018si dans la page de l\u2019application Branchement du Particulier \u00e0 l\u2019\u00e9gout on coche pr\u00e9sence d\u2019un BPE (Paris)\u00a0: oui faire apparaitre ce titre et le texte ci dessous\u00a0:\u2019",
                    "")
                break

        # Gestion des schémas : garder uniquement celui correspondant au type
        # Fermé : P119-P124 | Ouvert : P126-P131 | Canalisé : P134-P138
        FERME_SCHEMA_CAPTION  = "Sch\u00e9ma de principe du branchement particulier ferm\u00e9"
        OUVERT_SCHEMA_CAPTION = "Sch\u00e9ma de principe du branchement particulier ouvert"
        CANAL_SCHEMA_CAPTION  = "Sch\u00e9ma de principe du branchement particulier canalis\u00e9"
        FERME_INSTR  = "\u2018mettre le sch\u00e9ma et la l\u00e9gende ci dessus si dans la page branchement du particulier \u00e0 l\u2019\u00e9gout on coche BPE ferm\u00e9\u2019"
        OUVERT_INSTR = "\u2018mettre le sch\u00e9ma et la l\u00e9gende ci dessus si dans la page branchement du particulier \u00e0 l\u2019\u00e9gout on coche BPE ouvert\u2019"
        CANAL_INSTR  = "\u2018mettre le sch\u00e9ma et la l\u00e9gende ci dessus si dans la page branchement du particulier \u00e0 l\u2019\u00e9gout on coche BPE canalis\u00e9\u2019"

        if bpe_type == 'ferme':
            # Garder fermé, supprimer ouvert et canalisé
            delete_single_paragraph(doc, FERME_INSTR)
            delete_elements_by_text_range(doc,
                start_contains=OUVERT_SCHEMA_CAPTION,
                end_contains=OUVERT_INSTR,
                include_start=True, include_end=True)
            delete_elements_by_text_range(doc,
                start_contains=CANAL_SCHEMA_CAPTION,
                end_contains=CANAL_INSTR,
                include_start=True, include_end=True)
            # Supprimer aussi les lignes vides orphelines entre les schémas
        elif bpe_type == 'ouvert':
            # Garder ouvert, supprimer fermé et canalisé
            delete_elements_by_text_range(doc,
                start_contains=FERME_SCHEMA_CAPTION,
                end_contains=FERME_INSTR,
                include_start=True, include_end=True)
            delete_single_paragraph(doc, OUVERT_INSTR)
            delete_elements_by_text_range(doc,
                start_contains=CANAL_SCHEMA_CAPTION,
                end_contains=CANAL_INSTR,
                include_start=True, include_end=True)
        elif bpe_type == 'canalise':
            # Garder canalisé, supprimer fermé et ouvert
            delete_elements_by_text_range(doc,
                start_contains=FERME_SCHEMA_CAPTION,
                end_contains=FERME_INSTR,
                include_start=True, include_end=True)
            delete_elements_by_text_range(doc,
                start_contains=OUVERT_SCHEMA_CAPTION,
                end_contains=OUVERT_INSTR,
                include_start=True, include_end=True)
            delete_single_paragraph(doc, CANAL_INSTR)
        else:
            # Type non défini : supprimer tous les schémas
            delete_elements_by_text_range(doc,
                start_contains=FERME_SCHEMA_CAPTION,
                end_contains=FERME_INSTR,
                include_start=True, include_end=True)
            delete_elements_by_text_range(doc,
                start_contains=OUVERT_SCHEMA_CAPTION,
                end_contains=OUVERT_INSTR,
                include_start=True, include_end=True)
            delete_elements_by_text_range(doc,
                start_contains=CANAL_SCHEMA_CAPTION,
                end_contains=CANAL_INSTR,
                include_start=True, include_end=True)

        # Insérer la phrase BPE générée par Claude
        for para in doc.paragraphs:
            if BPE_PHRASE_PH.split('\u00e0')[0][1:] in para.text:
                replace_text_in_paragraph(para,
                    para.text,
                    bpe_phrase)
                break

    # ── 5. Section V – Zones de canalisations ─────────────────────────────
    # Supprimer les paragraphes génériques (143-151) et insérer les vrais
    CANALAPP_GENERIC  = "Canalisations apparentes en caves b\u00e2timent \u2026."
    CANALENT_BAT      = "Canalisations enterr\u00e9es sous b\u00e2timent"
    CANALENT_EXT      = "Canalisations enterr\u00e9es sous espaces exterieurs"
    CANALAPP_PHRASE   = "\u2018phrase de conclusion g\u00e9n\u00e9r\u00e9e par Claude IA\u2019"
    CANALAPP_INSTR    = "\u2018faire plusieurs paragraphes si il y"
    CANALENT_INSTR    = "\u2018faire plusieurs paragraphes si il y"

    # Supprimer les 3 blocs génériques (apparentes bat, enterrées bat, enterrées ext)
    # On les supprime du premier (CANALAPP_GENERIC) jusqu'au titre de conclusion suivant (Colonne EP)
    COLONNE_EP_TITLE = "Colonne d\u2019eaux pluviales de fa\u00e7ade"

    delete_elements_by_text_range(doc,
        start_contains=CANALAPP_GENERIC,
        end_contains=COLONNE_EP_TITLE,
        include_start=True,
        include_end=False)

    # Insérer les paragraphes de zones avant le premier titre de conclusion (Colonne EP)
    # ou avant "V – CONCLUSIONS" si aucune conclusion
    zone_paragraphs = []

    show_app = "Canalisations Apparentes" in paragraphes
    show_ent = "Canalisations Enterrées" in paragraphes

    # Bâtiments
    for zone in batiment_zones:
        zone_name = zone.get("zoneName", "Bâtiment")
        app_presentes = zone.get("appPresentes", False)
        app_phrase    = zone.get("appPhrase", "")
        ent_presentes = zone.get("entPresentes", False)
        ent_phrase    = zone.get("entPhrase", "")

        if show_app and app_presentes and app_phrase:
            zone_paragraphs.append({'text': f"Canalisations apparentes en caves {zone_name}", 'bold': True, 'underline': True})
            zone_paragraphs.append({'text': app_phrase, 'bold': False})

        if show_ent and ent_presentes and ent_phrase:
            zone_paragraphs.append({'text': f"Canalisations enterr\u00e9es sous {zone_name}", 'bold': True, 'underline': True})
            zone_paragraphs.append({'text': ent_phrase, 'bold': False})

    # Cours
    for zone in cour_zones:
        zone_name = zone.get("zoneName", "espace extérieur")
        ent_presentes = zone.get("entPresentes", False)
        ent_phrase    = zone.get("entPhrase", "")

        if show_ent and ent_presentes and ent_phrase:
            zone_paragraphs.append({'text': f"Canalisations enterr\u00e9es sous espaces ext\u00e9rieurs ({zone_name})", 'bold': True, 'underline': True})
            zone_paragraphs.append({'text': ent_phrase, 'bold': False})

    if zone_paragraphs:
        # Insérer avant le premier titre de conclusion (ou avant VI si aucun)
        anchor = None
        for candidate in [COLONNE_EP_TITLE, "Regard de limite de propri\u00e9t\u00e9",
                           "Installations sanitaires en sous-sol",
                           "Ancienne fosse d\u2019aisance",
                           "Regards de visite non \u00e9tanches",
                           "R\u00e9seau s\u00e9paratif",
                           "Cas des eaux us\u00e9es provenant",
                           "Cas des eaux provenant des garages",
                           "VI \u2013 COMPL\u00c9MENT"]:
            for para in doc.paragraphs:
                if candidate in para.text:
                    anchor = candidate
                    break
            if anchor:
                break

        if anchor:
            insert_paragraphs_before(doc, anchor, zone_paragraphs)

    # ── 6. Sections de conclusions – supprimer les non sélectionnées ──────
    # Map : chaîne Flutter exacte → extrait du titre dans le template
    SECTION_MAP = {
        "Colonne d\u2019eaux pluviales de fa\u00e7ade":               "Colonne d\u2019eaux pluviales de fa\u00e7ade",
        "Regard de limite de propri\u00e9t\u00e9":                    "Regard de limite de propri\u00e9t\u00e9",
        "Installations sanitaires en sous-sol":                       "Installations sanitaires en sous-sol",
        "Ancienne fosse d\u2019aisance":                              "Ancienne fosse d\u2019aisance",
        "Regards de visite non \u00e9tanches":                        "Regards de visite non \u00e9tanches",
        "R\u00e9seau s\u00e9paratif":                                 "R\u00e9seau s\u00e9paratif",
        "Cas des eaux us\u00e9es \u2013 Restaurants":                 "Cas des eaux us\u00e9es provenant des restaurants",
        "Cas des eaux \u2013 Garages & Stations de lavage":           "Cas des eaux provenant des garages",
    }

    # "Ventilations des réseaux" est lié à "Réseau séparatif"
    keep_ventilation = "R\u00e9seau s\u00e9paratif" in reglementations

    for flutter_key, template_title in SECTION_MAP.items():
        if flutter_key not in reglementations:
            delete_section_by_title(doc, template_title)

    if not keep_ventilation:
        delete_section_by_title(doc, "Ventilations des r\u00e9seaux")

    # ── 7. Supprimer les instructions résiduelles dans les titres ─────────
    # Les titres de sections de conclusion contiennent des instructions en smart quotes
    # qu'il faut supprimer si la section est gardée
    for para in doc.paragraphs:
        if _has_numpr(para._element) and _has_underline(para._element):
            t = para.text
            # Supprimer la partie instruction (entre \u2018 et fin de phrase)
            # En remplaçant le placeholder d'instruction
            for instr_marker in [
                " \u2018int\u00e9grer ce paragraphe si coch\u00e9 dans selection des paragraphes\u2019",
                "\u2018int\u00e9grer ce paragraphe si coch\u00e9 dans selection des paragraphes\u2019",
            ]:
                if instr_marker in t:
                    replace_text_in_paragraph(para, instr_marker, "")
                    break

    # ── 8. Règlement d'assainissement ─────────────────────────────────────
    # (garder le texte d'origine si règlement = Paris, sinon adapter)
    if reglement and "Paris" not in reglement:
        replace_in_doc(doc, {
            "r\u00e8glement d\u2019assainissement de la ville de Paris":
                f"r\u00e8glement d\u2019assainissement de {reglement}",
        })

    # ── 9. Photo de façade (placeholder texte P6 → image) ────────────────
    FACADE_PH = "\u2018ici photo de la fa\u00e7ade en mode portrait, 9,87cm de hauteur\u2019"
    photo_facade_b64 = data.get("photoFacade")
    if photo_facade_b64:
        try:
            facade_bytes = base64.b64decode(photo_facade_b64)
            _insert_facade_photo(doc, facade_bytes)
        except Exception as e:
            print(f"[WARN] Erreur décodage photo façade : {e}")
    else:
        # Effacer le placeholder si aucune photo fournie
        for para in doc.paragraphs:
            if FACADE_PH in para.text:
                replace_text_in_paragraph(para, FACADE_PH, "")
                break

    # ── 10. Photos avec commentaires (section VI) ─────────────────────────
    photos_commentees = data.get("photosCommentees", [])
    _fill_photo_tables(doc, photos_commentees)

    # ── 11. Sérialise en mémoire ──────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Envoi du mail
# ─────────────────────────────────────────────

def send_email(docx_bytes: bytes, filename: str, devis: str, adresse: str):
    """Envoie le .docx via l'API HTTP Brevo (pas de SMTP, pas de blocage réseau)."""
    if not BREVO_API_KEY:
        raise Exception("BREVO_API_KEY manquante dans les variables d'environnement")

    body_text = (
        f"Bonjour,\n\n"
        f"Le rapport [{filename}] a été correctement généré et peut être téléchargé en pièce jointe.\n\n"
        f"Devis : {devis}\n"
        f"Adresse : {adresse}\n\n"
        f"Application Immeau."
    )

    payload = {
        "sender":      {"name": "Immeau", "email": MAIL_FROM},
        "to":          [{"email": MAIL_TO}],
        "subject":     f"Création du rapport [{filename.replace('.docx', '')}]",
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

        devis    = data.get("devis", "00000")
        adresse  = data.get("adresseProjet", "")
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
