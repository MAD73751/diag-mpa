#!/usr/bin/env python3
"""
serveur_diagnostic_mpa.py — EI David MARIE-ANNE / DMAAI
Serveur local Flask pour la génération de rapports MaPrimeAdapt'
Lance avec : python serveur_diagnostic_mpa.py
Puis ouvre : agent_diagnostic_mpa.html dans Firefox
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap
from datetime import datetime
from logos_b64 import LOGO_DMAAI_B64, LOGO_MPA_B64
import json, io, sys, base64, re, traceback
import os

# ─── RÈGLES ANAH : référentiel partagé (regles_anah_2026.json, même dossier que ce serveur) ───
# Aucune valeur d'aide en dur dans ce fichier : tout vient du JSON généré/maintenu avec build_regles.py.
_REGLES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'regles_anah_2026.json')
try:
    with open(_REGLES_PATH, encoding='utf-8') as _fh:
        REGLES = json.load(_fh)
except Exception as _e:
    sys.exit(f"Référentiel des règles Anah introuvable ou illisible ({_REGLES_PATH}) : {_e}\n"
             "Copier regles_anah_2026.json dans le dossier du serveur.")

def plafond_ressources(categorie, personnes, idf):
    """Plafond de ressources 2026 (même logique que regles_anah.js)."""
    zone = REGLES['plafonds_ressources']['ile_de_france' if idf else 'hors_ile_de_france']
    t_ = zone[categorie]
    n = max(1, int(personnes))
    base = t_['1_a_5_personnes']
    return base[n - 1] if n <= 5 else base[4] + (n - 5) * t_['par_personne_supplementaire']

def eur_txt(n):
    return f"{n:,.0f}".replace(',', '\u202f') + ' €'

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# ─── COULEURS DMAAI ───────────────────────────────────────────────────────────
BLEU       = RGBColor(0x1F, 0x4E, 0x79)
BLEU_TITRE = RGBColor(0x2E, 0x75, 0xB6)
BLEU_CLAIR = RGBColor(0xD6, 0xE4, 0xF0)
BLANC      = RGBColor(0xFF, 0xFF, 0xFF)
ROUGE      = RGBColor(0xC0, 0x39, 0x2B)
VERT       = RGBColor(0x1E, 0x84, 0x49)
ORANGE     = RGBColor(0xD3, 0x54, 0x00)
GRIS_T     = RGBColor(0x55, 0x65, 0x75)
NOIR       = RGBColor(0x1A, 0x1A, 0x2E)

PIECES = [
    'Cheminement extérieur', 'Parties communes', 'Garage / stationnement',
    'Entrée', 'Salon', 'Salle à manger', 'Salle de bain', 'Sanitaires',
    'Chambre 1 (usager)', 'Chambre 2', 'Cuisine', 'Escalier', 'Couloirs', 'Autre'
]

# ─── HABILITATION ANAH ────────────────────────────────────────────────────────
HABILIT_REF  = 'HABILIT./MARIE-ANNE/2024/06'
HABILIT_DATE = '12 décembre 2024'
HABILIT_DEPT = 'Loiret (45)'

# ─── NORMALISATION JSON (PWA → agent) ────────────────────────────────────────
def normaliser(data: dict) -> dict:
    """
    Accepte les deux formats JSON :
    - Format PWA GitHub : { data: { fields, radios, checks, pieces, trav }, photos }
    - Format agent local : { flds, rad, chk, pieces, trav, photos }
    Retourne toujours le format agent.
    """
    # Format PWA
    if 'data' in data and isinstance(data['data'], dict):
        src = data['data']
        f   = src.get('fields', {})
        r   = src.get('radios', {})
        c   = src.get('checks', {})
        p   = src.get('pieces', {})
        tv  = src.get('trav', {})
        ph  = data.get('photos', {})

        flds = {
            'amo_rso':    f.get('amo_structure', 'EI David MARIE-ANNE'),
            'amo_diag':   f.get('amo_diagnostiqueur', 'David MARIE-ANNE'),
            'amo_fonc':   f.get('amo_fonction', 'Architecte DPLG / Ingénieur TCE'),
            'v_date':     f.get('visite_date', ''),
            'd_nom':      f.get('dem_nom', ''),
            'd_prn':      f.get('dem_prenom', ''),
            'd_nais':     f.get('dem_naissance', ''),
            'd_ret':      f.get('dem_retraite', ''),
            'l_adr':      f.get('log_adresse', ''),
            'l_an':       f.get('log_annee', ''),
            'l_surf':     f.get('log_surface', ''),
            'l_pces':     f.get('log_nb_pieces', ''),
            'l_niv':      f.get('log_niveaux', ''),
            'l_etg':      f.get('log_etage', ''),
            'l_occ':      f.get('log_nb_occupants', ''),
            'fr_ct':      f.get('fr_coords', ''),
            'atec_autre': f.get('aides_tech_autre', ''),
            'conf_id':    f.get('confiance_identite', ''),
            'obs_aut':    f.get('obs_autonomie', ''),
            'e_sal': '', 'e_pen': '', 'e_alloc': '', 'e_autre': '', 'e_ep': '',
            'c_loy': '', 'c_pret': '', 'c_cred': '', 'c_mut': '',
            'c_elec': '', 'c_ass': '', 'c_autre': '',
            'prec':       f.get('preconisations', ''),
            'crp':        f.get('creation_piece', ''),
            'rem_men':    f.get('remarques_menage', ''),
        }

        rad = {
            'tlog':  'Maison individuelle' if r.get('type_logement') == 'Maison individuelle'
                     else 'Bâtiment collectif' if r.get('type_logement') in ('BHC','Bâtiment collectif')
                     else r.get('type_logement', ''),
            'stat':  r.get('statut_occupation', ''),
            'zone':  r.get('zone', ''),
            'fr':    r.get('fr_renov', ''),
            'handi': r.get('handicap', ''),
            'gir':   '' if r.get('gir') in ('Non évalué', None, '') else r.get('gir', ''),
            'gir_raw': r.get('gir', ''),
            'chute': r.get('chute', ''),
            'hosp':  r.get('hospit', ''),
            'paut':  r.get('perte_aut', ''),
            'ergo':  r.get('ergo', ''),
            'rergo': r.get('rapport_ergo', ''),
            'aid2':  r.get('aidant', ''),
            'conf':  r.get('confiance', ''),
        }

        chk = {
            'sit':  c.get('situation', []),
            'aid':  c.get('aides', []),
            'ori':  c.get('orientation', []),
            'atec': c.get('aides_tech', []),
            'eval': c.get('eval_log', []),
        }

        # Normalisation pièces — PWA utilise statut/etage/remarque
        pieces_norm = {}
        trav_norm   = {}
        for piece in PIECES:
            pd = p.get(piece, {})
            # Statut : 'adapte'/'non-adapte'/'' (PWA) ou 'ok'/'nok'/'st' (agent)
            st_raw = pd.get('statut', pd.get('st', ''))
            if st_raw in ('adapte', 'ok', 'Adapté'):
                st = 'ok'
            elif st_raw in ('non-adapte', 'nok', 'Non adapté'):
                st = 'nok'
            else:
                st = ''

            pieces_norm[piece] = {
                'st':    st,
                'circ':  pd.get('circ', ''),
                'equip': pd.get('equip', ''),
                'etg':   pd.get('etage', pd.get('etg', '')),
                'rem':   pd.get('remarque', pd.get('rem', '')),
            }

            td = tv.get(piece, {})
            trav_norm[piece] = {
                'prio': td.get('priorite', td.get('prio', '')),
                'sol':  td.get('solutions', td.get('sol', '')),
                'ret':  td.get('retenues', td.get('ret', '')),
            }

        return {
            'flds': flds, 'rad': rad, 'chk': chk,
            'pieces': pieces_norm, 'trav': trav_norm,
            'photos': ph
        }

    # Format agent local — déjà bon, on normalise juste les pièces
    else:
        pieces_norm = {}
        trav_norm   = {}
        for piece in PIECES:
            pd = data.get('pieces', {}).get(piece, {})
            st_raw = pd.get('st', pd.get('statut', ''))
            if st_raw in ('adapte', 'ok'):
                st = 'ok'
            elif st_raw in ('non-adapte', 'nok'):
                st = 'nok'
            else:
                st = ''
            pieces_norm[piece] = {
                'st': st,
                'circ': pd.get('circ', ''),
                'equip': pd.get('equip', ''),
                'etg': pd.get('etg', pd.get('etage', '')),
                'rem': pd.get('rem', pd.get('remarque', '')),
            }
            td = data.get('trav', {}).get(piece, {})
            trav_norm[piece] = {
                'prio': td.get('prio', td.get('priorite', '')),
                'sol':  td.get('sol', td.get('solutions', '')),
                'ret':  td.get('ret', td.get('retenues', '')),
            }
        data['pieces'] = pieces_norm
        data['trav']   = trav_norm
        return data


# ─── PHOTOS ───────────────────────────────────────────────────────────────────
def b64_to_bytes(data_uri):
    try:
        m = re.match(r'data:(image/[\w+]+);base64,(.+)', data_uri, re.DOTALL)
        if not m:
            return None, None
        mime = m.group(1)
        raw  = base64.b64decode(m.group(2))
        ext  = mime.split('/')[1].replace('jpeg','jpeg').replace('jpg','jpeg')
        return raw, ext
    except Exception:
        return None, None

def add_photos_in_cell(cell, photo_list, max_w_cm=2.8, max_count=4):
    """Ajoute des photos dans une cellule — 2 par ligne."""
    added = 0
    para  = None
    for ph in photo_list:
        if added >= max_count:
            break
        raw, ext = b64_to_bytes(ph.get('data', ''))
        if raw is None:
            continue
        # Nouvelle ligne tous les 2
        if added % 2 == 0:
            para = cell.add_paragraph()
            para.paragraph_format.space_before = Pt(2)
            para.paragraph_format.space_after  = Pt(2)
        try:
            run = para.add_run()
            run.add_picture(io.BytesIO(raw), width=Cm(max_w_cm))
            if added % 2 == 0:
                para.add_run(' ')
            added += 1
        except Exception:
            run = para.add_run(f'[Photo {added+1}]')
            run.font.size = Pt(7)
            run.font.color.rgb = GRIS_T
            added += 1

def photos_par_piece(photos_dict):
    """Retourne un dict {piece_name: [photos]} depuis le dict photos."""
    result = {}
    for i, pn in enumerate(PIECES):
        key = f'piece_{i}'
        result[pn] = photos_dict.get(key, [])
    return result


# ─── HELPERS DOCX ─────────────────────────────────────────────────────────────
def cell_bg(cell, hex_col):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_col)
    tcPr.append(shd)

def cell_borders(cell, color='D0D0D0', all_sides=True):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcB = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcB.append(b)
    tcPr.append(tcB)

def cell_pad(cell, top=80, right=100, bottom=80, left=100):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for side, val in [('top',top),('right',right),('bottom',bottom),('left',left)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        mar.append(m)
    tcPr.append(mar)

def add_sep(doc, color='1F4E79', thick=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single')
    bot.set(qn('w:sz'), str(thick))
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), color)
    pBdr.append(bot)
    pPr.append(pBdr)

def add_h1(doc, text, space_before=14):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(5)
    # Fond bleu simulé avec bordure basse épaisse
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '8')
    bot.set(qn('w:space'), '2'); bot.set(qn('w:color'), '1F4E79')
    pBdr.append(bot); pPr.append(pBdr)
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(12); r.font.color.rgb = BLEU
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = BLEU_TITRE
    return p

def add_field(doc, label, value):
    v = str(value).strip() if value else ''
    if not v or v in ('—', 'None', 'null', '0'):
        return
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    r1 = p.add_run(f'{label} : ')
    r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = GRIS_T
    r2 = p.add_run(v)
    r2.font.size = Pt(9); r2.font.color.rgb = NOIR

def tbl_header_row(table, headers, widths_cm, bg='1F4E79'):
    row = table.rows[0]
    for i, (h, w) in enumerate(zip(headers, widths_cm)):
        cell = row.cells[i]
        cell.width = Cm(w)
        cell_bg(cell, bg)
        cell_borders(cell, bg)
        cell_pad(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True; run.font.size = Pt(8); run.font.color.rgb = BLANC


# ─── GÉNÉRATION DOCX ──────────────────────────────────────────────────────────
def generer_docx(raw_data: dict) -> bytes:
    data   = normaliser(raw_data)
    f      = data.get('flds', {})
    r      = data.get('rad', {})
    ch     = data.get('chk', {})
    pieces = data.get('pieces', {})
    trav   = data.get('trav', {})
    ph_all = data.get('photos', {})
    ph_ext = ph_all.get('ext', [])
    ph_trav_gen = ph_all.get('trav', [])
    ph_pieces = photos_par_piece(ph_all)

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(2.2)
        sec.right_margin  = Cm(1.8)

    sn = doc.styles['Normal']
    sn.font.name = 'Calibri'
    sn.font.size = Pt(9)

    # ── NUMÉROTATION DES PAGES ───────────────────────────────────────────────
    def add_page_numbers(section):
        footer = section.footer
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.space_before = Pt(4)
        # "Page X / Y"
        run1 = fp.add_run('Page ')
        run1.font.size = Pt(8); run1.font.color.rgb = GRIS_T
        # Champ PAGE
        fld_page = OxmlElement('w:fldChar')
        fld_page.set(qn('w:fldCharType'), 'begin')
        r_p = OxmlElement('w:r')
        r_p.append(fld_page)
        fp._p.append(r_p)
        r_pi = OxmlElement('w:r')
        it = OxmlElement('w:instrText')
        it.text = ' PAGE '
        it.set(qn('xml:space'), 'preserve')
        r_pi.append(it)
        fp._p.append(r_pi)
        r_pe = OxmlElement('w:r')
        fe = OxmlElement('w:fldChar')
        fe.set(qn('w:fldCharType'), 'end')
        r_pe.append(fe)
        fp._p.append(r_pe)
        run2 = fp.add_run(' / ')
        run2.font.size = Pt(8); run2.font.color.rgb = GRIS_T
        # Champ NUMPAGES
        fld_np = OxmlElement('w:fldChar')
        fld_np.set(qn('w:fldCharType'), 'begin')
        r_np = OxmlElement('w:r')
        r_np.append(fld_np)
        fp._p.append(r_np)
        r_npi = OxmlElement('w:r')
        itn = OxmlElement('w:instrText')
        itn.text = ' NUMPAGES '
        itn.set(qn('xml:space'), 'preserve')
        r_npi.append(itn)
        fp._p.append(r_npi)
        r_npe = OxmlElement('w:r')
        fen = OxmlElement('w:fldChar')
        fen.set(qn('w:fldCharType'), 'end')
        r_npe.append(fen)
        fp._p.append(r_npe)
        # Ligne latérale : ref dossier + habilitation
        fp2 = footer.add_paragraph()
        fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rh = fp2.add_run(f'Habilitation ANAH {HABILIT_REF} · {HABILIT_DATE} · {HABILIT_DEPT}')
        rh.font.size = Pt(7); rh.font.color.rgb = GRIS_T

    for sec in doc.sections:
        add_page_numbers(sec)

    # ── EN-TÊTE ───────────────────────────────────────────────────────────────
    # Logos + titre sur 3 colonnes
    t_hdr = doc.add_table(rows=1, cols=3)
    t_hdr.style = 'Table Grid'
    t_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Col gauche : logo DMAAI
    c_logo = t_hdr.rows[0].cells[0]
    c_logo.width = Cm(3.8)
    cell_bg(c_logo, 'FFFFFF')
    cell_borders(c_logo, 'FFFFFF')
    cell_pad(c_logo, 60, 60, 60, 60)
    try:
        logo_bytes = base64.b64decode(LOGO_DMAAI_B64)
        p_logo = c_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(io.BytesIO(logo_bytes), width=Cm(3.4))
    except Exception as e:
        c_logo.paragraphs[0].add_run('DMAAI').bold = True

    # Col centrale : titre bleu
    c_left = t_hdr.rows[0].cells[1]
    c_left.width = Cm(10)
    cell_bg(c_left, '1F4E79')
    cell_borders(c_left, '1F4E79')
    cell_pad(c_left, 100, 120, 80, 140)
    p_t = c_left.paragraphs[0]
    rt = p_t.add_run('DIAGNOSTIC LOGEMENT AUTONOMIE')
    rt.bold = True; rt.font.size = Pt(13); rt.font.color.rgb = BLANC
    p_t2 = c_left.add_paragraph()
    rt2 = p_t2.add_run("MaPrimeAdapt'")
    rt2.bold = True; rt2.font.size = Pt(11); rt2.font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)
    p_sub = c_left.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(4)
    rs = p_sub.add_run('EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911 · Agréé MAR')
    rs.font.size = Pt(7.5); rs.font.color.rgb = RGBColor(0xB0, 0xC8, 0xE8)
    # Habilitation ANAH
    p_hab = c_left.add_paragraph()
    p_hab.paragraph_format.space_before = Pt(3)
    rh1 = p_hab.add_run(f'Habilitation ANAH N° {HABILIT_REF}')
    rh1.font.size = Pt(7.5); rh1.bold = True; rh1.font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)
    rh2 = p_hab.add_run(f'  ·  accordée le {HABILIT_DATE}  ·  Département du {HABILIT_DEPT}')
    rh2.font.size = Pt(7); rh2.font.color.rgb = RGBColor(0xB0, 0xC8, 0xE8)

    # Col droite : logo MPA + date
    c_right = t_hdr.rows[0].cells[2]
    c_right.width = Cm(3.7)
    cell_bg(c_right, 'FFFFFF')
    cell_borders(c_right, 'FFFFFF')
    cell_pad(c_right, 60, 60, 60, 60)
    try:
        mpa_bytes = base64.b64decode(LOGO_MPA_B64)
        p_mpa = c_right.paragraphs[0]
        p_mpa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_mpa = p_mpa.add_run()
        run_mpa.add_picture(io.BytesIO(mpa_bytes), width=Cm(3.2))
    except Exception:
        c_right.paragraphs[0].add_run("MaPrimeAdapt'")

    date_str = '—'
    if f.get('v_date'):
        try:
            date_str = datetime.strptime(f['v_date'], '%Y-%m-%d').strftime('%d/%m/%Y')
        except:
            date_str = f['v_date']
    p_date = c_right.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(4)
    rd1 = p_date.add_run('Visite du\n')
    rd1.font.size = Pt(7.5); rd1.font.color.rgb = GRIS_T
    rd2 = p_date.add_run(date_str)
    rd2.bold = True; rd2.font.size = Pt(10); rd2.font.color.rgb = BLEU

    # Référence dossier
    ref_dossier = f.get('ref_dossier', '')
    if ref_dossier:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.space_before = Pt(4)
        p_ref.paragraph_format.space_after  = Pt(2)
        rr1 = p_ref.add_run('Référence dossier : ')
        rr1.bold = True; rr1.font.size = Pt(9); rr1.font.color.rgb = BLEU
        rr2 = p_ref.add_run(ref_dossier)
        rr2.bold = True; rr2.font.size = Pt(9); rr2.font.color.rgb = NOIR

    doc.add_paragraph()  # espace

    # ── SOMMAIRE ─────────────────────────────────────────────────────────────
    p_toc_title = doc.add_paragraph()
    p_toc_title.paragraph_format.space_before = Pt(2)
    p_toc_title.paragraph_format.space_after  = Pt(4)
    rt = p_toc_title.add_run('SOMMAIRE')
    rt.bold = True; rt.font.size = Pt(10); rt.font.color.rgb = BLEU_TITRE

    sommaire_items = [
        ('1', 'Informations générales'),
        ('2', "Situation d'autonomie"),
        ('3', 'Situation économique (facultatif)'),
        ('4', 'État des lieux global du logement'),
        ('5', 'Projet de travaux'),
    ]
    for num, titre in sommaire_items:
        p_s = doc.add_paragraph()
        p_s.paragraph_format.space_before = Pt(1)
        p_s.paragraph_format.space_after  = Pt(1)
        p_s.paragraph_format.left_indent  = Cm(0.5)
        run_num = p_s.add_run(f'Partie {num}  —  ')
        run_num.bold = True; run_num.font.size = Pt(9); run_num.font.color.rgb = BLEU
        run_tit = p_s.add_run(titre)
        run_tit.font.size = Pt(9); run_tit.font.color.rgb = GRIS_T

    add_sep(doc, '2E75B6', 4)

    # Photos extérieures sous l'en-tête
    if ph_ext:
        p_ph = doc.add_paragraph()
        p_ph.paragraph_format.space_before = Pt(2)
        p_ph.paragraph_format.space_after  = Pt(6)
        for ph in ph_ext[:4]:
            raw, _ = b64_to_bytes(ph.get('data',''))
            if raw:
                try:
                    run = p_ph.add_run()
                    run.add_picture(io.BytesIO(raw), width=Cm(3.8))
                    p_ph.add_run('  ')
                except Exception:
                    pass

    # ── PARTIE 1 ─────────────────────────────────────────────────────────────
    add_h1(doc, 'PARTIE 1 — INFORMATIONS GÉNÉRALES', space_before=6)

    # Tableau 2 colonnes info
    t_info = doc.add_table(rows=1, cols=2)
    t_info.style = 'Table Grid'
    c1 = t_info.rows[0].cells[0]
    c2 = t_info.rows[0].cells[1]
    c1.width = Cm(8.5); c2.width = Cm(9)
    for c in [c1, c2]:
        cell_bg(c, 'EBF3FB'); cell_borders(c, 'D6E4F0')
        cell_pad(c, 80, 120, 80, 120)

    # Colonne gauche : AMO + Demandeur
    p_amo = c1.paragraphs[0]
    ra = p_amo.add_run('STRUCTURE AMO')
    ra.bold = True; ra.font.size = Pt(8); ra.font.color.rgb = BLEU_TITRE

    def add_kv(cell, label, value):
        if not value or str(value).strip() in ('', '—', 'None', '0'):
            return
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        r1 = p.add_run(f'{label} : ')
        r1.bold = True; r1.font.size = Pt(8.5); r1.font.color.rgb = GRIS_T
        r2 = p.add_run(str(value).strip())
        r2.font.size = Pt(8.5); r2.font.color.rgb = NOIR

    add_kv(c1, 'Raison sociale', f.get('amo_rso'))
    add_kv(c1, 'Diagnostiqueur', f.get('amo_diag'))
    add_kv(c1, 'Fonction', f.get('amo_fonc'))

    p_dem = c1.add_paragraph()
    p_dem.paragraph_format.space_before = Pt(8)
    rd = p_dem.add_run('DEMANDEUR')
    rd.bold = True; rd.font.size = Pt(8); rd.font.color.rgb = BLEU_TITRE

    nom_complet = f'{f.get("d_nom","")}'.strip() + ' ' + f'{f.get("d_prn","")}'.strip()
    add_kv(c1, 'Identité', nom_complet.strip())
    if f.get('d_nais'):
        try:
            nais_fmt = datetime.strptime(f['d_nais'], '%Y-%m-%d').strftime('%d/%m/%Y')
        except:
            nais_fmt = f['d_nais']
        add_kv(c1, 'Date de naissance', nais_fmt)
    sit = ', '.join(ch.get('sit', [])) or None
    add_kv(c1, 'Situation', sit)
    add_kv(c1, 'Aides perçues', ', '.join(ch.get('aid', [])) or None)
    add_kv(c1, 'Retraite', f.get('d_ret'))

    # Colonne droite : Logement
    p_log = c2.paragraphs[0]
    rl = p_log.add_run('LOGEMENT')
    rl.bold = True; rl.font.size = Pt(8); rl.font.color.rgb = BLEU_TITRE

    add_kv(c2, 'Adresse', f.get('l_adr'))
    add_kv(c2, 'Type', r.get('tlog'))
    add_kv(c2, "Statut d'occupation", r.get('stat'))
    add_kv(c2, 'Année construction', f.get('l_an'))
    surf = f.get('l_surf', '')
    add_kv(c2, 'Surface', f'{surf} m²' if surf else None)
    add_kv(c2, 'Nb pièces', f.get('l_pces'))
    etg = f.get('l_etg', '')
    niv = f.get('l_niv', '')
    add_kv(c2, 'Étage / Niveaux', f'{etg} / {niv} niveau(x)' if etg or niv else None)
    add_kv(c2, 'Nb occupants', f.get('l_occ'))
    add_kv(c2, 'Zone', r.get('zone'))

    ori = ', '.join(ch.get('ori', [])) or None
    p_ori = c2.add_paragraph()
    p_ori.paragraph_format.space_before = Pt(8)
    ro = p_ori.add_run('ORIENTATION MaPrimeAdapt\'')
    ro.bold = True; ro.font.size = Pt(8); ro.font.color.rgb = BLEU_TITRE
    add_kv(c2, 'Canal', ori)
    add_kv(c2, "France Rénov'", r.get('fr'))

    # ── PARTIE 2 ─────────────────────────────────────────────────────────────
    add_h1(doc, "PARTIE 2 — SITUATION D'AUTONOMIE")

    t_aut = doc.add_table(rows=1, cols=2)
    t_aut.style = 'Table Grid'
    a1 = t_aut.rows[0].cells[0]
    a2 = t_aut.rows[0].cells[1]
    a1.width = Cm(8.5); a2.width = Cm(9)
    for c in [a1, a2]:
        cell_bg(c, 'EBF3FB'); cell_borders(c, 'D6E4F0')
        cell_pad(c, 80, 120, 80, 120)

    p_a1 = a1.paragraphs[0]
    p_a1.add_run('AUTONOMIE & HANDICAP').bold = True
    p_a1.runs[0].font.size = Pt(8); p_a1.runs[0].font.color.rgb = BLEU_TITRE

    add_kv(a1, 'Handicap', r.get('handi'))
    gir_val = r.get('gir', '')
    gir_raw = r.get('gir_raw', '')
    if gir_raw and gir_raw not in ('', 'Non évalué'):
        add_kv(a1, 'GIR', f'GIR {gir_raw}')
    else:
        add_kv(a1, 'GIR', 'Non évalué')
    atec = ch.get('atec', [])
    if f.get('atec_autre'):
        atec = atec + [f.get('atec_autre')]
    add_kv(a1, 'Aides techniques', ', '.join(atec) if atec else '—')
    add_kv(a1, 'Évaluations récentes', ', '.join(ch.get('eval', [])) or None)

    p_a2 = a2.paragraphs[0]
    p_a2.add_run('CHUTES & SUIVI').bold = True
    p_a2.runs[0].font.size = Pt(8); p_a2.runs[0].font.color.rgb = BLEU_TITRE

    add_kv(a2, 'Chute dans les 6 mois', r.get('chute'))
    add_kv(a2, 'Hospitalisation', r.get('hosp'))
    add_kv(a2, "Perte d'autonomie", r.get('paut'))
    add_kv(a2, 'Suivi ergothérapeute', r.get('ergo'))
    add_kv(a2, 'Rapport ergo disponible', r.get('rergo'))
    add_kv(a2, 'Aidant / aide pro', r.get('aid2'))
    add_kv(a2, 'Personne de confiance', r.get('conf'))
    add_kv(a2, 'Identité conf.', f.get('conf_id'))
    if f.get('obs_aut'):
        p_obs = a2.add_paragraph()
        p_obs.paragraph_format.space_before = Pt(4)
        r_obs1 = p_obs.add_run('Observations : ')
        r_obs1.bold = True; r_obs1.font.size = Pt(8.5); r_obs1.font.color.rgb = GRIS_T
        r_obs2 = p_obs.add_run(f.get('obs_aut'))
        r_obs2.font.size = Pt(8.5)

    # ── PARTIE 3 — ÉCONOMIQUE ─────────────────────────────────────────────────
    rev_ids = [('e_sal','Salaire'),('e_pen','Pension'),('e_alloc','Allocations'),('e_autre','Autres revenus')]
    ch_ids  = [('c_loy','Loyer'),('c_pret','Prêt immobilier'),('c_cred','Crédits'),
               ('c_mut','Mutuelle'),('c_elec','Électricité'),('c_ass','Assurances'),('c_autre','Autres')]
    rev_vals = {k: float(f.get(k,0) or 0) for k,_ in rev_ids}
    ch_vals  = {k: float(f.get(k,0) or 0) for k,_ in ch_ids}
    tot_rev  = sum(rev_vals.values())
    tot_ch   = sum(ch_vals.values())

    if tot_rev or tot_ch:
        add_h1(doc, 'PARTIE 3 — SITUATION ÉCONOMIQUE (facultatif)')
        t_eco = doc.add_table(rows=1, cols=3)
        t_eco.style = 'Table Grid'
        we = [6.5, 4.0, 4.0]
        tbl_header_row(t_eco, ['Poste', 'Ressources (€/mois)', 'Charges (€/mois)'], we)
        for i, (k, label) in enumerate(rev_ids + ch_ids):
            is_r = i < len(rev_ids)
            val  = rev_vals.get(k,0) if is_r else ch_vals.get(k,0)
            if not val:
                continue
            rw = t_eco.add_row()
            bg = 'EBF3FB' if i % 2 == 0 else 'FFFFFF'
            for j, (v, w) in enumerate(zip([label, f'{val:,.0f} €' if is_r else '', f'{val:,.0f} €' if not is_r else ''], we)):
                cl = rw.cells[j]
                cl.width = Cm(w)
                cell_bg(cl, bg); cell_borders(cl); cell_pad(cl)
                p_ = cl.paragraphs[0]
                run = p_.add_run(v)
                run.font.size = Pt(8.5)
        # Total
        rw = t_eco.add_row()
        for j, (txt, w) in enumerate(zip(['TOTAL', f'{tot_rev:,.0f} €', f'{tot_ch:,.0f} €'], we)):
            cl = rw.cells[j]
            cl.width = Cm(w); cell_bg(cl, 'D6E4F0'); cell_borders(cl, '2E75B6'); cell_pad(cl)
            p_ = cl.paragraphs[0]
            p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p_.add_run(txt)
            run.bold = True; run.font.size = Pt(9); run.font.color.rgb = BLEU

        rav = tot_rev - tot_ch
        p_rav = doc.add_paragraph()
        p_rav.paragraph_format.space_before = Pt(5)
        p_rav.add_run('Reste à vivre estimé : ').bold = True
        p_rav.runs[0].font.size = Pt(10)
        rv2 = p_rav.add_run(f'{rav:,.0f} €/mois')
        rv2.bold = True; rv2.font.size = Pt(12)
        rv2.font.color.rgb = VERT if rav >= 0 else ROUGE

    # ── PARTIE 4 — ÉTAT DES LIEUX ─────────────────────────────────────────────
    doc.add_page_break()
    add_h1(doc, 'PARTIE 4 — ÉTAT DES LIEUX GLOBAL DU LOGEMENT', space_before=6)

    # Légende
    p_leg = doc.add_paragraph()
    p_leg.paragraph_format.space_after = Pt(5)
    for txt, col in [('■ Adapté  ', '1E8449'), ('■ Non adapté  ', 'C0392B'), ('■ Non renseigné', '888888')]:
        rl = p_leg.add_run(txt)
        rl.font.size = Pt(8)
        rl.font.color.rgb = RGBColor(int(col[:2],16), int(col[2:4],16), int(col[4:],16))

    # Colonnes selon présence de photos
    has_photos = any(ph_pieces.get(pn) for pn in PIECES)
    if has_photos:
        headers_edl = ['Pièce', 'Statut', 'Circulation', 'Équipements', 'Étage', 'Remarques', 'Photos']
        widths_edl  = [2.8, 1.5, 2.8, 2.8, 1.2, 2.2, 4.2]
    else:
        headers_edl = ['Pièce', 'Statut', 'Circulation', 'Équipements', 'Étage', 'Remarques']
        widths_edl  = [3.2, 1.8, 3.5, 3.5, 1.4, 4.1]

    t_edl = doc.add_table(rows=1, cols=len(headers_edl))
    t_edl.style = 'Table Grid'
    tbl_header_row(t_edl, headers_edl, widths_edl)

    for i, pn in enumerate(PIECES):
        pd     = pieces.get(pn, {})
        st     = pd.get('st', '')
        statut = 'Adapté' if st == 'ok' else 'Non adapté' if st == 'nok' else '—'
        piece_ph = ph_pieces.get(pn, [])
        bg = 'FFFFFF' if i % 2 == 0 else 'F5F8FC'

        row = t_edl.add_row()
        vals = [pn, statut, pd.get('circ',''), pd.get('equip',''), pd.get('etg',''), pd.get('rem','')]

        for j, (val, w) in enumerate(zip(vals, widths_edl)):
            cell = row.cells[j]
            cell.width = Cm(w)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_pad(cell, 60, 80, 60, 80)

            if j == 1:  # statut
                if st == 'ok':
                    cell_bg(cell, 'EAF3DE'); cell_borders(cell, '7DCEA0')
                elif st == 'nok':
                    cell_bg(cell, 'FDEDEC'); cell_borders(cell, 'F1948A')
                else:
                    cell_bg(cell, bg); cell_borders(cell)
            else:
                cell_bg(cell, bg); cell_borders(cell)

            p_ = cell.paragraphs[0]
            run = p_.add_run(str(val) if val else '')
            run.font.size = Pt(8.5)
            if j == 0:
                run.bold = True; run.font.color.rgb = BLEU
            if j == 1:
                run.bold = True
                run.font.color.rgb = VERT if st=='ok' else ROUGE if st=='nok' else GRIS_T
                p_.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Colonne photos
        if has_photos:
            ph_cell = row.cells[6]
            ph_cell.width = Cm(widths_edl[6])
            ph_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_bg(ph_cell, bg); cell_borders(ph_cell)
            cell_pad(ph_cell, 40, 40, 40, 40)
            if piece_ph:
                add_photos_in_cell(ph_cell, piece_ph, max_w_cm=3.8, max_count=4)
            else:
                p_ = ph_cell.paragraphs[0]
                r_ = p_.add_run('—')
                r_.font.size = Pt(8); r_.font.color.rgb = GRIS_T

    # ── PARTIE 5 — TRAVAUX ────────────────────────────────────────────────────
    doc.add_page_break()
    add_h1(doc, 'PARTIE 5 — PROJET DE TRAVAUX', space_before=6)

    if has_photos:
        headers_trav = ['Pièce', 'Priorité', 'Solutions proposées AMO', 'Solutions retenues', 'Photos avant']
        widths_trav  = [2.8, 1.8, 4.8, 4.8, 3.3]
    else:
        headers_trav = ['Pièce', 'Priorité', 'Solutions proposées AMO', 'Solutions retenues']
        widths_trav  = [3.2, 2.0, 6.0, 6.3]

    t_trav = doc.add_table(rows=1, cols=len(headers_trav))
    t_trav.style = 'Table Grid'
    tbl_header_row(t_trav, headers_trav, widths_trav, bg='2E75B6')

    PCOL = {'Haute': 'FDEDEC', 'Moyenne': 'FEF9E7', 'Basse': 'EAF3DE'}
    PTXT = {'Haute': ROUGE,    'Moyenne': ORANGE,   'Basse': VERT}

    for i, pn in enumerate(PIECES):
        td       = trav.get(pn, {})
        prio     = td.get('prio', '')
        sol      = td.get('sol', '')
        ret      = td.get('ret', '')
        piece_ph = ph_pieces.get(pn, [])
        bg = 'FFFFFF' if i % 2 == 0 else 'F5F8FC'

        row = t_trav.add_row()
        vals = [pn, prio, sol, ret]

        for j, (val, w) in enumerate(zip(vals, widths_trav)):
            cell = row.cells[j]
            cell.width = Cm(w)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_pad(cell, 60, 80, 60, 80)

            if j == 1 and prio in PCOL:
                cell_bg(cell, PCOL[prio]); cell_borders(cell)
            else:
                cell_bg(cell, bg); cell_borders(cell)

            p_ = cell.paragraphs[0]
            run = p_.add_run(str(val) if val else '')
            run.font.size = Pt(8.5)
            if j == 0:
                run.bold = True; run.font.color.rgb = BLEU
            if j == 1 and prio in PTXT:
                run.bold = True; run.font.color.rgb = PTXT[prio]
                p_.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if has_photos:
            ph_cell = row.cells[4]
            ph_cell.width = Cm(widths_trav[4])
            ph_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_bg(ph_cell, bg); cell_borders(ph_cell)
            cell_pad(ph_cell, 40, 40, 40, 40)
            if piece_ph:
                add_photos_in_cell(ph_cell, piece_ph, max_w_cm=3.0, max_count=2)
            else:
                p_ = ph_cell.paragraphs[0]
                r_ = p_.add_run('—')
                r_.font.size = Pt(8); r_.font.color.rgb = GRIS_T

    # Photos travaux générales
    if ph_trav_gen:
        add_h2(doc, 'Photos complémentaires — zones à traiter')
        p_ph = doc.add_paragraph()
        p_ph.paragraph_format.space_after = Pt(6)
        for ph in ph_trav_gen[:6]:
            raw, _ = b64_to_bytes(ph.get('data',''))
            if raw:
                try:
                    run = p_ph.add_run()
                    run.add_picture(io.BytesIO(raw), width=Cm(4.2))
                    p_ph.add_run('  ')
                except Exception:
                    pass

    # Préconisations
    if f.get('prec'):
        add_h2(doc, 'Détail des recommandations AMO')
        p_ = doc.add_paragraph(f.get('prec'))
        p_.paragraph_format.space_after = Pt(4)
        if p_.runs: p_.runs[0].font.size = Pt(9)

    if f.get('crp'):
        add_h2(doc, 'Création de pièce nouvelle')
        p_ = doc.add_paragraph(f.get('crp'))
        if p_.runs: p_.runs[0].font.size = Pt(9)

    if f.get('rem_men'):
        add_h2(doc, 'Remarques du ménage')
        p_ = doc.add_paragraph(f.get('rem_men'))
        if p_.runs: p_.runs[0].font.size = Pt(9)

    # ── PIED ──────────────────────────────────────────────────────────────────
    add_sep(doc)
    pied = doc.add_paragraph()
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pied.paragraph_format.space_before = Pt(3)
    rp = pied.add_run(
        'EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911 · SIRET 808 408 165 000 13\n'
        'ei_dmaai@outlook.fr · 06 34 40 79 32 · Orléans'
    )
    rp.font.size = Pt(7.5); rp.font.color.rgb = GRIS_T

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── ROUTES ───────────────────────────────────────────────────────────────────
def generer_plan(raw_data: dict) -> bytes:
    data = normaliser(raw_data)
    f    = data.get('flds', {})
    r    = data.get('rad', {})
    ch   = data.get('chk', {})
    devis_list = raw_data.get('devis', [])

    # ── Calcul profil ménage ──────────────────────────────────────────────────
    PLAFOND_HT  = REGLES['mpa']['plafond_travaux_ht']
    FORFAIT_AMO = REGLES['mpa']['amo']['forfait_complet_ttc']
    TAUX_MPA    = {k: round(v*100) for k,v in REGLES['mpa']['taux'].items()}

    nb_occ  = int(f.get('l_occ', 1) or 1)
    adresse = (f.get('l_adr', '') or '').lower()
    codes_idf = ['75','77','78','91','92','93','94','95']
    is_idf = any(c in adresse for c in codes_idf)
    zone_label = 'Île-de-France' if is_idf else 'Hors Île-de-France'

    # RFR annuel — champ direct saisi dans l'appli
    rfr_direct = float(f.get('e_rfr', 0) or 0)
    if rfr_direct > 0:
        rev_an = rfr_direct
    else:
        # Fallback : revenus mensuels × 12 (ancien format)
        rev_ids = ['e_sal','e_pen','e_alloc','e_autre']
        rev_mens = sum(float(f.get(k, 0) or 0) for k in rev_ids)
        rev_an = rev_mens * 12

    p_tmo = plafond_ressources('tres_modeste', nb_occ, is_idf)
    p_mo  = plafond_ressources('modeste', nb_occ, is_idf)

    if rev_an > 0 and rev_an <= p_tmo:
        profil = 'Très modeste (TMO)'; taux = TAUX_MPA['tres_modeste']
    elif rev_an > 0 and rev_an <= p_mo:
        profil = 'Modeste (MO)'; taux = TAUX_MPA['modeste']
    else:
        profil = "Non déterminé — vérifier les avis d'imposition"; taux = 0

    statut   = r.get('stat', '—')
    nom_dem  = f'{f.get("d_nom","")}'.strip() + ' ' + f'{f.get("d_prn","")}'.strip()
    adresse_log = f.get('l_adr', '—')

    date_str = '—'
    if f.get('v_date'):
        try: date_str = datetime.strptime(f['v_date'], '%Y-%m-%d').strftime('%d/%m/%Y')
        except: date_str = f['v_date']

    # ── Devis ────────────────────────────────────────────────────────────────
    devis_retenus = [d for d in devis_list if d.get('retenu') and (parse_montant(d.get('ht', 0)) > 0)]
    total_ht_trav  = sum(parse_montant(d.get('ht', 0)) for d in devis_retenus)
    total_ttc_trav = sum(parse_montant(d.get('ttc', 0)) for d in devis_retenus)

    # Honoraires AMO
    hono_ht  = float(f.get('hono_ht', 0) or 0)
    hono_tva = float(f.get('hono_tva', 0) or 0)
    hono_ttc = hono_ht * (1 + hono_tva / 100)
    hono_mpa = min(float(f.get('hono_mpa', FORFAIT_AMO) or FORFAIT_AMO), FORFAIT_AMO)

    total_ht  = total_ht_trav + hono_ht
    total_ttc = total_ttc_trav + hono_ttc

    ht_subv   = min(total_ht_trav, PLAFOND_HT)
    aide_trav = round(ht_subv * taux / 100) if taux > 0 else 0
    aide_amo  = hono_mpa if taux > 0 else 0
    aide_total = aide_trav + aide_amo
    reste_ch   = max(0, total_ttc - aide_total)

    # Autres aides complémentaires
    autres_aides = raw_data.get('autres_aides', []) or []
    total_autres = sum(float(a.get('montant', 0) or 0) for a in autres_aides)

    # ── Document Word ─────────────────────────────────────────────────────────
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    # En-tête
    t_hdr = doc.add_table(rows=1, cols=3)
    t_hdr.style = 'Table Grid'
    from docx.enum.table import WD_TABLE_ALIGNMENT
    t_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT

    c_logo = t_hdr.rows[0].cells[0]
    c_logo.width = Cm(3.8)
    cell_bg(c_logo, 'FFFFFF'); cell_borders(c_logo, 'FFFFFF')
    cell_pad(c_logo, 60, 60, 60, 60)
    try:
        logo_bytes = base64.b64decode(LOGO_DMAAI_B64)
        p_logo = c_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Cm(3.4))
    except: pass

    c_mid = t_hdr.rows[0].cells[1]
    c_mid.width = Cm(10)
    cell_bg(c_mid, '1F4E79'); cell_borders(c_mid, '1F4E79')
    cell_pad(c_mid, 100, 120, 80, 140)
    rt = c_mid.paragraphs[0].add_run("PLAN DE FINANCEMENT — MaPrimeAdapt'")
    rt.bold = True; rt.font.size = Pt(13); rt.font.color.rgb = BLANC
    p_sub = c_mid.add_paragraph()
    rs = p_sub.add_run('EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911')
    rs.font.size = Pt(7.5); rs.font.color.rgb = RGBColor(0xB0, 0xC8, 0xE8)
    p_hab = c_mid.add_paragraph()
    rh = p_hab.add_run(f"Habilitation ANAH N° {HABILIT_REF} · {HABILIT_DATE}")
    rh.font.size = Pt(7.5); rh.font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)

    c_mpa = t_hdr.rows[0].cells[2]
    c_mpa.width = Cm(3.7)
    cell_bg(c_mpa, 'FFFFFF'); cell_borders(c_mpa, 'FFFFFF')
    cell_pad(c_mpa, 60, 60, 60, 60)
    try:
        mpa_bytes = base64.b64decode(LOGO_MPA_B64)
        pm = c_mpa.paragraphs[0]
        pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pm.add_run().add_picture(io.BytesIO(mpa_bytes), width=Cm(3.2))
    except: pass

    doc.add_paragraph()

    # Bloc identification
    t_id = doc.add_table(rows=1, cols=2)
    t_id.style = 'Table Grid'
    t_id.alignment = WD_TABLE_ALIGNMENT.LEFT
    ci1 = t_id.rows[0].cells[0]; ci2 = t_id.rows[0].cells[1]
    ci1.width = Cm(8.5); ci2.width = Cm(9)
    for c in [ci1, ci2]:
        cell_bg(c, 'EBF3FB'); cell_borders(c, 'D6E4F0'); cell_pad(c, 80,120,80,120)

    def kv(cell, label, val):
        if not val or str(val).strip() in ('','—'): return
        p_ = cell.add_paragraph()
        p_.paragraph_format.space_before = Pt(1)
        r1 = p_.add_run(f'{label} : ')
        r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = GRIS_T
        r2 = p_.add_run(str(val).strip())
        r2.font.size = Pt(9)

    h1 = ci1.paragraphs[0]; h1.add_run('DEMANDEUR').bold = True
    h1.runs[0].font.size = Pt(8); h1.runs[0].font.color.rgb = BLEU_TITRE
    kv(ci1, 'Identité', nom_dem.strip())
    kv(ci1, 'Adresse du logement', adresse_log)
    kv(ci1, 'Statut', statut)
    kv(ci1, 'Nb personnes foyer', str(nb_occ))
    kv(ci1, 'Zone', zone_label)

    h2 = ci2.paragraphs[0]; h2.add_run("PROFIL MaPrimeAdapt'").bold = True
    h2.runs[0].font.size = Pt(8); h2.runs[0].font.color.rgb = BLEU_TITRE
    kv(ci2, 'Profil ménage', profil)
    kv(ci2, 'Taux de subvention', f'{taux} %' if taux else '—')
    kv(ci2, 'Plafond travaux HT', f'{eur_txt(PLAFOND_HT)} HT')
    kv(ci2, "Forfait AMO pris en charge", f"{FORFAIT_AMO} €")
    kv(ci2, 'Date du diagnostic', date_str)
    if f.get('ref_dossier'):
        kv(ci2, 'Référence dossier', f.get('ref_dossier'))

    # Avertissement si revenus non renseignés
    if rev_an == 0:
        p_warn = doc.add_paragraph()
        p_warn.paragraph_format.space_before = Pt(6)
        rw = p_warn.add_run("⚠️  Profil calculé sur la base des revenus mensuels saisis. Pour le dépôt du dossier, l'ANAH prend en compte les revenus fiscaux de référence (avis d'imposition N-2).")
        rw.font.size = Pt(8); rw.font.color.rgb = RGBColor(0x85, 0x64, 0x04)
        rw.italic = True

    # Titre tableau devis
    add_h1(doc, "DEVIS RETENUS")

    if not devis_retenus:
        p_nd = doc.add_paragraph('Aucun devis retenu.')
        p_nd.runs[0].font.size = Pt(10); p_nd.runs[0].font.color.rgb = GRIS_T
    else:
        t_dev = doc.add_table(rows=1, cols=6)
        t_dev.style = 'Table Grid'
        t_dev.alignment = WD_TABLE_ALIGNMENT.LEFT
        wd = [1.2, 2.5, 3.5, 2.5, 2.2, 2.2]
        tbl_header_row(t_dev, ['N°', 'SIRET', 'Entreprise', 'Type de travaux', 'Montant HT', 'Montant TTC'], wd)
        for i, d in enumerate(devis_retenus):
            ht  = parse_montant(d.get('ht', 0))
            ttc = parse_montant(d.get('ttc', 0))
            row = t_dev.add_row()
            bg  = 'FFFFFF' if i%2==0 else 'F5F8FC'
            vals = [str(i+1), d.get('siret',''), d.get('raison',''), d.get('type',''), 
                    fmt_eur(ht), fmt_eur(ttc)]
            for j,(val,w) in enumerate(zip(vals, wd)):
                cl = row.cells[j]; cl.width = Cm(w)
                cl.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                cell_bg(cl, bg); cell_borders(cl); cell_pad(cl, 60,80,60,80)
                p_ = cl.paragraphs[0]
                run = p_.add_run(str(val))
                run.font.size = Pt(8.5)
                if j == 0: run.bold = True; p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if j in (4,5): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Ligne honoraires AMO si renseigné
        total_ht_fmt = fmt_eur(total_ht)
        if hono_ht > 0:
            rh = t_dev.add_row()
            bg_h = 'EBF3FB'
            hono_vals = ['', '808 408 165 000 13', 'EI David MARIE-ANNE', 'Mission AMO',
                        fmt_eur(hono_ht), fmt_eur(hono_ttc)]
            for j2,(val,w) in enumerate(zip(hono_vals, wd)):
                cl = rh.cells[j2]; cl.width = Cm(w)
                cell_bg(cl, bg_h); cell_borders(cl); cell_pad(cl, 60,80,60,80)
                p_ = cl.paragraphs[0]
                run = p_.add_run(str(val)); run.font.size = Pt(8.5); run.italic = True
                if j2 in (4,5): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Ligne total devis
        rt_ = t_dev.add_row()
        for j,(txt,w) in enumerate(zip(['','','TOTAL HT', '', fmt_eur(total_ht), fmt_eur(total_ttc)], wd)):
            cl = rt_.cells[j]; cl.width = Cm(w)
            cell_bg(cl, 'D6E4F0'); cell_borders(cl, '2E75B6'); cell_pad(cl, 60,80,60,80)
            p_ = cl.paragraphs[0]
            run = p_.add_run(txt)
            run.bold = True; run.font.size = Pt(9); run.font.color.rgb = BLEU
            if j in (2,4,5): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j!=2 else WD_ALIGN_PARAGRAPH.RIGHT

        # Note dépassement plafond
        if total_ht > PLAFOND_HT:
            p_dep = doc.add_paragraph()
            p_dep.paragraph_format.space_before = Pt(4)
            rd = p_dep.add_run(f"⚠️  Le montant HT total ({total_ht_fmt})  d\u00e9passe le plafond subventionnable de {eur_txt(PLAFOND_HT)} HT. Seuls {eur_txt(PLAFOND_HT)} HT sont pris en compte pour le calcul de l'aide.")
            rd.font.size = Pt(8.5); rd.font.color.rgb = ORANGE; rd.italic = True

    # Plan de financement
    add_h1(doc, "PLAN DE FINANCEMENT MaPrimeAdapt'")

    t_plan = doc.add_table(rows=1, cols=4)
    t_plan.style = 'Table Grid'
    t_plan.alignment = WD_TABLE_ALIGNMENT.LEFT
    wp = [5.5, 3.0, 3.0, 3.5]
    tbl_header_row(t_plan, ['Poste', 'Montant', "Taux MPA'", "Aide MPA'"], wp, bg='2E75B6')

    rows_plan = [
        (f"Travaux d'adaptation (plafond {eur_txt(PLAFOND_HT)} HT)",
         fmt_eur(ht_subv).replace(' €',' € HT'),
         f'{taux} %' if taux else '—',
         f'{aide_trav:,.0f} €'),
        (f"Forfait accompagnement AMO{' (honoraires : ' + str(round(hono_ht,2)) + ' € HT)' if hono_ht > 0 else ''}",
         fmt_eur(hono_ht).replace(' €',' € HT') if hono_ht > 0 else 'Forfait',
         'Forfait fixe',
         f'{aide_amo:,.0f} €'),
    ]
    for i,(label,mont,tx,aide) in enumerate(rows_plan):
        row = t_plan.add_row()
        for j,(val,w) in enumerate(zip([label,mont,tx,aide], wp)):
            cl = row.cells[j]; cl.width = Cm(w)
            bg = 'EBF3FB' if i%2==0 else 'FFFFFF'
            cell_bg(cl, bg); cell_borders(cl); cell_pad(cl, 80,100,80,100)
            p_ = cl.paragraphs[0]
            run = p_.add_run(val)
            run.font.size = Pt(9)
            if j > 0: p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Ligne TOTAL aide
    rt_plan = t_plan.add_row()
    for j,(val,w) in enumerate(zip(["AIDE TOTALE MaPrimeAdapt'", fmt_eur(total_ttc).replace(' €',' € TTC'), '', f'{aide_total:,.0f} €'], wp)):
        cl = rt_plan.cells[j]; cl.width = Cm(w)
        cell_bg(cl, '1F4E79'); cell_borders(cl, '1F4E79'); cell_pad(cl, 80,100,80,100)
        p_ = cl.paragraphs[0]
        run = p_.add_run(val)
        run.bold = True; run.font.size = Pt(10); run.font.color.rgb = BLANC
        if j > 0: p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Autres aides complémentaires
    if autres_aides and total_autres > 0:
        add_h1(doc, "AUTRES AIDES COMPLÉMENTAIRES")
        t_autres = doc.add_table(rows=1, cols=3)
        t_autres.style = 'Table Grid'
        t_autres.alignment = WD_TABLE_ALIGNMENT.LEFT
        wa = [9.0, 3.5, 3.0]
        tbl_header_row(t_autres, ['Organisme / libellé', 'Montant', 'Type'], wa)
        for i, a in enumerate(autres_aides):
            if not float(a.get('montant', 0) or 0): continue
            row = t_autres.add_row()
            bg = 'EBF3FB' if i%2==0 else 'FFFFFF'
            for j2,(val,w) in enumerate(zip([a.get('label','—'), fmt_eur(float(a.get('montant',0))), a.get('source','—')], wa)):
                cl = row.cells[j2]; cl.width = Cm(w)
                cell_bg(cl, bg); cell_borders(cl); cell_pad(cl)
                p_ = cl.paragraphs[0]
                run = p_.add_run(str(val)); run.font.size = Pt(9)
                if j2 == 1: p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        reste_final = max(0, reste_ch - total_autres)
        p_af = doc.add_paragraph()
        p_af.paragraph_format.space_before = Pt(4)
        r1 = p_af.add_run(f'Total autres aides : {total_autres:,.0f} €  →  Reste à charge final : ')
        r1.font.size = Pt(9); r1.font.color.rgb = GRIS_T
        r2 = p_af.add_run(fmt_eur(reste_final))
        r2.bold = True; r2.font.size = Pt(12)
        r2.font.color.rgb = VERT if reste_final < reste_ch else ROUGE

    # Reste à charge — grand bloc visible
    p_rc = doc.add_paragraph()
    p_rc.paragraph_format.space_before = Pt(10)
    p_rc.paragraph_format.space_after  = Pt(4)
    r_lbl = p_rc.add_run('Reste à charge estimé : ')
    r_lbl.bold = True; r_lbl.font.size = Pt(12); r_lbl.font.color.rgb = GRIS_T
    r_val = p_rc.add_run(fmt_eur(reste_ch))
    r_val.bold = True; r_val.font.size = Pt(16)
    r_val.font.color.rgb = VERT if reste_ch < total_ttc * 0.5 else ROUGE

    # Avance possible
    avance = round(aide_total * 0.30)
    p_av = doc.add_paragraph()
    p_av.paragraph_format.space_before = Pt(4)
    ra = p_av.add_run(f"ℹ️  Avance possible sur l'aide MPA' : jusqu'à 30 % soit environ {avance:,.0f} €")
    ra.font.size = Pt(9); ra.italic = True; ra.font.color.rgb = BLEU_TITRE

    # Notes légales
    add_sep(doc, '2E75B6', 4)
    p_notes = doc.add_paragraph()
    p_notes.paragraph_format.space_before = Pt(6)
    rn = p_notes.add_run(
        "Ce plan de financement est établi à titre indicatif sur la base des informations communiquées lors du diagnostic. "
        "Le montant définitif de l'aide MaPrimeAdapt' sera déterminé par l'ANAH après instruction du dossier complet "
        "(avis d'imposition N-2, devis définitifs, accord de l'ANAH avant travaux). "
        "Les travaux ne doivent pas débuter avant l'accord de financement de l'ANAH."
    )
    rn.font.size = Pt(8); rn.italic = True; rn.font.color.rgb = GRIS_T

    # Pied de page
    add_sep(doc)
    pied = doc.add_paragraph()
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp = pied.add_run(
        "EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911 · SIRET 808 408 165 000 13"
        "ei_dmaai@outlook.fr · 06 34 40 79 32 · Orléans"
    )
    rp.font.size = Pt(7.5); rp.font.color.rgb = GRIS_T

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── ROUTES ───────────────────────────────────────────────────────────────────

@app.route('/reception')
def serve_reception():
    import os
    folder = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(folder, 'agent_reception_mpa.html')
    if os.path.exists(p):
        with open(p, 'r', encoding='utf-8') as fh:
            content = fh.read()
        from flask import Response
        return Response(content, mimetype='text/html; charset=utf-8')
    return '<h2>agent_reception_mpa.html introuvable</h2>', 404


@app.route('/attestation', methods=['POST'])
def route_attestation():
    try:
        data = request.get_json(force=True)
        docx_bytes = generer_attestation(data)
        b = data.get('beneficiaire', {})
        nom = (b.get('nom', 'X') or 'X').replace(' ', '_')
        date = b.get('recep', datetime.today().strftime('%Y-%m-%d'))
        filename = f'Attestation_Fin_Mission_MPA_{nom}_{date}.docx'
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/plan_definitif', methods=['POST'])
def route_plan_definitif():
    try:
        data = request.get_json(force=True)
        docx_bytes = generer_plan_definitif(data)
        b = data.get('beneficiaire', {})
        nom = (b.get('nom', 'X') or 'X').replace(' ', '_')
        date = data.get('date_gen', datetime.today().strftime('%Y-%m-%d'))
        filename = f'Plan_Financement_Definitif_MPA_{nom}_{date}.docx'
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


# ─── GÉNÉRATION ATTESTATION DE FIN DE MISSION ────────────────────────────────
def generer_attestation(data: dict) -> bytes:
    b = data.get('beneficiaire', {})
    postes = data.get('postes', [])
    factures = data.get('factures', [])
    photos_generales = data.get('photos_generales', []) or []

    nom       = b.get('nom', '—')
    adresse   = b.get('adresse', '—')
    dossier   = b.get('dossier', '—')
    entreprises = b.get('entreprises', '—')
    permis    = b.get('permis', '')
    reserves  = b.get('reserves', 'Aucune réserve')

    recep_raw = b.get('recep', '')
    try:
        recep_str = datetime.strptime(recep_raw, '%Y-%m-%d').strftime('%d/%m/%Y') if recep_raw else '—'
    except:
        recep_str = recep_raw or '—'

    deb_raw = b.get('deb', '')
    try:
        deb_str = datetime.strptime(deb_raw, '%Y-%m-%d').strftime('%d/%m/%Y') if deb_raw else '—'
    except:
        deb_str = deb_raw or '—'

    postes_realises = [p['nom'] for p in postes if p.get('recu')]
    postes_tous     = [p['nom'] for p in postes]
    nb_recu = len(postes_realises)
    nb_total = len(postes_tous)

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin   = Cm(3)
        sec.right_margin  = Cm(2.5)
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    # En-tête avec logos
    t_hdr = doc.add_table(rows=1, cols=3)
    t_hdr.style = 'Table Grid'
    t_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT
    c_logo = t_hdr.rows[0].cells[0]; c_logo.width = Cm(3.5)
    cell_bg(c_logo, 'FFFFFF'); cell_borders(c_logo, 'FFFFFF'); cell_pad(c_logo, 60,60,60,60)
    try:
        logo_bytes = base64.b64decode(LOGO_DMAAI_B64)
        p_logo = c_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Cm(3.2))
    except: pass

    c_mid = t_hdr.rows[0].cells[1]; c_mid.width = Cm(10.5)
    cell_bg(c_mid, '1F4E79'); cell_borders(c_mid, '1F4E79'); cell_pad(c_mid, 100,120,80,140)
    rt = c_mid.paragraphs[0].add_run("ATTESTATION DE FIN DE MISSION AMO")
    rt.bold = True; rt.font.size = Pt(13); rt.font.color.rgb = BLANC
    p2 = c_mid.add_paragraph()
    p2.add_run("MaPrimeAdapt' — Assistance à Maîtrise d'Ouvrage").font.size = Pt(9)
    p2.runs[0].font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)
    p3 = c_mid.add_paragraph()
    rs = p3.add_run('EI David MARIE-ANNE · Architecte DPLG N°049576 · SIRET 808 408 165 000 13')
    rs.font.size = Pt(7.5); rs.font.color.rgb = RGBColor(0xB0, 0xC8, 0xE8)
    p4 = c_mid.add_paragraph()
    rh = p4.add_run(f'Habilitation ANAH N° {HABILIT_REF} · {HABILIT_DATE} · {HABILIT_DEPT}')
    rh.font.size = Pt(7.5); rh.bold = True; rh.font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)

    c_mpa = t_hdr.rows[0].cells[2]; c_mpa.width = Cm(3.5)
    cell_bg(c_mpa, 'FFFFFF'); cell_borders(c_mpa, 'FFFFFF'); cell_pad(c_mpa, 60,60,60,60)
    try:
        mpa_bytes = base64.b64decode(LOGO_MPA_B64)
        pm = c_mpa.paragraphs[0]; pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pm.add_run().add_picture(io.BytesIO(mpa_bytes), width=Cm(3.0))
    except: pass

    doc.add_paragraph()

    # Corps
    add_h1(doc, 'JE SOUSSIGNÉ(E)')
    t_amo = doc.add_table(rows=1, cols=2)
    t_amo.style = 'Table Grid'
    ca1 = t_amo.rows[0].cells[0]; ca2 = t_amo.rows[0].cells[1]
    ca1.width = Cm(8.5); ca2.width = Cm(9)
    for c in [ca1, ca2]:
        cell_bg(c, 'EBF3FB'); cell_borders(c, 'D6E4F0'); cell_pad(c, 80,120,80,120)

    def kv2(cell, label, val):
        if not val or str(val).strip() in ('','—','None'): return
        p_ = cell.add_paragraph()
        p_.paragraph_format.space_before = Pt(1)
        r1 = p_.add_run(f'{label} : '); r1.bold = True
        r1.font.size = Pt(9); r1.font.color.rgb = GRIS_T
        r2 = p_.add_run(str(val).strip()); r2.font.size = Pt(9)

    ca1.paragraphs[0].add_run('AMO').bold = True
    ca1.paragraphs[0].runs[0].font.size = Pt(8); ca1.paragraphs[0].runs[0].font.color.rgb = BLEU_TITRE
    kv2(ca1, 'Nom', 'David MARIE-ANNE')
    kv2(ca1, 'Qualité', 'Architecte DPLG / Ingénieur TCE')
    kv2(ca1, 'Structure', 'EI DMAAI (EI David MARIE-ANNE)')
    kv2(ca1, 'SIRET', '808 408 165 000 13')
    kv2(ca1, 'Contact', 'ei_dmaai@outlook.fr · 06 34 40 79 32')
    kv2(ca1, 'Habilitation ANAH', f'{HABILIT_REF} · {HABILIT_DATE}')

    ca2.paragraphs[0].add_run('BÉNÉFICIAIRE').bold = True
    ca2.paragraphs[0].runs[0].font.size = Pt(8); ca2.paragraphs[0].runs[0].font.color.rgb = BLEU_TITRE
    kv2(ca2, 'Nom complet', nom)
    kv2(ca2, 'Adresse des travaux', adresse)
    kv2(ca2, 'N° dossier ANAH', dossier)
    kv2(ca2, 'Entreprise(s)', entreprises)
    kv2(ca2, 'Date début mission AMO', deb_str)
    kv2(ca2, 'Date réception travaux', recep_str)
    if permis:
        kv2(ca2, 'Permis / déclaration', permis)

    add_h1(doc, 'CERTIFIE')
    p_cert = doc.add_paragraph()
    p_cert.paragraph_format.space_before = Pt(4)
    p_cert.paragraph_format.space_after  = Pt(6)
    r_cert = p_cert.add_run(
        f"Avoir accompagné M./Mme {nom} dans le cadre de la mission d'Assistance à Maîtrise d'Ouvrage "
        f"pour le dispositif MaPrimeAdapt', conformément à l'habilitation ANAH N° {HABILIT_REF}, "
        f"depuis le {deb_str} jusqu'à la réception des travaux le {recep_str}."
    )
    r_cert.font.size = Pt(10)

    if postes_tous:
        add_h1(doc, 'POSTES DE TRAVAUX RÉCEPTIONNÉS')
        t_pos = doc.add_table(rows=1, cols=3)
        t_pos.style = 'Table Grid'
        tbl_header_row(t_pos, ['Poste de travaux', 'Réceptionné', 'Observations'], [7, 2.5, 8])
        for p_ in postes:
            row = t_pos.add_row()
            recu = '✓' if p_.get('recu') else '—'
            bg = 'EAF3DE' if p_.get('recu') else 'FFFFFF'
            for j,(val,w) in enumerate(zip([p_.get('nom',''), recu, p_.get('obs','')], [7,2.5,8])):
                cl = row.cells[j]; cl.width = Cm(w)
                cell_bg(cl, bg); cell_borders(cl); cell_pad(cl, 60,80,60,80)
                p__ = cl.paragraphs[0]
                run = p__.add_run(str(val)); run.font.size = Pt(9)
                if j == 1:
                    p__.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run.bold = True
                    run.font.color.rgb = VERT if p_.get('recu') else GRIS_T

        p_bilan = doc.add_paragraph()
        p_bilan.paragraph_format.space_before = Pt(6)
        rb = p_bilan.add_run(f'Bilan : {nb_recu}/{nb_total} poste(s) réceptionné(s) sans réserve.')
        rb.bold = True; rb.font.size = Pt(10)
        rb.font.color.rgb = VERT if nb_recu == nb_total else ORANGE

    # Photos par poste
    postes_avec_photos = [p for p in postes if p.get('photos')]
    if postes_avec_photos:
        add_h1(doc, 'PHOTOS APRÈS TRAVAUX — PAR POSTE')
        for p_ in postes_avec_photos:
            ph_list = p_.get('photos', [])
            if not ph_list:
                continue
            # Titre du poste
            p_tit = doc.add_paragraph()
            p_tit.paragraph_format.space_before = Pt(6)
            p_tit.paragraph_format.space_after  = Pt(2)
            rt = p_tit.add_run(p_.get('nom', ''))
            rt.bold = True; rt.font.size = Pt(9); rt.font.color.rgb = BLEU_TITRE
            # Tableau 4 photos max par ligne (2 colonnes × 2 lignes)
            n = min(len(ph_list), 4)
            cols = 2
            rows_needed = (n + cols - 1) // cols
            t_ph = doc.add_table(rows=rows_needed, cols=cols)
            t_ph.style = 'Table Grid'
            idx = 0
            for ri in range(rows_needed):
                for ci in range(cols):
                    if idx >= n:
                        break
                    cell = t_ph.rows[ri].cells[ci]
                    cell.width = Cm(8.5)
                    cell_bg(cell, 'F5F9FC'); cell_borders(cell); cell_pad(cell, 60, 60, 60, 60)
                    raw, _ = b64_to_bytes(ph_list[idx].get('data', ''))
                    if raw:
                        try:
                            p_img = cell.paragraphs[0]
                            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_img.add_run().add_picture(io.BytesIO(raw), width=Cm(7.5))
                        except Exception:
                            cell.paragraphs[0].add_run(ph_list[idx].get('name', '(photo)'))
                    # Légende nom fichier
                    p_leg = cell.add_paragraph()
                    rl = p_leg.add_run(ph_list[idx].get('name', ''))
                    rl.font.size = Pt(7); rl.font.color.rgb = GRIS_T
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    idx += 1

    # Photos générales
    if photos_generales:
        add_h1(doc, 'PHOTOS GÉNÉRALES DE CHANTIER')
        n = min(len(photos_generales), 6)
        cols = 2
        rows_needed = (n + cols - 1) // cols
        t_gen = doc.add_table(rows=rows_needed, cols=cols)
        t_gen.style = 'Table Grid'
        idx = 0
        for ri in range(rows_needed):
            for ci in range(cols):
                if idx >= n:
                    break
                cell = t_gen.rows[ri].cells[ci]
                cell.width = Cm(8.5)
                cell_bg(cell, 'F5F9FC'); cell_borders(cell); cell_pad(cell, 60, 60, 60, 60)
                raw, _ = b64_to_bytes(photos_generales[idx].get('data', ''))
                if raw:
                    try:
                        p_img = cell.paragraphs[0]
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.add_run().add_picture(io.BytesIO(raw), width=Cm(7.5))
                    except Exception:
                        cell.paragraphs[0].add_run(photos_generales[idx].get('name', '(photo)'))
                p_leg = cell.add_paragraph()
                rl = p_leg.add_run(photos_generales[idx].get('name', ''))
                rl.font.size = Pt(7); rl.font.color.rgb = GRIS_T
                p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                idx += 1

    add_h1(doc, 'RÉSERVES')
    p_res = doc.add_paragraph()
    p_res.paragraph_format.space_before = Pt(2)
    p_res.paragraph_format.left_indent  = Cm(1)
    r_res = p_res.add_run(reserves or 'Aucune réserve.')
    r_res.font.size = Pt(10)

    add_h1(doc, 'SIGNATURE')
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(8)
    r_sig = p_sig.add_run(
        f"Fait à Orléans, le {datetime.today().strftime('%d/%m/%Y')}\n\n"
        "David MARIE-ANNE\nArchitecte DPLG — EI DMAAI\n"
        "SIRET 808 408 165 000 13\n\n\n"
        "_______________________________\n(Signature)"
    )
    r_sig.font.size = Pt(10)

    add_sep(doc)
    pied = doc.add_paragraph()
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp = pied.add_run(
        'EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911 · SIRET 808 408 165 000 13\n'
        'ei_dmaai@outlook.fr · 06 34 40 79 32 · Orléans'
    )
    rp.font.size = Pt(7.5); rp.font.color.rgb = GRIS_T

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── GÉNÉRATION PLAN DE FINANCEMENT DÉFINITIF ────────────────────────────────
def generer_plan_definitif(data: dict) -> bytes:
    b = data.get('beneficiaire', {})
    factures = data.get('factures', [])
    autres_aides = data.get('autres_aides', []) or []
    postes = data.get('postes', []) or []
    photos_generales = data.get('photos_generales', []) or []

    nom       = b.get('nom', '—')
    adresse   = b.get('adresse', '—')
    dossier   = b.get('dossier', '—')

    hono_ht   = float(data.get('hono_ht', 0) or 0)
    hono_tva  = float(data.get('hono_tva', 0) or 0)
    hono_ttc  = hono_ht * (1 + hono_tva / 100)
    aide_mpa  = float(data.get('aide_mpa', 0) or 0)
    aide_amo  = float(data.get('aide_amo', 0) or 0)

    total_trav_ht  = sum(parse_montant(f.get('ht', 0)) for f in factures)
    total_trav_ttc = sum(parse_montant(f.get('ht', 0)) * (1 + parse_montant(f.get('tva', 0))/100) for f in factures)
    total_autres   = sum(parse_montant(a.get('montant', 0)) for a in autres_aides)
    total_charges_ht  = total_trav_ht + hono_ht
    total_charges_ttc = total_trav_ttc + hono_ttc
    total_aides = aide_mpa + aide_amo + total_autres
    reste_ch = max(0, total_charges_ttc - total_aides)

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2)
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    # En-tête
    t_hdr = doc.add_table(rows=1, cols=3)
    t_hdr.style = 'Table Grid'
    t_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT
    c_logo = t_hdr.rows[0].cells[0]; c_logo.width = Cm(3.8)
    cell_bg(c_logo, 'FFFFFF'); cell_borders(c_logo, 'FFFFFF'); cell_pad(c_logo, 60,60,60,60)
    try:
        logo_bytes = base64.b64decode(LOGO_DMAAI_B64)
        pl = c_logo.paragraphs[0]; pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pl.add_run().add_picture(io.BytesIO(logo_bytes), width=Cm(3.4))
    except: pass

    c_mid = t_hdr.rows[0].cells[1]; c_mid.width = Cm(10)
    cell_bg(c_mid, '1F4E79'); cell_borders(c_mid, '1F4E79'); cell_pad(c_mid, 100,120,80,140)
    rt = c_mid.paragraphs[0].add_run("PLAN DE FINANCEMENT DÉFINITIF — MaPrimeAdapt'")
    rt.bold = True; rt.font.size = Pt(12); rt.font.color.rgb = BLANC
    p2 = c_mid.add_paragraph()
    rs = p2.add_run('EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911')
    rs.font.size = Pt(7.5); rs.font.color.rgb = RGBColor(0xB0, 0xC8, 0xE8)
    p3 = c_mid.add_paragraph()
    rh = p3.add_run(f'Habilitation ANAH N° {HABILIT_REF} · {HABILIT_DATE}')
    rh.font.size = Pt(7.5); rh.bold = True; rh.font.color.rgb = RGBColor(0xFA, 0xD7, 0xA0)

    c_mpa = t_hdr.rows[0].cells[2]; c_mpa.width = Cm(3.7)
    cell_bg(c_mpa, 'FFFFFF'); cell_borders(c_mpa, 'FFFFFF'); cell_pad(c_mpa, 60,60,60,60)
    try:
        mpa_bytes = base64.b64decode(LOGO_MPA_B64)
        pm = c_mpa.paragraphs[0]; pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pm.add_run().add_picture(io.BytesIO(mpa_bytes), width=Cm(3.2))
    except: pass

    doc.add_paragraph()

    # Bloc identification
    t_id = doc.add_table(rows=1, cols=2)
    t_id.style = 'Table Grid'
    t_id.alignment = WD_TABLE_ALIGNMENT.LEFT
    ci1 = t_id.rows[0].cells[0]; ci2 = t_id.rows[0].cells[1]
    ci1.width = Cm(8.5); ci2.width = Cm(9)
    for c in [ci1, ci2]:
        cell_bg(c, 'EBF3FB'); cell_borders(c, 'D6E4F0'); cell_pad(c, 80,120,80,120)

    def kv3(cell, label, val):
        if not val or str(val).strip() in ('','—','None'): return
        p_ = cell.add_paragraph()
        p_.paragraph_format.space_before = Pt(1)
        r1 = p_.add_run(f'{label} : '); r1.bold = True
        r1.font.size = Pt(9); r1.font.color.rgb = GRIS_T
        r2 = p_.add_run(str(val).strip()); r2.font.size = Pt(9)

    ci1.paragraphs[0].add_run('BÉNÉFICIAIRE').bold = True
    ci1.paragraphs[0].runs[0].font.size = Pt(8); ci1.paragraphs[0].runs[0].font.color.rgb = BLEU_TITRE
    kv3(ci1, 'Identité', nom)
    kv3(ci1, 'Adresse', adresse)
    kv3(ci1, 'N° dossier ANAH', dossier)

    recep_raw = b.get('recep', '')
    try:
        recep_str = datetime.strptime(recep_raw, '%Y-%m-%d').strftime('%d/%m/%Y') if recep_raw else '—'
    except:
        recep_str = recep_raw or '—'

    ci2.paragraphs[0].add_run('RÉCEPTION').bold = True
    ci2.paragraphs[0].runs[0].font.size = Pt(8); ci2.paragraphs[0].runs[0].font.color.rgb = BLEU_TITRE
    kv3(ci2, 'Date réception', recep_str)
    kv3(ci2, 'Entreprise(s)', b.get('entreprises', ''))
    kv3(ci2, 'Document établi le', datetime.today().strftime('%d/%m/%Y'))

    # Tableau factures
    add_h1(doc, 'FACTURES RÉELLES')
    if not factures:
        doc.add_paragraph('Aucune facture renseignée.').runs[0].font.size = Pt(9)
    else:
        t_fac = doc.add_table(rows=1, cols=5)
        t_fac.style = 'Table Grid'
        t_fac.alignment = WD_TABLE_ALIGNMENT.LEFT
        wf = [1.5, 5.5, 2.5, 1.5, 2.5]
        tbl_header_row(t_fac, ['N°', 'Désignation / Entreprise', 'Montant HT', 'TVA', 'Montant TTC'], wf)
        for i, f_ in enumerate(factures):
            ht  = parse_montant(f_.get('ht', 0))
            tva = parse_montant(f_.get('tva', 0))
            ttc = ht * (1 + tva/100)
            row = t_fac.add_row()
            bg  = 'FFFFFF' if i%2==0 else 'F5F8FC'
            for j,(val,w) in enumerate(zip([str(i+1), f_.get('label','—'), fmt_eur(ht), f'{tva:.0f} %', fmt_eur(ttc)], wf)):
                cl = row.cells[j]; cl.width = Cm(w)
                cell_bg(cl, bg); cell_borders(cl); cell_pad(cl, 60,80,60,80)
                p_ = cl.paragraphs[0]; run = p_.add_run(val); run.font.size = Pt(8.5)
                if j in (0,2,3,4): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j!=0 else WD_ALIGN_PARAGRAPH.CENTER

        # Honoraires AMO
        if hono_ht > 0:
            rh_ = t_fac.add_row()
            for j,(val,w) in enumerate(zip(['', 'Honoraires AMO — EI DMAAI (Mission AMO)', fmt_eur(hono_ht), f'{hono_tva:.0f} %', fmt_eur(hono_ttc)], wf)):
                cl = rh_.cells[j]; cl.width = Cm(w)
                cell_bg(cl, 'EBF3FB'); cell_borders(cl); cell_pad(cl, 60,80,60,80)
                p_ = cl.paragraphs[0]; run = p_.add_run(val)
                run.font.size = Pt(8.5); run.italic = True
                if j in (2,3,4): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Total
        rt_ = t_fac.add_row()
        for j,(val,w) in enumerate(zip(['', 'TOTAL', fmt_eur(total_charges_ht), '', fmt_eur(total_charges_ttc)], wf)):
            cl = rt_.cells[j]; cl.width = Cm(w)
            cell_bg(cl, 'D6E4F0'); cell_borders(cl, '2E75B6'); cell_pad(cl, 60,80,60,80)
            p_ = cl.paragraphs[0]; run = p_.add_run(val)
            run.bold = True; run.font.size = Pt(9); run.font.color.rgb = BLEU
            if j in (2,4): p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Plan financement définitif
    add_h1(doc, "PLAN DE FINANCEMENT DÉFINITIF")
    t_plan = doc.add_table(rows=1, cols=3)
    t_plan.style = 'Table Grid'
    t_plan.alignment = WD_TABLE_ALIGNMENT.LEFT
    wp = [9.0, 3.5, 3.5]
    tbl_header_row(t_plan, ['Poste', 'Référence', 'Montant'], wp, bg='2E75B6')

    rows_plan = [
        ("Coût total travaux + honoraires AMO (TTC)", f'Factures réelles', fmt_eur(total_charges_ttc)),
    ]
    if aide_mpa > 0:
        rows_plan.append(("— MaPrimeAdapt' travaux versée", "ANAH", f"- {fmt_eur(aide_mpa)}"))
    if aide_amo > 0:
        rows_plan.append(("— Forfait accompagnement AMO versé", "ANAH", f"- {fmt_eur(aide_amo)}"))
    for a in autres_aides:
        if parse_montant(a.get('montant', 0)) > 0:
            rows_plan.append((f"— {a.get('label', 'Autre aide')}", a.get('source', '—'), f"- {fmt_eur(parse_montant(a.get('montant', 0)))}"))

    for i,(label,ref,mont) in enumerate(rows_plan):
        row = t_plan.add_row()
        bg = 'EBF3FB' if i==0 else ('E2F0E8' if i>0 and i<len(rows_plan) else 'FFFFFF')
        for j,(val,w) in enumerate(zip([label,ref,mont], wp)):
            cl = row.cells[j]; cl.width = Cm(w)
            cell_bg(cl, bg); cell_borders(cl); cell_pad(cl, 80,100,80,100)
            p_ = cl.paragraphs[0]; run = p_.add_run(val); run.font.size = Pt(9)
            if j == 2: p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if i == 0: run.bold = True

    # Ligne reste à charge
    rt_ = t_plan.add_row()
    for j,(val,w) in enumerate(zip(['RESTE À CHARGE BÉNÉFICIAIRE', 'TOTAL', fmt_eur(reste_ch)], wp)):
        cl = rt_.cells[j]; cl.width = Cm(w)
        cell_bg(cl, 'C55A11' if reste_ch > 0 else '1B7A3C')
        cell_borders(cl, 'C55A11' if reste_ch > 0 else '1B7A3C')
        cell_pad(cl, 80,100,80,100)
        p_ = cl.paragraphs[0]; run = p_.add_run(val)
        run.bold = True; run.font.size = Pt(10); run.font.color.rgb = BLANC
        if j == 2: p_.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    add_sep(doc)
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(4)
    rn = p_note.add_run(
        "Ce plan de financement définitif est établi sur la base des factures réelles acquittées et des aides versées. "
        "Document établi par EI David MARIE-ANNE, AMO habilité ANAH, dans le cadre de la mission MaPrimeAdapt'."
    )
    rn.font.size = Pt(8); rn.italic = True; rn.font.color.rgb = GRIS_T

    # ─── Annexe photos ────────────────────────────────────────────────────────
    postes_avec_photos = [p for p in postes if p.get('photos')]
    if postes_avec_photos or photos_generales:
        doc.add_page_break()
        add_h1(doc, 'ANNEXE — PHOTOS APRÈS TRAVAUX')

        for p_ in postes_avec_photos:
            ph_list = p_.get('photos', [])
            if not ph_list:
                continue
            p_tit = doc.add_paragraph()
            p_tit.paragraph_format.space_before = Pt(8)
            p_tit.paragraph_format.space_after  = Pt(2)
            rt = p_tit.add_run(p_.get('nom', ''))
            rt.bold = True; rt.font.size = Pt(9); rt.font.color.rgb = BLEU_TITRE
            n = min(len(ph_list), 4)
            cols = 2
            rows_needed = (n + cols - 1) // cols
            t_ph = doc.add_table(rows=rows_needed, cols=cols)
            t_ph.style = 'Table Grid'
            idx = 0
            for ri in range(rows_needed):
                for ci in range(cols):
                    if idx >= n:
                        break
                    cell = t_ph.rows[ri].cells[ci]
                    cell.width = Cm(8.5)
                    cell_bg(cell, 'F5F9FC'); cell_borders(cell); cell_pad(cell, 60, 60, 60, 60)
                    raw, _ = b64_to_bytes(ph_list[idx].get('data', ''))
                    if raw:
                        try:
                            p_img = cell.paragraphs[0]
                            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_img.add_run().add_picture(io.BytesIO(raw), width=Cm(7.5))
                        except Exception:
                            cell.paragraphs[0].add_run(ph_list[idx].get('name', '(photo)'))
                    p_leg = cell.add_paragraph()
                    rl = p_leg.add_run(ph_list[idx].get('name', ''))
                    rl.font.size = Pt(7); rl.font.color.rgb = GRIS_T
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    idx += 1

        if photos_generales:
            p_tit = doc.add_paragraph()
            p_tit.paragraph_format.space_before = Pt(10)
            p_tit.paragraph_format.space_after  = Pt(2)
            rt = p_tit.add_run('Photos générales de chantier')
            rt.bold = True; rt.font.size = Pt(9); rt.font.color.rgb = BLEU_TITRE
            n = min(len(photos_generales), 6)
            cols = 2
            rows_needed = (n + cols - 1) // cols
            t_gen = doc.add_table(rows=rows_needed, cols=cols)
            t_gen.style = 'Table Grid'
            idx = 0
            for ri in range(rows_needed):
                for ci in range(cols):
                    if idx >= n:
                        break
                    cell = t_gen.rows[ri].cells[ci]
                    cell.width = Cm(8.5)
                    cell_bg(cell, 'F5F9FC'); cell_borders(cell); cell_pad(cell, 60, 60, 60, 60)
                    raw, _ = b64_to_bytes(photos_generales[idx].get('data', ''))
                    if raw:
                        try:
                            p_img = cell.paragraphs[0]
                            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_img.add_run().add_picture(io.BytesIO(raw), width=Cm(7.5))
                        except Exception:
                            cell.paragraphs[0].add_run(photos_generales[idx].get('name', '(photo)'))
                    p_leg = cell.add_paragraph()
                    rl = p_leg.add_run(photos_generales[idx].get('name', ''))
                    rl.font.size = Pt(7); rl.font.color.rgb = GRIS_T
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    idx += 1

    add_sep(doc)
    pied = doc.add_paragraph()
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp = pied.add_run(
        'EI David MARIE-ANNE · Architecte DPLG N°049576 · Ingénieur TCE OPQIBI 1905/1911 · SIRET 808 408 165 000 13\n'
        'ei_dmaai@outlook.fr · 06 34 40 79 32 · Orléans'
    )
    rp.font.size = Pt(7.5); rp.font.color.rgb = GRIS_T

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


@app.route('/plan_financement', methods=['POST'])
def plan_financement():
    try:
        data       = request.get_json(force=True)
        docx_bytes = generer_plan(data)
        flds = data.get('flds') or data.get('data', {}).get('fields', {})
        nom  = flds.get('d_nom',  flds.get('dem_nom',  'X')).replace(' ','_')
        prn  = flds.get('d_prn',  flds.get('dem_prenom','') ).replace(' ','_')
        date = flds.get('v_date', flds.get('visite_date','2026'))
        filename = f'Plan_Financement_MPA_{nom}_{prn}_{date}.docx'
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500



def fmt_eur(val: float) -> str:
    """Formate un montant en euros format français : 1 585,60 €"""
    # Utiliser les locales python ou formatage manuel
    s = f"{val:,.2f}"           # "1,585.60"
    s = s.replace(',', ' ').replace('.', ',')  # "1 585,60"
    return s + " \u20ac"


def parse_montant(val) -> float:
    """Convertit un montant en float, gère virgule et point comme séparateurs."""
    if val is None: return 0.0
    s = str(val).strip()
    if not s: return 0.0
    # Supprimer espaces, €, symboles
    s = s.replace('\xa0', '').replace(' ', '').replace('\u202f', '').replace('€', '')
    # Si virgule ET point : format anglais "1,585.60" → supprimer virgule
    if ',' in s and '.' in s:
        s = s.replace(',', '')
    # Si seule virgule : format français "1585,60" → remplacer par point
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0



@app.route('/')
def serve_html():
    import os
    folder = os.path.dirname(os.path.abspath(__file__))
    for name in ['agent_diagnostic_mpa.html', 'diagnostic_mpa.html']:
        p = os.path.join(folder, name)
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                content = f.read()
            from flask import Response
            return Response(content, mimetype='text/html; charset=utf-8')
    htmls = [f for f in os.listdir(folder) if f.endswith('.html')]
    return f'<h2>HTML non trouvé.</h2><p>Dossier: {folder}</p><p>HTML dispo: {htmls}</p>', 404

@app.route('/ping', methods=['GET'])
def ping():
    return jsonify({'status': 'ok', 'service': 'Diagnostic MaPrimeAdapt — DMAAI'})

@app.route('/generer', methods=['POST'])
def generer():
    try:
        data       = request.get_json(force=True)
        docx_bytes = generer_docx(data)
        # Nom du fichier
        flds = data.get('flds') or data.get('data', {}).get('fields', {})
        nom  = flds.get('d_nom',  flds.get('dem_nom',  'X')).replace(' ','_')
        prn  = flds.get('d_prn',  flds.get('dem_prenom','') ).replace(' ','_')
        date = flds.get('v_date', flds.get('visite_date','2026'))
        filename = f'Rapport_MPA_{nom}_{prn}_{date}.docx'
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


# ─── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print('=' * 60)
    print('  Serveur MaPrimeAdapt — DMAAI')
    print('  Diagnostic : http://localhost:5052')
    print('  Reception  : http://localhost:5052/reception')
    print('=' * 60)
    try:
        import flask_cors
    except ImportError:
        print('\n⚠️  Dépendances manquantes. Lancez :')
        print('   pip install flask flask-cors python-docx')
        sys.exit(1)
    app.run(host='localhost', port=5052, debug=False)
