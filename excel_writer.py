"""
Excel Writer v12 — Médiathèque
════════════════════════════════════════════════════════════════
Nouveautés v12 :
  - Planning_Agent : lignes 8h30-9h et 9h-10h (heure d'arrivée)
  - Nouvelles couleurs : lavande Jeunesse, hachures hors-horaires,
    rouge doux congés
  - Totaux SP dynamiques (formules IF) dans Planning_Agent
  - Récap Semaine_N dynamique (références cross-sheet)
════════════════════════════════════════════════════════════════
"""

import io
import re
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from collections import defaultdict

SECTIONS   = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SP   = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
MOIS_FR    = {1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
              7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',
              11:'Novembre',12:'Décembre'}

# ── Palette ────────────────────────────────────────────────────
C = {
    'hdr_dark':   '1A2F4A', 'hdr_rouge':  'C0392B', 'hdr_bleu':   '2471A3',
    'rdc':        'D6E8F7', 'adulte':     'D4EDD4', 'mf':         'FFF0CC',
    # Jeunesse : lavande douce (remplace rose/rouge)
    'jeunesse':   'FFE0E0',          # onglet Semaine (inchangé, repère visuel)
    'jeunesse_ag':'E8DAEF',          # Planning_Agent : lavande
    'closed':     'EBEBEB',
    'event_bg':   'FFF3CD', 'agent_txt': '2C3E50',
    'gray':       'AAAAAA', 'alt1':      'FDFEFE', 'alt2':        'EBF5FB',
    'dark':       '2C3E50', 'sp_ok':     'D5F5E3', 'sp_warn':     'FFF3CD',
    'sp_alert':   'FFCCCC', 'sp_ok_t':   '1E8449', 'sp_warn_t':   'E67E22',
    'sp_alert_t': '8B0000',
    'hdr_rdc':    '2980B9', 'hdr_adulte':'27AE60', 'hdr_mf':      'E67E22',
    'hdr_j':      'E74C3C', 'hdr_purple':'8E44AD', 'hdr_dark2':   '34495E',
    'oos_bg':     'FF0000', 'oos_txt':   'FFFFFF',
    'vac_bg':     'F0E6FF', 'vac_txt':   '6A0DAD',
    # Congé : rouge très doux (remplace FFCCCC/8B0000)
    'conge_bg':   'FADBD8', 'conge_txt': '922B21',
    # Hors-horaires : hachures (défini via _hatch())
    'bureau_bg':  'FEF9E7',          # au bureau (présent, hors SP) → jaune doux
    'off_bg':     'D5D8DC',          # off / hors horaires → gris uniforme
    'off_txt':    '999999',          # texte tiret gris
    # Arrivée
    'arrival_bg': 'EAF2FF',
    'sp_dyn_bg':  'E8DAEF',   # lavande neutre pour cellules SP dynamiques
    'arrival_txt':'1F4E79',
}
SEC_BG  = {'RDC': C['rdc'],      'Adulte': C['adulte'],
           'MF':  C['mf'],       'Jeunesse': C['jeunesse']}
SEC_BG_AG = {'RDC': C['rdc'],    'Adulte': C['adulte'],
             'MF':  C['mf'],     'Jeunesse': C['jeunesse_ag']}
SEC_HDR = {'RDC': C['hdr_rdc'],  'Adulte': C['hdr_adulte'],
           'MF':  C['hdr_mf'],   'Jeunesse': C['hdr_j']}


# ── Helpers styles ─────────────────────────────────────────────
def _fill(h):
    return PatternFill('solid', fgColor=h)

def _hatch(fg='BBBBBB', bg='F5F5F5'):
    """Motif hachuré diagonal — représente hors-horaires."""
    return PatternFill(patternType='lightDown', fgColor=fg, bgColor=bg)

def _font(bold=False, size=10, color='000000', italic=False, name='Calibri'):
    return Font(bold=bold, size=size, color=color, italic=italic, name=name)

def _aln(h='center', v='center', wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _brd(color='CCCCCC', style='thin'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _set(cell, value=None, bg=None, fnt=None, aln=None, brd=None, hatch=False):
    if value is not None:
        cell.value = value
    if hatch:
        cell.fill = _hatch()
    elif bg:
        cell.fill = _fill(bg)
    if fnt:
        cell.font = fnt
    if aln:
        cell.alignment = aln
    if brd:
        cell.border = brd


def is_vacataire(agent):
    return 'vacataire' in str(agent).lower()

def min_to_hhmm(m):
    return f"{int(m)//60}h{int(m)%60:02d}" if m else '0h00'

def min_to_dec(m):
    return round(m / 60, 2) if m else 0.0


# ══════════════════════════════════════════════════════════════
#  ONGLET SEMAINE
# ══════════════════════════════════════════════════════════════

def write_week_sheet(wb, week_data, metadata, agent_sp_cells=None):
    """
    agent_sp_cells : dict {agent: {jour: 'C42'}} — références dans
    Planning_Agent_Semaine_N pour le récap dynamique.
    Si None : utilise les valeurs statiques du moteur.
    """
    week_num       = week_data['week_num']
    week_dates     = week_data['week_dates']
    plan           = week_data['plan']
    sp_jour_cum    = week_data.get('sp_jour_cum', {})
    sp_alerts      = week_data.get('sp_alerts', {})
    creneaux       = metadata['creneaux']
    sp_minmax_all  = metadata.get('sp_minmax_all', {})
    sp_minmax_week = sp_minmax_all.get(week_num, sp_minmax_all.get(2, {}))
    affectations   = metadata.get('affectations', {})
    is_sam_week    = 'Samedi' in week_dates

    mois_num = {'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
                'juillet':7,'août':8,'septembre':9,'octobre':10,
                'novembre':11,'décembre':12
                }.get(metadata['mois'].lower(), 1)
    mois_cap  = MOIS_FR.get(mois_num, metadata['mois'].capitalize())
    annee     = metadata['annee']
    sem_type  = week_data.get('semaine_type', '')

    dates_list = sorted([d for d in week_dates.values() if d])
    if dates_list:
        d0, d1 = dates_list[0], dates_list[-1]
        m2 = MOIS_FR.get(d1.month, '')
        date_range = (f"{d0.day} au {d1.day} {mois_cap} {annee}"
                      if d0.month == d1.month
                      else f"{d0.day} {mois_cap} au {d1.day} {m2} {annee}")
    else:
        date_range = f"Semaine {week_num}"

    sname = f"Semaine_{week_num}"
    if sname in wb.sheetnames:
        wb.remove(wb[sname])
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False

    for col, w in {'A':5,'B':14,'C':20,'D':20,'E':20,'F':26,'G':28,'H':30}.items():
        ws.column_dimensions[col].width = w

    # Mapping jour → (row_cren_start, row_cren_end) dans cet onglet
    # Utilisé par Planning_Agent pour les formules cross-sheet
    jour_cren_rows = {}

    row = 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value=f"PLANNING SERVICE PUBLIC — Semaine {week_num}  |  {date_range}  |  {sem_type}"),
         bg=C['hdr_dark'], fnt=_font(bold=True, size=13, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="  RDC     Adulte     Musique & Films     Jeunesse     "
                       "Fermé     Événement     ALERTE     Rouge=hors section"),
         bg=C['dark'], fnt=_font(size=9, color='FFFFFF', italic=True), aln=_aln('left'))
    ws.row_dimensions[row].height = 15
    row += 2

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        jour_plan   = plan.get(jour, {})
        samedi_type = jour_plan.get('_samedi_type')
        is_sam      = jour == 'Samedi'

        j_label = f"  {jour.upper()}  {date.day} {MOIS_FR.get(date.month,'')} {date.year}"
        if is_sam and samedi_type:
            j_label += f"  —  SAMEDI {samedi_type}"
        j_color = (C['hdr_rouge'] if (is_sam and samedi_type == 'ROUGE')
                   else C['hdr_bleu']  if (is_sam and samedi_type == 'BLEU')
                   else C['hdr_dark'])

        ws.merge_cells(f'A{row}:H{row}')
        _set(ws.cell(row=row, column=1, value=j_label),
             bg=j_color, fnt=_font(bold=True, size=12, color='FFFFFF'), aln=_aln('left'))
        ws.row_dimensions[row].height = 26
        row += 1

        hdrs = [('N°',C['dark']),('Créneau',C['dark']),('RDC',SEC_HDR['RDC']),
                ('Adulte',SEC_HDR['Adulte']),('M & F',SEC_HDR['MF']),
                ('Jeunesse',SEC_HDR['Jeunesse']),('Événement',C['hdr_purple']),
                ('Alertes',C['hdr_rouge'])]
        for ci,(h_txt,hc) in enumerate(hdrs):
            _set(ws.cell(row=row, column=ci+1, value=h_txt),
                 bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'),
                 aln=_aln('center'), brd=_brd())
        ws.row_dimensions[row].height = 20
        row += 1

        # Mémoriser la première ligne de créneaux pour ce jour
        _cren_start_row = row

        for ni,(cren_name,cs,ce) in enumerate(creneaux):
            slot = jour_plan.get(cren_name)
            if slot is None:
                for ci,val in enumerate([ni+1, cren_name,'—','—','—','—','',''],1):
                    _set(ws.cell(row=row, column=ci, value=val or None),
                         bg=C['closed'],
                         fnt=_font(size=9 if ci>2 else 10, color=C['gray'], italic=ci>1),
                         aln=_aln(), brd=_brd())
            else:
                assgn  = slot.get('assignment', {s:[] for s in SECTIONS})
                events = slot.get('events', [])
                alerts = slot.get('alerts', [])
                oos    = slot.get('out_of_section', {})

                _set(ws.cell(row=row, column=1, value=ni+1),
                     bg='F2F3F4', fnt=_font(size=9, color='7F8C8D'), aln=_aln(), brd=_brd())
                _set(ws.cell(row=row, column=2, value=cren_name),
                     bg='FDFEFE', fnt=_font(bold=True, size=10), aln=_aln(), brd=_brd())

                for si, sect in enumerate(SECTIONS):
                    col = si + 3
                    agents  = assgn.get(sect, [])
                    has_al  = any(sect in al for al in alerts)
                    is_oos  = oos.get(sect, False)

                    if is_oos and agents:
                        val = '⚠ ' + '  /  '.join(agents)
                        bg  = C['oos_bg']
                        fc  = C['oos_txt']
                    elif agents:
                        val = '  /  '.join(agents)
                        bg  = SEC_BG[sect]
                        fc  = C['vac_txt'] if all(is_vacataire(a) for a in agents) else C['agent_txt']
                    elif has_al:
                        val, bg, fc = 'ALERTE', C['conge_bg'], C['sp_alert_t']
                    else:
                        val, bg, fc = '—', 'FDF2F8', 'C0392B'

                    _set(ws.cell(row=row, column=col, value=val),
                         bg=bg, fnt=_font(size=10, color=fc, bold=bool(agents)),
                         aln=_aln('center'), brd=_brd())

                ev_txt = '\n'.join(f"{ev['nom']} ({', '.join(ev['agents'][:3])})"
                                    for ev in events) if events else None
                _set(ws.cell(row=row, column=7, value=ev_txt),
                     bg=C['event_bg'] if ev_txt else 'FAFAFA',
                     fnt=_font(size=9, italic=True,
                               color='6E2F09' if ev_txt else C['gray']),
                     aln=_aln('left'), brd=_brd())

                al_txt = '\n'.join(alerts) if alerts else None
                _set(ws.cell(row=row, column=8, value=al_txt),
                     bg=C['conge_bg'] if al_txt else 'FAFAFA',
                     fnt=_font(size=9, bold=bool(al_txt),
                               color=C['sp_alert_t'] if al_txt else C['gray']),
                     aln=_aln('left'), brd=_brd())

                h = max(20, 16*max(len(events), len(alerts), 1))
                ws.row_dimensions[row].height = h
            row += 1

        # Mémoriser la dernière ligne de créneaux pour ce jour
        jour_cren_rows[jour] = (_cren_start_row, row - 1)

        for ci in range(1, 9):
            ws.cell(row=row, column=ci).fill = _fill('F0F3F4')
        ws.row_dimensions[row].height = 4
        row += 1

    # ── Récap SP ────────────────────────────────────────────────
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="RÉCAP — Heures de Service Public par agent (heures décimales)"),
         bg='34495E', fnt=_font(bold=True, size=11, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 24
    row += 1

    hdrs2 = ['Agent','Section','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Total (h)']
    hcols = ['34495E','34495E','1A5276','1A5276','1A5276','1A5276','922B21','145A32']
    for ci,(h_txt,hc) in enumerate(zip(hdrs2,hcols)):
        _set(ws.cell(row=row, column=ci+1, value=h_txt),
             bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'),
             aln=_aln(), brd=_brd())
    ws.row_dimensions[row].height = 20
    row += 1

    def primary_section(agent):
        sects = affectations.get(agent, [])
        return sects[0] if sects else '?'

    agents_shown = sorted(
        [a for a in sp_minmax_week.keys() if not is_vacataire(a)],
        key=lambda x: (primary_section(x), x)
    ) + [a for a in affectations.keys() if is_vacataire(a)]

    pa_sheet = f"Planning_Agent_Semaine_{week_num}"

    for i, agent in enumerate(agents_shown):
        bg   = C['alt1'] if i % 2 == 0 else C['alt2']
        sect = primary_section(agent)

        _set(ws.cell(row=row, column=1, value=agent),
             bg=bg, fnt=_font(bold=True, size=10), aln=_aln('left'), brd=_brd())
        _set(ws.cell(row=row, column=2,
                     value='VAC' if is_vacataire(agent) else sect),
             bg=C['vac_bg'] if is_vacataire(agent) else SEC_BG.get(sect, bg),
             fnt=_font(size=10, color=C['vac_txt'] if is_vacataire(agent) else '000000'),
             aln=_aln(), brd=_brd())

        total_h_static = 0.0
        for ji, jour in enumerate(JOURS_SP):
            col  = ji + 3
            cell = ws.cell(row=row, column=col)
            sp_min = sp_jour_cum.get(jour, {}).get(agent, 0)
            sp_h   = min_to_dec(sp_min)
            total_h_static += sp_h

            if week_dates.get(jour) is not None:
                # Formule dynamique si on a les refs Planning_Agent
                if agent_sp_cells and agent in agent_sp_cells:
                    cell_ref = agent_sp_cells[agent].get(jour)
                    if cell_ref:
                        cell.value = f"='{pa_sheet}'!{cell_ref}"
                        cell.number_format = '0.00'
                    else:
                        cell.value = sp_h if sp_h > 0 else 0
                        cell.number_format = '0.00'
                else:
                    cell.value = sp_h if sp_h > 0 else 0
                    cell.number_format = '0.00'

                bg_c = ('FFF9C4' if jour == 'Samedi' and sp_h > 0
                         else C['alt2'] if sp_h > 0 else bg)
                cell.fill      = _fill(bg_c)
                cell.font      = _font(size=10, bold=sp_h > 0)
                cell.alignment = _aln()
                cell.border    = _brd()
            else:
                cell.value     = '—'
                cell.fill      = _fill(bg)
                cell.font      = _font(size=9, color=C['gray'])
                cell.alignment = _aln()
                cell.border    = _brd()

        # Total — formule SUM des colonnes jour
        tc = ws.cell(row=row, column=8)
        # Colonnes C..G (indices 3..7) = 5 jours max
        # On somme les colonnes qui ont une date
        sum_cols = [get_column_letter(3 + ji)
                    for ji, jour in enumerate(JOURS_SP)
                    if week_dates.get(jour) is not None]
        if sum_cols:
            tc.value = f"=SUM({','.join(c+str(row) for c in sum_cols)})"
        else:
            tc.value = 0
        tc.number_format = '0.00'

        # Coloration total (basée sur valeur moteur pour le style)
        total_h = total_h_static
        if not is_vacataire(agent):
            mm     = sp_minmax_week.get(agent, {})
            min_sp = (mm.get('Min_MarSam',0) if is_sam_week else mm.get('Min_MarVen',0))
            max_sp = (mm.get('Max_MarSam',99) if is_sam_week else mm.get('Max_MarVen',99))
            if agent in sp_alerts:
                t_bg, t_fc = C['sp_alert'], C['sp_alert_t']
            elif total_h < min_sp:
                t_bg, t_fc = C['sp_alert'], C['sp_alert_t']
            elif total_h > max_sp:
                t_bg, t_fc = C['sp_warn'], C['sp_warn_t']
            else:
                t_bg, t_fc = C['sp_ok'], C['sp_ok_t']
        else:
            t_bg = C['vac_bg'] if total_h > 0 else bg
            t_fc = C['vac_txt'] if total_h > 0 else C['gray']

        tc.fill = _fill(t_bg)
        tc.font = _font(bold=True, size=10, color=t_fc)
        tc.alignment = _aln()
        tc.border    = _brd()
        ws.row_dimensions[row].height = 18
        row += 1

    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="Vert=OK  |  Rouge=sous min  |  Orange=sur max  |  "
                       "Violet=vacataire  |  Rouge vif cellule planning=agent hors section"),
         bg='F4F6F7', fnt=_font(size=9, italic=True, color='5D6D7E'), aln=_aln('left'))
    ws.row_dimensions[row].height = 16
    return ws, jour_cren_rows


# ══════════════════════════════════════════════════════════════
#  ONGLET PLANNING PAR AGENT — 1 PAR SEMAINE
# ══════════════════════════════════════════════════════════════

def write_planning_agent_week_sheet(wb, week_data, metadata, jour_cren_rows=None):
    """
    Retourne agent_sp_cells = {agent: {jour: 'C42'}}
    pour utilisation dans les formules cross-sheet du récap Semaine_N.
    """
    week_num     = week_data['week_num']
    week_dates   = week_data['week_dates']
    plan         = week_data['plan']
    sp_jour_cum  = week_data.get('sp_jour_cum', {})
    creneaux     = metadata['creneaux']
    affectations = metadata['affectations']
    horaires_ag  = metadata['horaires_agents']
    evenements   = metadata['evenements']
    sem_type     = week_data.get('semaine_type', '')

    mois_num = {'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
                'juillet':7,'août':8,'septembre':9,'octobre':10,
                'novembre':11,'décembre':12
                }.get(metadata['mois'].lower(), 1)
    mois_cap = MOIS_FR.get(mois_num, '')
    annee    = metadata['annee']

    sname = f"Planning_Agent_Semaine_{week_num}"
    if sname in wb.sheetnames:
        wb.remove(wb[sname])
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False

    jours_presents = [j for j in JOURS_SP if week_dates.get(j) is not None]
    nb_jours       = len(jours_presents)
    nb_cols        = 2 + nb_jours
    end_col        = get_column_letter(nb_cols)

    agents_sorted = sorted(
        [a for a in affectations.keys() if not is_vacataire(a)],
        key=lambda x: (affectations.get(x, [''])[0], x)
    ) + [a for a in affectations.keys() if is_vacataire(a)]

    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 15
    for ci in range(3, 3 + nb_jours):
        ws.column_dimensions[get_column_letter(ci)].width = 20

    # Mapping jour → colonne lettre
    col_map = {}
    for ji, jour in enumerate(jours_presents):
        col_map[jour] = 3 + ji

    row = 1

    # ── Titre global ────────────────────────────────────────────
    ws.merge_cells(f'A{row}:{end_col}{row}')
    dates_list = sorted([d for d in week_dates.values() if d])
    d0, d1 = dates_list[0], dates_list[-1]
    titre = (f"PLANNING PAR AGENT — Semaine {week_num}  |  "
             f"{d0.day} au {d1.day} {mois_cap} {annee}  |  {sem_type}")
    _set(ws.cell(row=row, column=1, value=titre),
         bg=C['hdr_dark'], fnt=_font(bold=True, size=13, color='FFFFFF'),
         aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    # ── Légende ─────────────────────────────────────────────────
    ws.merge_cells(f'A{row}:{end_col}{row}')
    _set(ws.cell(row=row, column=1,
                 value="  SP = section (couleur)     Arr. XXhYY = arrivée hors heure ronde     "
                       "Jaune = au bureau (hors SP)     Gris — = absent / hors horaires     "
                       "Rose = Congé/RTT     Rouge vif = hors section habituelle"),
         bg='34495E', fnt=_font(size=9, italic=True, color='FFFFFF'), aln=_aln('left'))
    ws.row_dimensions[row].height = 15
    row += 2

    # Créneaux additionnels avant 10h
    EARLY_SLOTS = [
        ('8:30-9:00',  510, 540),
        ('9:00-10:00', 540, 600),
    ]

    # Durées en MINUTES entières pour les créneaux SP (évite pb décimal Excel FR)
    cren_durations_min = [(ce - cs) for (_, cs, ce) in creneaux]

    # Résultat : {agent: {jour: 'C42'}} pour cross-sheet refs
    agent_sp_cells = {}

    for agent in agents_sorted:
        sects      = affectations.get(agent, [])
        sect_prim  = sects[0] if sects else '?'
        hdr_bg     = C['vac_bg'] if is_vacataire(agent) else SEC_HDR.get(sect_prim, C['hdr_dark'])
        hdr_fc     = C['vac_txt'] if is_vacataire(agent) else 'FFFFFF'

        # ── En-tête agent ───────────────────────────────────────
        ws.merge_cells(f'A{row}:{end_col}{row}')
        sects_label = ' / '.join(sects)
        _set(ws.cell(row=row, column=1,
                     value=f"  {agent.upper()}   ({sects_label})"),
             bg=hdr_bg, fnt=_font(bold=True, size=11, color=hdr_fc),
             aln=_aln('left'))
        ws.row_dimensions[row].height = 24
        row += 1

        # ── En-tête colonnes ────────────────────────────────────
        ws.cell(row=row, column=1).fill = _fill(C['hdr_dark2'])
        ws.cell(row=row, column=1).value = ''
        _set(ws.cell(row=row, column=2, value='Créneau'),
             bg=C['hdr_dark2'], fnt=_font(bold=True, size=10, color='FFFFFF'),
             aln=_aln(), brd=_brd())
        for jour in jours_presents:
            c    = col_map[jour]
            date = week_dates.get(jour)
            sam_t = plan.get(jour, {}).get('_samedi_type')
            lbl  = f"{jour[:3]}\n{date.day}/{date.month}"
            if sam_t:
                lbl += f"\n{sam_t}"
            bg_h = (C['hdr_rouge'] if sam_t == 'ROUGE'
                    else C['hdr_bleu'] if sam_t == 'BLEU'
                    else C['hdr_dark2'])
            _set(ws.cell(row=row, column=c, value=lbl),
                 bg=bg_h, fnt=_font(bold=True, size=9, color='FFFFFF'),
                 aln=_aln('center'), brd=_brd())
        ws.row_dimensions[row].height = 36
        row += 1

        # ── Lignes d'arrivée (8h30-9h et 9h-10h) ───────────────
        # Même formatage colonne B que les créneaux SP
        for slot_name, es, ee in EARLY_SLOTS:
            _set(ws.cell(row=row, column=1), bg='F8F9FA')
            ws.cell(row=row, column=1).value = ''
            # Colonne B : même style que les créneaux SP normaux
            _set(ws.cell(row=row, column=2, value=slot_name),
                 bg='FDFEFE', fnt=_font(bold=True, size=9),
                 aln=_aln(), brd=_brd())

            for jour in jours_presents:
                c        = col_map[jour]
                date     = week_dates.get(jour)
                date_str = date.strftime('%Y-%m-%d')
                cell     = ws.cell(row=row, column=c)

                # Vérifier congé sur la journée entière
                # (congé = événement couvrant le début de journée ou le créneau)
                ev_all = evenements.get(date_str, [])
                ev_agent = [ev for ev in ev_all if agent in ev['agents']]
                is_conge = any(
                    n.lower() in ('congé','conge','rtt','vacation')
                    for ev in ev_agent
                    for n in [ev['nom']]
                )

                # Horaires habituels de l'agent ce jour
                h  = horaires_ag.get(agent, {}).get(jour)
                sm = h[0] if h else None   # début matin (en minutes)

                if is_conge:
                    # Congé : affiché dès 8h30 si l'agent travaille habituellement
                    _set(cell, value='Congé',
                         bg=C['conge_bg'],
                         fnt=_font(size=9, color=C['conge_txt'], italic=True),
                         aln=_aln(), brd=_brd())

                elif sm is not None and es <= sm < ee:
                    # L'agent arrive pendant ce créneau
                    # → mention précise SEULEMENT si heure ≠ début exact du créneau
                    if sm == es:
                        # Arrive pile à l'heure du créneau → cellule bureau (blanc)
                        _set(cell, value='',
                             bg=C['bureau_bg'],
                             fnt=_font(size=8, color=C['gray']),
                             aln=_aln(), brd=_brd())
                    else:
                        # Arrive en cours de créneau → indiquer l'heure précise
                        _set(cell, value=f"Arr. {min_to_hhmm(sm)}",
                             bg=C['arrival_bg'],
                             fnt=_font(size=9, bold=True, color=C['arrival_txt']),
                             aln=_aln(), brd=_brd())

                elif sm is not None and sm < es:
                    # Agent déjà arrivé avant ce créneau → au bureau (blanc)
                    _set(cell, value='',
                         bg=C['bureau_bg'],
                         fnt=_font(size=8, color=C['gray']),
                         aln=_aln(), brd=_brd())

                else:
                    # Pas encore arrivé / ne travaille pas ce jour → gris + tiret
                    _set(cell, value='—',
                         bg=C['off_bg'],
                         fnt=_font(size=10, color=C['off_txt']),
                         aln=_aln(), brd=_brd())

            ws.row_dimensions[row].height = 26
            row += 1

        # ── Lignes créneaux SP (10h → 19h) ─────────────────────
        cren_start_row = row  # première ligne SP (pour la formule)

        for cren_idx, (cren_name, cs, ce) in enumerate(creneaux):
            # Colonne A : durée en heures décimales (police blanche = invisible)
            # Utilisée par la formule SUMPRODUCT pour éviter les tableaux littéraux
            dur_h = round((ce - cs) / 60, 4)
            dur_cell = ws.cell(row=row, column=1)
            dur_cell.value = dur_h
            dur_cell.number_format = '0.00'
            dur_cell.fill = _fill('F8F9FA')
            dur_cell.font = _font(size=8, color='F8F9FA')  # invisible (blanc sur blanc)
            dur_cell.alignment = _aln()
            _set(ws.cell(row=row, column=2, value=cren_name),
                 bg='FDFEFE', fnt=_font(bold=True, size=9), aln=_aln(), brd=_brd())

            for jour in jours_presents:
                c        = col_map[jour]
                date     = week_dates.get(jour)
                date_str = date.strftime('%Y-%m-%d')
                cell     = ws.cell(row=row, column=c)
                slot     = plan.get(jour, {}).get(cren_name)

                ev_names = [ev['nom'] for ev in evenements.get(date_str, [])
                            if agent in ev['agents']
                            and ev['debut'] <= cs < ev['fin']]
                is_conge = any(n.lower() in ('congé','conge','rtt','vacation')
                               for n in ev_names)
                event    = [n for n in ev_names
                            if n.lower() not in ('congé','conge','rtt','vacation')]

                h = horaires_ag.get(agent, {}).get(jour)
                in_horaires = False
                if h:
                    sm,em,sa,ea = h
                    in_m = sm is not None and em is not None and cs >= sm and ce <= em
                    in_a = sa is not None and ea is not None and cs >= sa and ce <= ea
                    in_horaires = in_m or in_a

                sp_section  = None
                is_oos_cell = False
                if slot and isinstance(slot, dict) and slot.get('open'):
                    for sect in SECTIONS:
                        if agent in slot.get('assignment', {}).get(sect, []):
                            sp_section = sect
                            break
                    if sp_section:
                        is_oos_cell = slot.get('out_of_section', {}).get(sp_section, False)

                if is_conge:
                    _set(cell, value='Congé',
                         bg=C['conge_bg'],
                         fnt=_font(size=9, color=C['conge_txt'], italic=True),
                         aln=_aln(), brd=_brd())
                elif event:
                    ev_label = f"Évén.\n{event[0][:18]}"
                    _set(cell, value=ev_label,
                         bg=C['event_bg'],
                         fnt=_font(size=9, italic=True, color='6E2F09'),
                         aln=_aln(), brd=_brd())
                elif sp_section and is_oos_cell:
                    # Hors section habituelle → statique (cas rare, pas de formule)
                    _set(cell, value=f"SP\n{sp_section}\n⚠hors sect.",
                         bg=C['oos_bg'],
                         fnt=_font(size=8, bold=True, color=C['oos_txt']),
                         aln=_aln(), brd=_brd())
                elif not in_horaires:
                    # Hors horaires → statique gris + tiret
                    _set(cell, value='—',
                         bg=C['off_bg'],
                         fnt=_font(size=10, color=C['off_txt']),
                         aln=_aln(), brd=_brd())
                else:
                    # SP ou bureau → formule dynamique si jour_cren_rows dispo
                    # =IF(SUMPRODUCT(ISNUMBER(SEARCH("Agent",'Sem'!C{r}:F{r})))>0,"SP","")
                    sem_row = None
                    if jour_cren_rows and jour in jour_cren_rows:
                        r1_sem, _ = jour_cren_rows[jour]
                        sem_row   = r1_sem + cren_idx   # ligne exacte dans Semaine_N
                    if sem_row is not None:
                        sem_sht = 'Semaine_' + str(week_num)
                        ag = agent.replace('"', '""')
                        # IF imbrique: cherche agent col C(RDC) D(Adulte) E(MF) F(Jeunesse)
                        # CHAR(10) = saut de ligne (EN); Excel FR accepte les deux
                        p1 = 'ISNUMBER(SEARCH("' + ag + '",\'' + sem_sht + '\'!C' + str(sem_row) + '))'
                        p2 = 'ISNUMBER(SEARCH("' + ag + '",\'' + sem_sht + '\'!D' + str(sem_row) + '))'
                        p3 = 'ISNUMBER(SEARCH("' + ag + '",\'' + sem_sht + '\'!E' + str(sem_row) + '))'
                        p4 = 'ISNUMBER(SEARCH("' + ag + '",\'' + sem_sht + '\'!F' + str(sem_row) + '))'
                        formula = (
                            '=IF(' + p1 + ',"SP"&CHAR(10)&"RDC",'
                            'IF(' + p2 + ',"SP"&CHAR(10)&"Adulte",'
                            'IF(' + p3 + ',"SP"&CHAR(10)&"M & F",'
                            'IF(' + p4 + ',"SP"&CHAR(10)&"Jeunesse",""))))'
                        )
                        cell.value     = formula
                        cell.alignment = _aln()
                        cell.border    = _brd()
                        # Couleur initiale selon état moteur (sera écrasée par CF)
                        if sp_section:
                            cell.fill = _fill(C['sp_dyn_bg'])
                            cell.font = _font(size=9, bold=True, color=C['agent_txt'])
                        else:
                            cell.fill = _fill(C['bureau_bg'])
                            cell.font = _font(size=8, color=C['gray'])
                    else:
                        # Fallback statique
                        if sp_section:
                            bg_sp = (C['vac_bg'] if is_vacataire(agent)
                                     else SEC_BG_AG.get(sp_section, C['jeunesse_ag']))
                            fc_sp = C['vac_txt'] if is_vacataire(agent) else C['agent_txt']
                            _set(cell, value=f"SP\n{sp_section}",
                                 bg=bg_sp, fnt=_font(size=9, bold=True, color=fc_sp),
                                 aln=_aln(), brd=_brd())
                        else:
                            _set(cell, value='',
                                 bg=C['bureau_bg'],
                                 fnt=_font(size=8, color=C['gray']),
                                 aln=_aln(), brd=_brd())

            ws.row_dimensions[row].height = 28
            row += 1

        cren_end_row = row - 1  # dernière ligne SP (inclusive)

        # ── Mise en forme conditionnelle sur la plage créneaux ───────
        # Appliquée sur toutes les colonnes jour × toutes les lignes créneaux
        # Règle 1 : cellule = "SP"  → lavande (fond SP)
        # Règle 2 : cellule = ""    → jaune (bureau)
        # Les cellules statiques (—, Congé, Évén.) ne matchent aucune règle
        if jours_presents and jour_cren_rows:
            col_start = col_map[jours_presents[0]]
            col_end   = col_map[jours_presents[-1]]
            cf_range  = (f"{get_column_letter(col_start)}{cren_start_row}:"
                         f"{get_column_letter(col_end)}{cren_end_row}")

            # Ancre = première cellule de la plage (la formule CF est relative)
            anc = f"{get_column_letter(col_start)}{cren_start_row}"

            fill_sp  = PatternFill(fill_type='solid', fgColor='E8DAEF')  # lavande
            fill_bur = PatternFill(fill_type='solid', fgColor='FEF9E7')  # jaune bureau

            font_sp  = Font(size=9, bold=True)
            font_bur = Font(size=8)

            rule_sp = FormulaRule(
                formula=[f'LEFT({anc},2)="SP"'],
                fill=fill_sp,
                font=font_sp,
                stopIfTrue=True
            )
            rule_bur = FormulaRule(
                formula=[f'{anc}=""'],
                fill=fill_bur,
                font=font_bur,
                stopIfTrue=True
            )
            ws.conditional_formatting.add(cf_range, rule_sp)
            ws.conditional_formatting.add(cf_range, rule_bur)

        # ── Ligne TOTAL SP (formule SUMPRODUCT dynamique) ──────
        # Les cellules SP seront converties en sharedStrings en post-processing
        # → SEARCH("SP", ...) fonctionnera nativement dans Excel
        _set(ws.cell(row=row, column=1), bg='F0F3F4')
        ws.cell(row=row, column=1).value = ''
        _set(ws.cell(row=row, column=2, value='TOTAL SP'),
             bg='34495E', fnt=_font(bold=True, size=9, color='FFFFFF'),
             aln=_aln(), brd=_brd())

        agent_sp_cells[agent] = {}

        # ── TOTAL SP : formule dynamique cross-sheet ──────────────
        # Cherche le nom de l'agent dans Semaine_N colonnes C(RDC),D(Adulte),E(MF),F(Jeunesse)
        # Multiplie par les durées en col A (heures décimales, police blanche)
        # Entièrement recalculé par Excel si Semaine_N est modifié

        sem_sheet = f"Semaine_{week_num}"   # nom de l'onglet source
        # Sections = colonnes C,D,E,F dans Semaine_N
        SEC_COLS_SEM = ['C', 'D', 'E', 'F']

        for jour in jours_presents:
            c          = col_map[jour]
            col_ltr    = get_column_letter(c)
            total_cell = ws.cell(row=row, column=c)
            sp_j       = sp_jour_cum.get(jour, {}).get(agent, 0)

            if week_dates.get(jour) is not None:
                # Lignes créneaux dans Semaine_N pour ce jour
                if jour_cren_rows and jour in jour_cren_rows:
                    r1, r2 = jour_cren_rows[jour]
                else:
                    # Fallback : valeur statique moteur
                    sp_h = round(sp_j / 60, 2)
                    total_cell.value = sp_h
                    total_cell.number_format = '0.00'
                    agent_sp_cells[agent][jour] = f"{col_ltr}{row}"
                    bg_t = 'D5F5E3' if sp_j > 0 else 'F0F3F4'
                    fc_t = '1E8449' if sp_j > 0 else C['gray']
                    total_cell.fill = _fill(bg_t)
                    total_cell.font = _font(bold=sp_j>0, size=9, color=fc_t)
                    total_cell.alignment = _aln()
                    total_cell.border = _brd()
                    continue

                # =SUMPRODUCT(
                #   ((ISNUMBER(SEARCH("agent",Sem!C6:C16))
                #    +ISNUMBER(SEARCH("agent",Sem!D6:D16))
                #    +ISNUMBER(SEARCH("agent",Sem!E6:E16))
                #    +ISNUMBER(SEARCH("agent",Sem!F6:F16)))>0)
                #   * A{cren_start}:A{cren_end})
                search_parts = '+'.join(
                    'ISNUMBER(SEARCH("' + agent + '",' + "'" + sem_sheet + "'!" + sc + str(r1) + ':' + sc + str(r2) + '))'
                    for sc in SEC_COLS_SEM
                )
                formula = (
                    '=SUMPRODUCT((' + search_parts + '>0)'
                    '*A' + str(cren_start_row) + ':A' + str(cren_end_row) + ')'
                )
                total_cell.value         = formula
                total_cell.number_format = '0.00'

                bg_t = 'D5F5E3' if sp_j > 0 else 'F0F3F4'
                fc_t = '1E8449' if sp_j > 0 else C['gray']
                total_cell.fill      = _fill(bg_t)
                total_cell.font      = _font(bold=sp_j > 0, size=9, color=fc_t)
                total_cell.alignment = _aln()
                total_cell.border    = _brd()

                # Stocker la référence pour le récap cross-sheet Semaine_N
                agent_sp_cells[agent][jour] = f"{col_ltr}{row}"
            else:
                _set(total_cell, value='—',
                     bg='F0F3F4',
                     fnt=_font(size=9, color=C['gray']),
                     aln=_aln(), brd=_brd())

        ws.row_dimensions[row].height = 20
        row += 1

        # ── Séparateur ──────────────────────────────────────────
        for ci in range(1, nb_cols + 1):
            ws.cell(row=row, column=ci).fill = _fill('C0C0C0')
        ws.row_dimensions[row].height = 3
        row += 2

    return ws, agent_sp_cells




# ══════════════════════════════════════════════════════════════
#  POST-PROCESSING : inlineStr → sharedStrings
# ══════════════════════════════════════════════════════════════

def _xml_decode(s):
    """Decode XML entities (named + numeric) via html.unescape + newline."""
    import html
    # html.unescape gère &amp; &lt; &gt; &quot; + toutes les entites numeriques &#NNN; &#xHHH;
    # On remplace d'abord &#10; par un placeholder pour le newline
    return html.unescape(s)


def _xml_encode(s):
    """Encode string for XML."""
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;')
             .replace('\n', '&#10;'))


def _convert_inlinestr_to_shared_strings(buf, pa_sheet_names, all_sp_cells=None, weeks_data=None):
    """
    Convertit les cellules inlineStr des onglets Planning_Agent en sharedStrings.
    openpyxl ecrit tout en inlineStr ; SEARCH() ne fonctionne pas sur ce format.
    Apres conversion, les formules SUMPRODUCT+SEARCH deviennent dynamiques.

    Transformation :
      <c r="C8" s="42" t="inlineStr"><is><t>SP&#10;Jeunesse</t></is></c>
      => <c r="C8" s="42" t="s"><v>17</v></c>
    """
    buf.seek(0)
    in_bytes = buf.read()

    # Lire les manifestes
    with zipfile.ZipFile(io.BytesIO(in_bytes), 'r') as zin:
        wb_xml   = zin.read('xl/workbook.xml').decode()
        rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode()
        ct_xml   = zin.read('[Content_Types].xml').decode()
        all_files = zin.namelist()

    # Mapping sheet name -> chemin ZIP
    sheet_names = re.findall(r'<sheet [^>]*name="([^"]+)"', wb_xml)
    sheet_rids  = re.findall(r'<sheet [^>]*r:id="([^"]+)"', wb_xml)
    rel_ids     = re.findall(r'<Relationship[^>]*\sId="([^"]+)"', rels_xml)
    rel_targets = re.findall(r'<Relationship[^>]*\sTarget="([^"]+)"', rels_xml)
    rid_to_target = dict(zip(rel_ids, rel_targets))
    name_to_path = {}
    for sname, rid in zip(sheet_names, sheet_rids):
        target = rid_to_target.get(rid, '')
        if target:
            name_to_path[sname] = target.lstrip('/')

    pa_paths = {sn: name_to_path[sn] for sn in pa_sheet_names if sn in name_to_path}

    # Construire table shared strings (partir de l'existant si present)
    str_table = []
    str_index = {}

    if 'xl/sharedStrings.xml' in all_files:
        with zipfile.ZipFile(io.BytesIO(in_bytes), 'r') as zin:
            ss_xml = zin.read('xl/sharedStrings.xml').decode()
        for m in re.finditer(r'<si><t[^>]*>(.*?)</t></si>', ss_xml, re.DOTALL):
            txt = _xml_decode(m.group(1))
            if txt not in str_index:
                str_index[txt] = len(str_table)
                str_table.append(txt)

    # Collecter tous les inlineStr des feuilles PA
    with zipfile.ZipFile(io.BytesIO(in_bytes), 'r') as zin:
        for sname, path in pa_paths.items():
            if path not in all_files:
                continue
            xml = zin.read(path).decode()
            for m in re.finditer(
                    r'<c [^>]*t="inlineStr"[^>]*><is><t[^>]*>(.*?)</t></is></c>',
                    xml, re.DOTALL):
                txt = _xml_decode(m.group(1))
                if txt not in str_index:
                    str_index[txt] = len(str_table)
                    str_table.append(txt)

    # Construire sharedStrings.xml
    NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    si_parts = []
    for s in str_table:
        s_enc = _xml_encode(s)
        space = ' xml:space="preserve"' if (' ' in s or '\n' in s) else ''
        si_parts.append('<si><t' + space + '>' + s_enc + '</t></si>')
    n = len(str_table)
    new_ss_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  '<sst ' + NS + ' count="' + str(n) + '" uniqueCount="' + str(n) + '">'
                  + ''.join(si_parts) + '</sst>')

    # Convertir les cellules inlineStr -> t="s" dans chaque feuille PA
    def convert_sheet_xml(xml):
        def repl(m):
            c_open = m.group(1)   # ex: <c r="C8" s="42" t="inlineStr">
            txt    = _xml_decode(m.group(2))
            idx    = str_index.get(txt, 0)
            # Retirer t="inlineStr" ET le > final pour reconstruire proprement
            c_open = re.sub(r'\s*t="inlineStr"', '', c_open)
            c_open = c_open.rstrip('>').rstrip()
            return c_open + ' t="s"><v>' + str(idx) + '</v></c>'
        return re.sub(
            r'(<c [^>]*t="inlineStr"[^>]*>)<is><t[^>]*>(.*?)</t></is></c>',
            repl, xml, flags=re.DOTALL)

    # Mettre a jour Content_Types et rels si necessaire
    SS_CT  = ('application/vnd.openxmlformats-officedocument'
              '.spreadsheetml.sharedStrings+xml')
    SS_REL = ('http://schemas.openxmlformats.org/officeDocument'
              '/2006/relationships/sharedStrings')

    if 'sharedStrings' not in ct_xml:
        ct_xml = ct_xml.replace(
            '</Types>',
            '<Override PartName="/xl/sharedStrings.xml"'
            ' ContentType="' + SS_CT + '"/></Types>')

    if 'sharedStrings' not in rels_xml:
        max_rid = max((int(r) for r in re.findall(r'Id="rId(\d+)"', rels_xml)),
                      default=0)
        new_rid = 'rId' + str(max_rid + 1)
        rels_xml = rels_xml.replace(
            '</Relationships>',
            '<Relationship Id="' + new_rid + '" Type="' + SS_REL + '"'
            ' Target="sharedStrings.xml"/></Relationships>')

    # Construire mapping cellule -> valeur calculee pour les formules TOTAL SP
    # Cela permet a Excel d'afficher les valeurs immediatement a l'ouverture
    # (sans attendre un recalcul manuel), tout en gardant la formule dynamique
    cached_vals = {}  # {pa_sheet_name: {cell_ref: float}}
    if all_sp_cells and weeks_data:
        for wk in weeks_data:
            wn  = wk['week_num']
            sn  = f"Planning_Agent_Semaine_{wn}"
            sp_jour_cum = wk.get('sp_jour_cum', {})
            asc = all_sp_cells.get(wn, {})
            cached_vals[sn] = {}
            for agent, jour_map in asc.items():
                for jour, cell_ref in jour_map.items():
                    val = round(sp_jour_cum.get(jour, {}).get(agent, 0) / 60, 4)
                    cached_vals[sn][cell_ref] = val

    def inject_cached(xml, sheet_name):
        """Injecte les valeurs calculees dans <v></v> des cellules formule."""
        vals = cached_vals.get(sheet_name, {})
        for cell_ref, val in vals.items():
            vs  = f"{val:.4f}".rstrip('0').rstrip('.') or '0'
            esc = re.escape(cell_ref)
            xml = re.sub(
                r'(<c r="' + esc + r'"[^>]*><f>[^<]*</f>)<v>[^<]*</v>(</c>)',
                r'\g<1><v>' + vs + r'</v>\g<2>',
                xml
            )
        return xml

    # Reecrire le ZIP
    out_buf = io.BytesIO()
    ss_written = False
    with zipfile.ZipFile(io.BytesIO(in_bytes), 'r') as zin, \
         zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = ct_xml.encode('utf-8')
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                data = rels_xml.encode('utf-8')
            elif item.filename == 'xl/sharedStrings.xml':
                data = new_ss_xml.encode('utf-8')
                ss_written = True
            else:
                for sname, path in pa_paths.items():
                    if item.filename == path:
                        xml_str = convert_sheet_xml(data.decode('utf-8'))
                        xml_str = inject_cached(xml_str, sname)
                        data = xml_str.encode('utf-8')
                        break
            zout.writestr(item, data)

        # Ajouter sharedStrings.xml s'il n'existait pas
        if not ss_written:
            zout.writestr('xl/sharedStrings.xml', new_ss_xml.encode('utf-8'))

    out_buf.seek(0)
    return out_buf


def _patch_recap_formulas(wb, week_data, metadata, agent_sp_cells):
    """
    Passe 3 : réécrit les cellules du récap SP dans Semaine_N
    avec des formules cross-sheet pointant vers Planning_Agent_Semaine_N.
    Ex : C77 = ='Planning_Agent_Semaine_1'!C19
         H77 = =SUM(C77:G77)
    """
    week_num     = week_data['week_num']
    week_dates   = week_data['week_dates']
    sp_jour_cum  = week_data.get('sp_jour_cum', {})
    sp_alerts    = week_data.get('sp_alerts', {})
    affectations = metadata.get('affectations', {})
    sp_minmax_all  = metadata.get('sp_minmax_all', {})
    sp_minmax_week = sp_minmax_all.get(week_num, sp_minmax_all.get(2, {}))
    is_sam_week  = 'Samedi' in week_dates

    sname = f"Semaine_{week_num}"
    pa_sheet = f"Planning_Agent_Semaine_{week_num}"
    if sname not in wb.sheetnames:
        return
    ws = wb[sname]

    # Trouver les lignes du récap (chercher l'en-tête "Agent")
    recap_hdr_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == 'Agent':
            recap_hdr_row = r
            break
    if recap_hdr_row is None:
        return

    # Agents dans le même ordre que write_week_sheet
    def primary_section(agent):
        sects = affectations.get(agent, [])
        return sects[0] if sects else '?'

    agents_shown = sorted(
        [a for a in sp_minmax_week.keys() if not is_vacataire(a)],
        key=lambda x: (primary_section(x), x)
    ) + [a for a in affectations.keys() if is_vacataire(a)]

    data_start_row = recap_hdr_row + 1

    for i, agent in enumerate(agents_shown):
        r   = data_start_row + i
        bg  = C['alt1'] if i % 2 == 0 else C['alt2']
        sp_total_static = 0.0

        for ji, jour in enumerate(JOURS_SP):
            col  = ji + 3   # C=3, D=4, E=5, F=6, G=7
            cell = ws.cell(row=r, column=col)
            sp_min = sp_jour_cum.get(jour, {}).get(agent, 0)
            sp_h   = min_to_dec(sp_min)
            sp_total_static += sp_h

            if week_dates.get(jour) is not None:
                cell_ref = (agent_sp_cells.get(agent, {}).get(jour)
                            if agent_sp_cells else None)
                if cell_ref:
                    cell.value = f"='{pa_sheet}'!{cell_ref}"
                else:
                    cell.value = sp_h if sp_h > 0 else 0
                cell.number_format = '0.00'

                bg_c = ('FFF9C4' if jour == 'Samedi' and sp_h > 0
                         else C['alt2'] if sp_h > 0 else bg)
                cell.fill      = _fill(bg_c)
                cell.font      = _font(size=10, bold=sp_h > 0)
                cell.alignment = _aln()
                cell.border    = _brd()
            # Jours absents : laisser tel quel (déjà '—')

        # Colonne H : SUM dynamique des colonnes jour présents
        tc = ws.cell(row=r, column=8)
        sum_cols = [get_column_letter(3 + ji)
                    for ji, jour in enumerate(JOURS_SP)
                    if week_dates.get(jour) is not None]
        if sum_cols:
            tc.value = f"=SUM({','.join(c + str(r) for c in sum_cols)})"
        tc.number_format = '0.00'

        # Couleur colonne H (basée sur valeur statique moteur)
        total_h = sp_total_static
        if not is_vacataire(agent):
            mm     = sp_minmax_week.get(agent, {})
            min_sp = mm.get('Min_MarSam', 0) if is_sam_week else mm.get('Min_MarVen', 0)
            max_sp = mm.get('Max_MarSam', 99) if is_sam_week else mm.get('Max_MarVen', 99)
            if agent in sp_alerts:
                t_bg, t_fc = C['sp_alert'], C['sp_alert_t']
            elif total_h < min_sp:
                t_bg, t_fc = C['sp_alert'], C['sp_alert_t']
            elif total_h > max_sp:
                t_bg, t_fc = C['sp_warn'], C['sp_warn_t']
            else:
                t_bg, t_fc = C['sp_ok'], C['sp_ok_t']
        else:
            t_bg = C['vac_bg'] if total_h > 0 else bg
            t_fc = C['vac_txt'] if total_h > 0 else C['gray']

        tc.fill      = _fill(t_bg)
        tc.font      = _font(bold=True, size=10, color=t_fc)
        tc.alignment = _aln()
        tc.border    = _brd()


# ══════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════

def generate_excel(source_filepath, weeks_data, metadata):
    wb = load_workbook(source_filepath)

    # Supprimer anciens onglets générés
    to_remove = [s for s in wb.sheetnames
                 if s.startswith('Semaine_') or s.startswith('Planning_Agent')]
    for name in to_remove:
        wb.remove(wb[name])

    # Passe 1 : onglets Semaine → récupère les numéros de lignes créneaux
    all_jour_cren_rows = {}
    for week_data in weeks_data:
        _ws, jcr = write_week_sheet(wb, week_data, metadata)
        all_jour_cren_rows[week_data['week_num']] = jcr

    # Passe 2 : onglets Planning_Agent → formules cross-sheet vers Semaine
    #           récupère agent_sp_cells = {agent: {jour: 'C19'}}
    all_sp_cells = {}
    for week_data in weeks_data:
        jcr = all_jour_cren_rows.get(week_data['week_num'], {})
        _, asc = write_planning_agent_week_sheet(wb, week_data, metadata, jour_cren_rows=jcr)
        all_sp_cells[week_data['week_num']] = asc

    # Passe 3 : injecter les formules cross-sheet dans le récap de chaque Semaine
    #           maintenant qu'on connaît les refs Planning_Agent (C19, C37, ...)
    for week_data in weeks_data:
        asc = all_sp_cells.get(week_data['week_num'], {})
        _patch_recap_formulas(wb, week_data, metadata, asc)

    # Réordonner : Semaine_1..5 → Planning_Agent_1..5 → autres
    week_sheets  = [f"Semaine_{i}" for i in range(1, len(weeks_data)+1)
                    if f"Semaine_{i}" in wb.sheetnames]
    agent_sheets = [f"Planning_Agent_Semaine_{i}" for i in range(1, len(weeks_data)+1)
                    if f"Planning_Agent_Semaine_{i}" in wb.sheetnames]
    other_sheets = [s for s in wb.sheetnames
                    if s not in week_sheets and s not in agent_sheets]
    wb._sheets   = ([wb[s] for s in week_sheets]
                    + [wb[s] for s in agent_sheets]
                    + [wb[s] for s in other_sheets])


    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    buf = io.BytesIO()
    wb.save(buf)

    # Note : _convert_inlinestr_to_shared_strings n'est plus appelée.
    # Les formules cross-sheet (IF/SEARCH vers Semaine_N) fonctionnent
    # nativement sans conversion inlineStr — les cellules Semaine_N
    # sont des sharedStrings dès que Excel recalcule le fichier.

    buf.seek(0)
    return buf
