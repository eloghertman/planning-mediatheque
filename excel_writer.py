"""
Excel Writer — génère le fichier de planning formaté
"""

import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict


# ─────────────────────────────────────────────
#  PALETTE COULEURS
# ─────────────────────────────────────────────

C = {
    'header_dark':   '1A2F4A',
    'header_rouge':  'C0392B',
    'header_bleu':   '2471A3',
    'sec_rdc':       'D6E8F7',
    'sec_adulte':    'D4EDD4',
    'sec_mf':        'FFF0CC',
    'sec_jeunesse':  'FFE0E0',
    'closed':        'EBEBEB',
    'event_bg':      'FFF3CD',
    'event_text':    '7D6608',
    'title_txt':     'FFFFFF',
    'agent_txt':     '2C3E50',
    'gray_txt':      'AAAAAA',
    'row_alt1':      'FDFEFE',
    'row_alt2':      'EBF5FB',
    'col_header':    '2E5F8A',
    'orange':        'E67E22',
    'purple':        '8E44AD',
    'green':         '27AE60',
    'red':           'E74C3C',
    'blue':          '2980B9',
    'dark':          '2C3E50',
}

SEC_COLORS = {
    'RDC':      C['sec_rdc'],
    'Adulte':   C['sec_adulte'],
    'MF':       C['sec_mf'],
    'Jeunesse': C['sec_jeunesse'],
}

SEC_HEADER_COLORS = {
    'RDC':      C['blue'],
    'Adulte':   C['green'],
    'MF':       C['orange'],
    'Jeunesse': C['red'],
}

def _fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def _font(bold=False, size=10, color='000000', italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name='Calibri')

def _aln(h='center', v='center', wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _brd(color='AAAAAA', style='thin'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _brd_medium():
    s = Side(style='medium', color='555555')
    return Border(left=s, right=s, top=s, bottom=s)

def _set(cell, value=None, bg=None, fnt=None, aln=None, brd=None, height=None):
    if value is not None:
        cell.value = value
    if bg:
        cell.fill = _fill(bg)
    if fnt:
        cell.font = fnt
    if aln:
        cell.alignment = aln
    if brd:
        cell.border = brd


# ─────────────────────────────────────────────
#  ÉCRITURE D'UN ONGLET SEMAINE
# ─────────────────────────────────────────────

JOURS_SP = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
SECTIONS = ['RDC', 'Adulte', 'MF', 'Jeunesse']

MOIS_FR = {
    1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
    7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'
}

def write_week_sheet(wb, week_data, metadata, sheet_name=None):
    """Écrit un onglet semaine dans le classeur wb."""
    week_num   = week_data['week_num']
    week_dates = week_data['week_dates']
    plan       = week_data['plan']
    creneaux   = metadata['creneaux']
    mois       = metadata['mois'].capitalize()
    annee      = metadata['annee']

    # Dates de début et fin de semaine
    dates_list = [d for d in week_dates.values() if d is not None]
    if dates_list:
        d_start = min(dates_list)
        d_end   = max(dates_list)
        date_range = f"{d_start.day} au {d_end.day} {mois} {annee}"
    else:
        date_range = f"Semaine {week_num}"

    sname = sheet_name or f"Semaine_{week_num}"
    if sname in wb.sheetnames:
        del wb[sname]
    ws = wb.create_sheet(sname)

    # Colonnes : A=N°, B=Créneau, C=RDC, D=Adulte, E=MF, F=Jeunesse, G=Événement, H=Agents événement
    widths = {'A':5, 'B':14, 'C':22, 'D':22, 'E':22, 'F':26, 'G':34, 'H':34}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    row = 1

    # ── TITRE ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws.cell(row=row, column=1,
        value=f"📅  PLANNING SERVICE PUBLIC — Semaine {week_num}  |  {date_range}")
    _set(c, bg=C['header_dark'],
         fnt=_font(bold=True, size=13, color='FFFFFF'),
         aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    # ── LÉGENDE ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws.cell(row=row, column=1,
        value="  🔵 RDC     🟢 Adulte     🟡 Musique & Films     🔴 Jeunesse     ⚫ Fermé     🟠 Événement")
    _set(c, bg='2C3E50',
         fnt=_font(size=9, color='FFFFFF', italic=True),
         aln=_aln('left'))
    ws.row_dimensions[row].height = 18
    row += 2

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        jour_plan = plan.get(jour, {})
        samedi_type = jour_plan.get('_samedi_type')

        # ── EN-TÊTE JOUR ──
        is_sam = jour == 'Samedi'
        sam_label = f"  🔴 SAMEDI ROUGE" if (is_sam and samedi_type == 'ROUGE') else \
                    f"  🔵 SAMEDI BLEU"  if (is_sam and samedi_type == 'BLEU')  else ""
        jour_color = C['header_rouge'] if (is_sam and samedi_type == 'ROUGE') else \
                     C['header_bleu']  if (is_sam and samedi_type == 'BLEU')  else \
                     C['header_dark']
        date_label = f"{jour.upper()}  {date.day} {mois} {annee}{sam_label}"
        ws.merge_cells(f'A{row}:H{row}')
        c = ws.cell(row=row, column=1, value=f"  {date_label}")
        _set(c, bg=jour_color,
             fnt=_font(bold=True, size=12, color='FFFFFF'),
             aln=_aln('left'))
        ws.row_dimensions[row].height = 26
        row += 1

        # ── EN-TÊTES COLONNES ──
        headers = [
            ('N°',     C['dark']),
            ('Créneau',C['dark']),
            ('RDC',    SEC_HEADER_COLORS['RDC']),
            ('Adulte', SEC_HEADER_COLORS['Adulte']),
            ('Musique & Films', SEC_HEADER_COLORS['MF']),
            ('Jeunesse',        SEC_HEADER_COLORS['Jeunesse']),
            ('Événement',       C['purple']),
            ('Agents concernés',C['purple']),
        ]
        for ci, (h, hc) in enumerate(headers):
            c = ws.cell(row=row, column=ci+1, value=h)
            _set(c, bg=hc,
                 fnt=_font(bold=True, size=10, color='FFFFFF'),
                 aln=_aln('center'),
                 brd=_brd())
        ws.row_dimensions[row].height = 20
        row += 1

        # ── LIGNES CRÉNEAUX ──
        for ni, (cren_name, cs, ce) in enumerate(creneaux):
            slot = jour_plan.get(cren_name)

            if slot is None:
                # Fermé
                c1 = ws.cell(row=row, column=1, value=ni+1)
                _set(c1, bg=C['closed'], fnt=_font(size=9, color=C['gray_txt']), aln=_aln(), brd=_brd())

                c2 = ws.cell(row=row, column=2, value=cren_name)
                _set(c2, bg=C['closed'], fnt=_font(size=10, color=C['gray_txt'], italic=True), aln=_aln(), brd=_brd())

                for ci in range(3, 7):
                    cx = ws.cell(row=row, column=ci, value='— Fermé —')
                    _set(cx, bg=C['closed'], fnt=_font(size=9, color=C['gray_txt'], italic=True), aln=_aln(), brd=_brd())

                for ci in [7, 8]:
                    cx = ws.cell(row=row, column=ci, value='')
                    _set(cx, bg='FAFAFA', brd=_brd())

            else:
                assignment = slot.get('assignment', {s: [] for s in SECTIONS})
                events     = slot.get('events', [])

                # N°
                c1 = ws.cell(row=row, column=1, value=ni+1)
                _set(c1, bg='F2F3F4', fnt=_font(size=9, color='7F8C8D'), aln=_aln(), brd=_brd())

                # Créneau
                c2 = ws.cell(row=row, column=2, value=cren_name)
                _set(c2, bg='FDFEFE', fnt=_font(bold=True, size=10, color=C['agent_txt']), aln=_aln(), brd=_brd())

                # Sections
                for si, sect in enumerate(SECTIONS):
                    col = si + 3
                    agents = assignment.get(sect, [])
                    val = '  /  '.join(agents) if agents else '—'
                    bg  = SEC_COLORS[sect] if agents else 'FDF2F8'
                    cx  = ws.cell(row=row, column=col, value=val)
                    _set(cx, bg=bg,
                         fnt=_font(size=10, color=C['agent_txt'] if agents else 'C0392B',
                                   bold=bool(agents)),
                         aln=_aln('center'), brd=_brd())

                # Événements
                if events:
                    ev_names  = '\n'.join(ev['nom'] for ev in events)
                    ev_agents = '\n'.join(', '.join(ev['agents']) for ev in events)
                    n_evts = len(events)
                else:
                    ev_names, ev_agents = '', ''
                    n_evts = 0

                c7 = ws.cell(row=row, column=7, value=ev_names)
                _set(c7,
                     bg=C['event_bg'] if ev_names else 'FAFAFA',
                     fnt=_font(size=9, bold=bool(ev_names),
                               color=C['event_text'] if ev_names else C['gray_txt']),
                     aln=_aln('left'), brd=_brd())

                c8 = ws.cell(row=row, column=8, value=ev_agents)
                _set(c8,
                     bg=C['event_bg'] if ev_agents else 'FAFAFA',
                     fnt=_font(size=9, italic=True,
                               color='6E2F09' if ev_agents else C['gray_txt']),
                     aln=_aln('left'), brd=_brd())

                if n_evts > 1:
                    ws.row_dimensions[row].height = max(20, 16 * n_evts)
                else:
                    ws.row_dimensions[row].height = 20

            row += 1

        # Séparateur entre jours
        for ci in range(1, 9):
            cx = ws.cell(row=row, column=ci, value='')
            cx.fill = _fill('F0F3F4')
        ws.row_dimensions[row].height = 6
        row += 1

    # ── RÉCAP SP ──
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    c = ws.cell(row=row, column=1, value="📊  RÉCAPITULATIF — Heures SP par agent cette semaine")
    _set(c, bg='34495E', fnt=_font(bold=True, size=11, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 24
    row += 1

    sp_count = week_data.get('sp_count', {})
    affectations = metadata.get('affectations', {})

    agent_primary = {}
    for agent, sects in affectations.items():
        agent_primary[agent] = sects[0] if sects else '?'

    # En-têtes récap
    recap_headers = ['Agent', 'Section', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Total semaine']
    rh_colors = ['34495E','34495E','1A5276','1A5276','1A5276','1A5276','922B21','145A32']
    for ci, (h, hc) in enumerate(zip(recap_headers, rh_colors)):
        c = ws.cell(row=row, column=ci+1, value=h)
        _set(c, bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'), aln=_aln(), brd=_brd())
    ws.row_dimensions[row].height = 20
    row += 1

    # Calcul par agent et par jour
    creneaux_mins = {cn: (ce - cs) for cn, cs, ce in creneaux}
    agent_day_min = defaultdict(lambda: defaultdict(int))

    for jour in JOURS_SP:
        jour_plan = plan.get(jour, {})
        for cren_name, cs, ce in creneaux:
            slot = jour_plan.get(cren_name)
            if not slot or slot is None:
                continue
            mins = creneaux_mins.get(cren_name, 60)
            assignment = slot.get('assignment', {})
            for sect, agents in assignment.items():
                for agent in agents:
                    agent_day_min[agent][jour] += mins

    agents_sorted = sorted(
        [a for a in affectations if any(agent_day_min[a].values()) or sp_count.get(a, 0) > 0],
        key=lambda x: (agent_primary.get(x, 'ZZ'), x)
    )

    for i, agent in enumerate(agents_sorted):
        bg = C['row_alt1'] if i % 2 == 0 else C['row_alt2']
        sect = agent_primary.get(agent, '')

        c = ws.cell(row=row, column=1, value=agent)
        _set(c, bg=bg, fnt=_font(bold=True, size=10), aln=_aln('left'), brd=_brd())

        c = ws.cell(row=row, column=2, value=sect)
        _set(c, bg=SEC_COLORS.get(sect, bg), fnt=_font(size=10), aln=_aln(), brd=_brd())

        total = 0
        for di, jour in enumerate(JOURS_SP):
            mins = agent_day_min[agent].get(jour, 0)
            total += mins
            val = f"{mins//60}h{mins%60:02d}" if mins > 0 else '—'
            bg_c = 'FFF9C4' if (jour == 'Samedi' and mins > 0) else (C['row_alt2'] if mins > 0 else bg)
            c = ws.cell(row=row, column=di+3, value=val)
            _set(c, bg=bg_c, fnt=_font(size=10, bold=mins>0), aln=_aln(), brd=_brd())

        c = ws.cell(row=row, column=8, value=f"{total//60}h{total%60:02d}" if total > 0 else '0h00')
        _set(c, bg='D5F5E3' if total > 0 else bg, fnt=_font(bold=True, size=10, color='1E8449'), aln=_aln(), brd=_brd())

        ws.row_dimensions[row].height = 18
        row += 1

    return ws


# ─────────────────────────────────────────────
#  GÉNÉRATION DU FICHIER COMPLET
# ─────────────────────────────────────────────

def generate_excel(source_filepath, weeks_data, metadata):
    """
    Charge le fichier source (pour garder toutes les feuilles de paramètres),
    remplace les onglets Semaine_1..4 par les plannings calculés,
    retourne un buffer bytes.
    """
    wb = load_workbook(source_filepath)

    # Supprimer les anciens onglets semaine
    for i in range(1, 6):
        name = f"Semaine_{i}"
        if name in wb.sheetnames:
            del wb[name]

    # Écrire les nouveaux onglets dans l'ordre
    for week_data in weeks_data:
        write_week_sheet(wb, week_data, metadata)

    # Remettre les onglets semaine en premier
    sheet_order = [f"Semaine_{i}" for i in range(1, len(weeks_data)+1)]
    other_sheets = [s for s in wb.sheetnames if s not in sheet_order]
    wb._sheets = [wb[s] for s in sheet_order if s in wb.sheetnames] + \
                 [wb[s] for s in other_sheets if s in wb.sheetnames]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
