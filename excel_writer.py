"""
Excel Writer v10 — Médiathèque
════════════════════════════════════════════════════════════════
Onglets générés :
  - Semaine_1 à Semaine_N  : planning hebdomadaire + récap SP
  - Planning_Agent_Semaine_1 à _N : vue par agent, un onglet par semaine
════════════════════════════════════════════════════════════════
"""

import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

SECTIONS = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SP = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
MOIS_FR  = {1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
            7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'}

C = {
    'hdr_dark':  '1A2F4A', 'hdr_rouge': 'C0392B', 'hdr_bleu':  '2471A3',
    'rdc':       'D6E8F7', 'adulte':    'D4EDD4', 'mf':        'FFF0CC',
    'jeunesse':  'FFE0E0', 'closed':    'EBEBEB', 'event_bg':  'FFF3CD',
    'alert_bg':  'FFCCCC', 'alert_txt': '8B0000', 'agent_txt': '2C3E50',
    'gray':      'AAAAAA', 'alt1':      'FDFEFE', 'alt2':      'EBF5FB',
    'dark':      '2C3E50', 'sp_ok':     'D5F5E3', 'sp_warn':   'FFF3CD',
    'sp_alert':  'FFCCCC', 'sp_ok_t':   '1E8449', 'sp_warn_t': 'E67E22',
    'sp_alert_t':'8B0000',
    'hdr_rdc':   '2980B9', 'hdr_adulte':'27AE60', 'hdr_mf':    'E67E22',
    'hdr_j':     'E74C3C', 'hdr_purple':'8E44AD', 'hdr_dark2': '34495E',
    'oos_bg':    'FF0000', 'oos_txt':   'FFFFFF',  # D16 hors section
    'vac_bg':    'F0E6FF', 'vac_txt':   '6A0DAD',  # vacataire
}
SEC_BG  = {'RDC': C['rdc'],      'Adulte': C['adulte'], 'MF': C['mf'],      'Jeunesse': C['jeunesse']}
SEC_HDR = {'RDC': C['hdr_rdc'],  'Adulte': C['hdr_adulte'], 'MF': C['hdr_mf'], 'Jeunesse': C['hdr_j']}


def _fill(h): return PatternFill('solid', fgColor=h)
def _font(bold=False, size=10, color='000000', italic=False, name='Calibri'):
    return Font(bold=bold, size=size, color=color, italic=italic, name=name)
def _aln(h='center', v='center', wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _brd(color='CCCCCC', style='thin'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def _set(cell, value=None, bg=None, fnt=None, aln=None, brd=None):
    if value is not None: cell.value = value
    if bg:  cell.fill      = _fill(bg)
    if fnt: cell.font      = fnt
    if aln: cell.alignment = aln
    if brd: cell.border    = brd

def is_vacataire(agent): return 'vacataire' in str(agent).lower()
def min_to_hhmm(m): return f"{int(m)//60}h{int(m)%60:02d}" if m else '0h00'
def min_to_dec(m): return round(m / 60, 2) if m else 0.0


# ══════════════════════════════════════════════════════════════
#  ONGLET SEMAINE
# ══════════════════════════════════════════════════════════════

def write_week_sheet(wb, week_data, metadata):
    week_num    = week_data['week_num']
    week_dates  = week_data['week_dates']
    plan        = week_data['plan']
    sp_count    = week_data['sp_count']
    sp_jour_cum = week_data.get('sp_jour_cum', {})
    sp_alerts   = week_data.get('sp_alerts', {})
    creneaux    = metadata['creneaux']
    sp_minmax_all  = metadata.get('sp_minmax_all', {})
    sp_minmax_week = sp_minmax_all.get(week_num, sp_minmax_all.get(2, {}))
    affectations   = metadata.get('affectations', {})
    is_sam_week    = 'Samedi' in week_dates

    mois_num = {'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
                'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,'décembre':12
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
    if sname in wb.sheetnames: wb.remove(wb[sname])
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False

    for col, w in {'A':5,'B':14,'C':20,'D':20,'E':20,'F':26,'G':28,'H':30}.items():
        ws.column_dimensions[col].width = w

    row = 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value=f"PLANNING SERVICE PUBLIC — Semaine {week_num}  |  {date_range}  |  {sem_type}"),
         bg=C['hdr_dark'], fnt=_font(bold=True, size=13, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="  RDC     Adulte     Musique & Films     Jeunesse     Fermé     Événement     ALERTE     Rouge=hors section"),
         bg=C['dark'], fnt=_font(size=9, color='FFFFFF', italic=True), aln=_aln('left'))
    ws.row_dimensions[row].height = 15
    row += 2

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None: continue
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
                ('Jeunesse',SEC_HDR['Jeunesse']),('Événement',C['hdr_purple']),('Alertes',C['hdr_rouge'])]
        for ci,(h_txt,hc) in enumerate(hdrs):
            _set(ws.cell(row=row, column=ci+1, value=h_txt),
                 bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'),
                 aln=_aln('center'), brd=_brd())
        ws.row_dimensions[row].height = 20
        row += 1

        for ni,(cren_name,cs,ce) in enumerate(creneaux):
            slot = jour_plan.get(cren_name)
            if slot is None:
                for ci,val in enumerate([ni+1, cren_name,'—','—','—','—','',''],1):
                    _set(ws.cell(row=row, column=ci, value=val or None),
                         bg=C['closed'], fnt=_font(size=9 if ci>2 else 10, color=C['gray'], italic=ci>1),
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

                for si,sect in enumerate(SECTIONS):
                    col = si + 3
                    agents = assgn.get(sect, [])
                    has_al = any(sect in al for al in alerts)
                    is_oos = oos.get(sect, False)

                    if is_oos and agents:
                        val = '⚠ ' + '  /  '.join(agents)
                        bg  = C['oos_bg']
                        fc  = C['oos_txt']
                    elif agents:
                        val = '  /  '.join(agents)
                        bg  = SEC_BG[sect]
                        fc  = C['vac_txt'] if all(is_vacataire(a) for a in agents) else C['agent_txt']
                    elif has_al:
                        val, bg, fc = 'ALERTE', C['alert_bg'], C['alert_txt']
                    else:
                        val, bg, fc = '—', 'FDF2F8', 'C0392B'

                    _set(ws.cell(row=row, column=col, value=val),
                         bg=bg, fnt=_font(size=10, color=fc, bold=bool(agents)),
                         aln=_aln('center'), brd=_brd())

                ev_txt = '\n'.join(f"{ev['nom']} ({', '.join(ev['agents'][:3])})"
                                    for ev in events) if events else None
                _set(ws.cell(row=row, column=7, value=ev_txt),
                     bg=C['event_bg'] if ev_txt else 'FAFAFA',
                     fnt=_font(size=9, italic=True, color='6E2F09' if ev_txt else C['gray']),
                     aln=_aln('left'), brd=_brd())

                al_txt = '\n'.join(alerts) if alerts else None
                _set(ws.cell(row=row, column=8, value=al_txt),
                     bg=C['alert_bg'] if al_txt else 'FAFAFA',
                     fnt=_font(size=9, bold=bool(al_txt),
                               color=C['alert_txt'] if al_txt else C['gray']),
                     aln=_aln('left'), brd=_brd())

                h = max(20, 16*max(len(events), len(alerts), 1))
                ws.row_dimensions[row].height = h
            row += 1

        for ci in range(1, 9):
            ws.cell(row=row, column=ci).fill = _fill('F0F3F4')
        ws.row_dimensions[row].height = 4
        row += 1

    # ── Récap SP ──
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

    # Réguliers + vacataires
    agents_shown = sorted(
        [a for a in sp_minmax_week.keys() if not is_vacataire(a)],
        key=lambda x: (primary_section(x), x)
    ) + [a for a in affectations.keys() if is_vacataire(a)]

    for i, agent in enumerate(agents_shown):
        bg   = C['alt1'] if i % 2 == 0 else C['alt2']
        sect = primary_section(agent)

        _set(ws.cell(row=row, column=1, value=agent),
             bg=bg, fnt=_font(bold=True, size=10), aln=_aln('left'), brd=_brd())
        _set(ws.cell(row=row, column=2, value='VAC' if is_vacataire(agent) else sect),
             bg=C['vac_bg'] if is_vacataire(agent) else SEC_BG.get(sect, bg),
             fnt=_font(size=10, color=C['vac_txt'] if is_vacataire(agent) else '000000'),
             aln=_aln(), brd=_brd())

        total_h = 0.0
        for ji,jour in enumerate(JOURS_SP):
            col    = ji + 3
            cell   = ws.cell(row=row, column=col)
            sp_min = sp_jour_cum.get(jour, {}).get(agent, 0)
            sp_h   = min_to_dec(sp_min)
            if week_dates.get(jour) is not None:
                cell.value         = sp_h if sp_h > 0 else 0
                cell.number_format = '0.00'
                bg_c = ('FFF9C4' if jour == 'Samedi' and sp_h > 0
                         else C['alt2'] if sp_h > 0 else bg)
                cell.fill      = _fill(bg_c)
                cell.font      = _font(size=10, bold=sp_h > 0)
                cell.alignment = _aln()
                cell.border    = _brd()
                total_h += sp_h
            else:
                cell.value = '—'; cell.fill=_fill(bg)
                cell.font=_font(size=9,color=C['gray'])
                cell.alignment=_aln(); cell.border=_brd()

        tc = ws.cell(row=row, column=8)
        tc.value = round(total_h, 2)
        tc.number_format = '0.00'

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

        tc.fill = _fill(t_bg); tc.font = _font(bold=True, size=10, color=t_fc)
        tc.alignment = _aln(); tc.border = _brd()
        ws.row_dimensions[row].height = 18
        row += 1

    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="Vert=OK  |  Rouge=sous min  |  Orange=sur max  |  "
                       "Violet=vacataire  |  Rouge vif cellule planning=agent hors section habituelle"),
         bg='F4F6F7', fnt=_font(size=9, italic=True, color='5D6D7E'), aln=_aln('left'))
    ws.row_dimensions[row].height = 16
    return ws


# ══════════════════════════════════════════════════════════════
#  ONGLET PLANNING PAR AGENT — 1 PAR SEMAINE
# ══════════════════════════════════════════════════════════════

def write_planning_agent_week_sheet(wb, week_data, metadata):
    """
    Un onglet Planning_Agent_Semaine_N par semaine.
    Tous les agents sur l'onglet, créneaux en lignes, jours en colonnes.
    """
    week_num    = week_data['week_num']
    week_dates  = week_data['week_dates']
    plan        = week_data['plan']
    sp_jour_cum = week_data.get('sp_jour_cum', {})
    creneaux    = metadata['creneaux']
    affectations= metadata['affectations']
    horaires_ag = metadata['horaires_agents']
    evenements  = metadata['evenements']
    samedi_type = week_data['samedi_type']
    sem_type    = week_data.get('semaine_type', '')

    mois_num = {'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
                'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,'décembre':12
                }.get(metadata['mois'].lower(), 1)
    mois_cap = MOIS_FR.get(mois_num, '')
    annee    = metadata['annee']

    sname = f"Planning_Agent_Semaine_{week_num}"
    if sname in wb.sheetnames: wb.remove(wb[sname])
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False

    # Jours présents dans la semaine
    jours_presents = [j for j in JOURS_SP if week_dates.get(j) is not None]
    nb_jours       = len(jours_presents)

    # Agents : réguliers triés par section primaire, puis vacataires
    agents_sorted = sorted(
        [a for a in affectations.keys() if not is_vacataire(a)],
        key=lambda x: (affectations.get(x, [''])[0], x)
    ) + [a for a in affectations.keys() if is_vacataire(a)]

    # Largeurs
    ws.column_dimensions['A'].width = 16   # Agent
    ws.column_dimensions['B'].width = 14   # Créneau
    for ci in range(3, 3 + nb_jours):
        ws.column_dimensions[get_column_letter(ci)].width = 20

    row = 1
    nb_cols  = 2 + nb_jours
    end_col  = get_column_letter(nb_cols)

    # Titre
    ws.merge_cells(f'A{row}:{end_col}{row}')
    dates_list = sorted([d for d in week_dates.values() if d])
    d0, d1 = dates_list[0], dates_list[-1]
    titre = (f"PLANNING PAR AGENT — Semaine {week_num}  |  "
             f"{d0.day} au {d1.day} {mois_cap} {annee}  |  {sem_type}")
    _set(ws.cell(row=row, column=1, value=titre),
         bg=C['hdr_dark'], fnt=_font(bold=True, size=13, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    # Légende
    ws.merge_cells(f'A{row}:{end_col}{row}')
    _set(ws.cell(row=row, column=1,
                 value="  SP = section (couleur)     Événement     Congé     "
                       "Grisé = hors horaires     Rouge vif = hors section habituelle"),
         bg='34495E', fnt=_font(size=9, italic=True, color='FFFFFF'), aln=_aln('left'))
    ws.row_dimensions[row].height = 15
    row += 2

    # Mapping jour → colonne
    col_map = {}
    for ji, jour in enumerate(jours_presents):
        col_map[jour] = 3 + ji

    for agent in agents_sorted:
        sect_prim = affectations.get(agent, ['?'])[0]
        hdr_bg    = C['vac_bg'] if is_vacataire(agent) else SEC_HDR.get(sect_prim, C['hdr_dark'])
        hdr_fc    = C['vac_txt'] if is_vacataire(agent) else 'FFFFFF'

        # En-tête agent
        ws.merge_cells(f'A{row}:{end_col}{row}')
        sects_label = ' / '.join(affectations.get(agent, []))
        _set(ws.cell(row=row, column=1,
                     value=f"  {agent.upper()}   ({sects_label})"),
             bg=hdr_bg, fnt=_font(bold=True, size=11, color=hdr_fc), aln=_aln('left'))
        ws.row_dimensions[row].height = 24
        row += 1

        # En-tête colonnes jours
        ws.cell(row=row, column=1).value = ''
        ws.cell(row=row, column=1).fill  = _fill(C['hdr_dark2'])
        _set(ws.cell(row=row, column=2, value='Créneau'),
             bg=C['hdr_dark2'], fnt=_font(bold=True, size=10, color='FFFFFF'),
             aln=_aln(), brd=_brd())
        for jour in jours_presents:
            c    = col_map[jour]
            date = week_dates.get(jour)
            sam_t = plan.get(jour, {}).get('_samedi_type')
            lbl  = f"{jour[:3]}\n{date.day}/{date.month}"
            if sam_t: lbl += f"\n{sam_t}"
            bg_h = (C['hdr_rouge'] if sam_t == 'ROUGE'
                    else C['hdr_bleu'] if sam_t == 'BLEU'
                    else C['hdr_dark2'])
            _set(ws.cell(row=row, column=c, value=lbl),
                 bg=bg_h, fnt=_font(bold=True, size=9, color='FFFFFF'),
                 aln=_aln('center'), brd=_brd())
        ws.row_dimensions[row].height = 36
        row += 1

        # Lignes créneaux
        for cren_name, cs, ce in creneaux:
            ws.cell(row=row, column=1).value = ''
            ws.cell(row=row, column=1).fill  = _fill('F8F9FA')
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
                conge    = any(n.lower() in ('congé','conge','rtt','vacation')
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

                sp_section = None
                is_oos_cell = False
                if slot and isinstance(slot, dict) and slot.get('open'):
                    for sect in SECTIONS:
                        if agent in slot.get('assignment', {}).get(sect, []):
                            sp_section = sect
                            break
                    if sp_section:
                        is_oos_cell = slot.get('out_of_section', {}).get(sp_section, False)

                if conge:
                    _set(cell, value='Congé', bg=C['alert_bg'],
                         fnt=_font(size=9, color=C['alert_txt'], italic=True),
                         aln=_aln(), brd=_brd())
                elif event:
                    ev_label = f"Évén.\n{event[0][:18]}"
                    _set(cell, value=ev_label, bg=C['event_bg'],
                         fnt=_font(size=9, italic=True, color='6E2F09'),
                         aln=_aln(), brd=_brd())
                elif sp_section and is_oos_cell:
                    _set(cell, value=f"SP\n{sp_section}\n⚠hors sect.",
                         bg=C['oos_bg'],
                         fnt=_font(size=8, bold=True, color=C['oos_txt']),
                         aln=_aln(), brd=_brd())
                elif sp_section:
                    fc_sp = C['vac_txt'] if is_vacataire(agent) else C['agent_txt']
                    _set(cell, value=f"SP\n{sp_section}",
                         bg=SEC_BG[sp_section] if not is_vacataire(agent) else C['vac_bg'],
                         fnt=_font(size=9, bold=True, color=fc_sp),
                         aln=_aln(), brd=_brd())
                elif not in_horaires:
                    _set(cell, value='', bg='E8E8E8',
                         fnt=_font(size=8, color=C['gray']),
                         aln=_aln(), brd=_brd())
                else:
                    _set(cell, value='', bg='FAFAFA',
                         fnt=_font(size=8, color=C['gray']),
                         aln=_aln(), brd=_brd())

            ws.row_dimensions[row].height = 24
            row += 1

        # Ligne total SP
        ws.cell(row=row, column=1).value = ''
        ws.cell(row=row, column=1).fill  = _fill('F0F3F4')
        _set(ws.cell(row=row, column=2, value='TOTAL SP'),
             bg='34495E', fnt=_font(bold=True, size=9, color='FFFFFF'),
             aln=_aln(), brd=_brd())
        for jour in jours_presents:
            c    = col_map[jour]
            sp_j = sp_jour_cum.get(jour, {}).get(agent, 0)
            val  = min_to_hhmm(sp_j) if sp_j > 0 else '—'
            _set(ws.cell(row=row, column=c), value=val,
                 bg='D5F5E3' if sp_j > 0 else 'F0F3F4',
                 fnt=_font(bold=sp_j>0, size=9,
                           color='1E8449' if sp_j>0 else C['gray']),
                 aln=_aln(), brd=_brd())
        ws.row_dimensions[row].height = 18
        row += 1

        # Séparateur
        for ci in range(1, nb_cols + 1):
            ws.cell(row=row, column=ci).fill = _fill('C8C8C8')
        ws.row_dimensions[row].height = 3
        row += 2

    return ws


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

    # Créer onglets semaines
    for week_data in weeks_data:
        write_week_sheet(wb, week_data, metadata)

    # Créer onglets planning agent par semaine
    for week_data in weeks_data:
        write_planning_agent_week_sheet(wb, week_data, metadata)

    # Réordonner onglets
    week_sheets  = [f"Semaine_{i}" for i in range(1, len(weeks_data)+1)
                    if f"Semaine_{i}" in wb.sheetnames]
    agent_sheets = [f"Planning_Agent_Semaine_{i}" for i in range(1, len(weeks_data)+1)
                    if f"Planning_Agent_Semaine_{i}" in wb.sheetnames]
    other_sheets = [s for s in wb.sheetnames
                    if s not in week_sheets and s not in agent_sheets]
    wb._sheets   = ([wb[s] for s in week_sheets]
                    + [wb[s] for s in agent_sheets]
                    + [wb[s] for s in other_sheets])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
