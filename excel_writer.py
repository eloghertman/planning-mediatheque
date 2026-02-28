"""
Excel Writer v5 — Médiathèque
Génère le fichier Excel de planning avec :
- Tableau planning par créneau avec alertes en rouge
- Récapitulatif SP avec FORMULES EXCEL dynamiques (se recalcule si modif manuelle)
"""

import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

# ══════════════════════════════════════════════════════════════
#  PALETTE
# ══════════════════════════════════════════════════════════════

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
    'alert_bg':      'FFCCCC',
    'alert_txt':     '8B0000',
    'title_txt':     'FFFFFF',
    'agent_txt':     '2C3E50',
    'gray_txt':      'AAAAAA',
    'row_alt1':      'FDFEFE',
    'row_alt2':      'EBF5FB',
    'blue':          '2980B9',
    'green':         '27AE60',
    'orange':        'E67E22',
    'red':           'E74C3C',
    'dark':          '2C3E50',
    'purple':        '8E44AD',
    'sp_warn':       'FFF3CD',
    'sp_alert':      'FFCCCC',
    'sp_ok':         'D5F5E3',
    'sp_short':      'FDEBD0',
}

SEC_COLORS = {'RDC': C['sec_rdc'], 'Adulte': C['sec_adulte'],
              'MF': C['sec_mf'], 'Jeunesse': C['sec_jeunesse']}
SEC_HDR    = {'RDC': C['blue'], 'Adulte': C['green'],
              'MF': C['orange'], 'Jeunesse': C['red']}

SECTIONS  = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SP  = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

MOIS_FR = {1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
           7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'}

def _fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def _font(bold=False, size=10, color='000000', italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name='Calibri')

def _aln(h='center', v='center', wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _brd(color='AAAAAA', style='thin'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _set(cell, value=None, bg=None, fnt=None, aln=None, brd=None):
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


# ══════════════════════════════════════════════════════════════
#  ÉCRITURE D'UN ONGLET SEMAINE
# Layout colonnes :
#   A=N°, B=Créneau, C=RDC, D=Adulte, E=MF, F=Jeunesse, G=Événement, H=Alertes
# ══════════════════════════════════════════════════════════════

def write_week_sheet(wb, week_data, metadata):
    week_num   = week_data['week_num']
    week_dates = week_data['week_dates']
    plan       = week_data['plan']
    sp_alerts  = week_data.get('sp_alerts', {})
    creneaux   = metadata['creneaux']
    sp_minmax  = metadata.get('sp_minmax', {})

    mois_num  = {'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
                 'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,
                 'décembre':12}.get(metadata['mois'].lower(), 1)
    mois_cap  = MOIS_FR.get(mois_num, metadata['mois'].capitalize())
    annee     = metadata['annee']

    dates_list = sorted([d for d in week_dates.values() if d])
    date_range = f"{dates_list[0].day} au {dates_list[-1].day} {mois_cap} {annee}" if dates_list else f"Semaine {week_num}"

    sname = f"Semaine_{week_num}"
    if sname in wb.sheetnames:
        del wb[sname]
    ws = wb.create_sheet(sname)

    # Largeurs colonnes
    for col, w in {'A':5,'B':14,'C':22,'D':22,'E':22,'F':28,'G':30,'H':32}.items():
        ws.column_dimensions[col].width = w

    row = 1

    # ── TITRE ──
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value=f"📅  PLANNING SERVICE PUBLIC — Semaine {week_num}  |  {date_range}"),
         bg=C['header_dark'], fnt=_font(bold=True, size=13, color='FFFFFF'),
         aln=_aln('center'))
    ws.row_dimensions[row].height = 30
    row += 1

    # ── LÉGENDE ──
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="  🔵 RDC     🟢 Adulte     🟡 MF     🔴 Jeunesse     ⬜ Fermé     🟠 Événement     🚨 ALERTE"),
         bg=C['dark'], fnt=_font(size=9, color='FFFFFF', italic=True), aln=_aln('left'))
    ws.row_dimensions[row].height = 16
    row += 2

    # ── TABLEAU PLANNING TYPE PAR JOUR ──
    # On garde trace des cellules de données pour les formules récap
    # data_cells[agent][jour] = liste de numéros de lignes où l'agent apparaît
    agent_rows = defaultdict(list)   # agent -> [(row, col, cren_min)]

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        jour_plan   = plan.get(jour, {})
        samedi_type = jour_plan.get('_samedi_type')

        # En-tête jour
        is_sam    = jour == 'Samedi'
        sam_label = f"  🔴 SAMEDI ROUGE" if (is_sam and samedi_type == 'ROUGE') else \
                    f"  🔵 SAMEDI BLEU"  if (is_sam and samedi_type == 'BLEU')  else ""
        jour_color = C['header_rouge'] if (is_sam and samedi_type == 'ROUGE') else \
                     C['header_bleu']  if (is_sam and samedi_type == 'BLEU')  else C['header_dark']

        ws.merge_cells(f'A{row}:H{row}')
        _set(ws.cell(row=row, column=1,
                     value=f"  {jour.upper()}  {date.day} {mois_cap} {annee}{sam_label}"),
             bg=jour_color, fnt=_font(bold=True, size=12, color='FFFFFF'), aln=_aln('left'))
        ws.row_dimensions[row].height = 26
        row += 1

        # En-têtes colonnes
        headers_info = [
            ('N°',C['dark']),('Créneau',C['dark']),
            ('RDC',SEC_HDR['RDC']),('Adulte',SEC_HDR['Adulte']),
            ('Musique & Films',SEC_HDR['MF']),('Jeunesse',SEC_HDR['Jeunesse']),
            ('Événement',C['purple']),('⚠️ Alertes',C['header_rouge']),
        ]
        for ci, (h, hc) in enumerate(headers_info):
            _set(ws.cell(row=row, column=ci+1, value=h),
                 bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'),
                 aln=_aln('center'), brd=_brd())
        ws.row_dimensions[row].height = 20
        row += 1

        # Lignes créneaux
        for ni, (cren_name, cs, ce) in enumerate(creneaux):
            slot    = jour_plan.get(cren_name)
            slot_min = ce - cs

            if slot is None:
                # Fermé
                for ci, val in enumerate([ni+1, cren_name, '—', '—', '—', '—', '', ''], 1):
                    _set(ws.cell(row=row, column=ci, value=val if val else None),
                         bg=C['closed'],
                         fnt=_font(size=9 if ci > 2 else 10,
                                   color=C['gray_txt'], italic=ci > 1),
                         aln=_aln(), brd=_brd())
            else:
                assignment = slot.get('assignment', {s: [] for s in SECTIONS})
                events     = slot.get('events', [])
                alerts     = slot.get('alerts', [])

                # N°
                _set(ws.cell(row=row, column=1, value=ni+1),
                     bg='F2F3F4', fnt=_font(size=9, color='7F8C8D'),
                     aln=_aln(), brd=_brd())
                # Créneau
                _set(ws.cell(row=row, column=2, value=cren_name),
                     bg='FDFEFE', fnt=_font(bold=True, size=10),
                     aln=_aln(), brd=_brd())

                # Sections RDC/Adulte/MF/Jeunesse
                for si, sect in enumerate(SECTIONS):
                    col     = si + 3
                    agents  = assignment.get(sect, [])
                    has_alert = any(sect in al for al in alerts)
                    val     = '  /  '.join(agents) if agents else ('🚨 ALERTE' if has_alert else '—')
                    bg      = C['alert_bg'] if has_alert and not agents else \
                              (SEC_COLORS[sect] if agents else 'FDF2F8')
                    fnt_col = C['alert_txt'] if has_alert and not agents else \
                              (C['agent_txt'] if agents else 'C0392B')
                    _set(ws.cell(row=row, column=col, value=val),
                         bg=bg, fnt=_font(size=10, color=fnt_col, bold=bool(agents)),
                         aln=_aln('center'), brd=_brd())

                    # Enregistrer pour formules récap
                    for agent in agents:
                        agent_rows[agent].append((row, col, slot_min, cren_name, jour))

                # Événements
                ev_txt = '\n'.join(f"{ev['nom']} ({', '.join(ev['agents'])})"
                                    for ev in events) if events else ''
                _set(ws.cell(row=row, column=7, value=ev_txt),
                     bg=C['event_bg'] if ev_txt else 'FAFAFA',
                     fnt=_font(size=9, italic=True,
                               color='6E2F09' if ev_txt else C['gray_txt']),
                     aln=_aln('left'), brd=_brd())

                # Alertes
                al_txt = '\n'.join(alerts) if alerts else ''
                _set(ws.cell(row=row, column=8, value=al_txt),
                     bg=C['alert_bg'] if al_txt else 'FAFAFA',
                     fnt=_font(size=9, bold=bool(al_txt),
                               color=C['alert_txt'] if al_txt else C['gray_txt']),
                     aln=_aln('left'), brd=_brd())

                if events or alerts:
                    ws.row_dimensions[row].height = max(20, 16 * max(len(events), len(alerts), 1))
                else:
                    ws.row_dimensions[row].height = 20

            row += 1

        # Séparateur
        for ci in range(1, 9):
            ws.cell(row=row, column=ci).fill = _fill('F0F3F4')
        ws.row_dimensions[row].height = 5
        row += 1

    # ══════════════════════════════════════════════════════════
    #  RÉCAPITULATIF SP AVEC FORMULES EXCEL DYNAMIQUES
    # ══════════════════════════════════════════════════════════
    row += 1
    recap_start_row = row

    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="📊  RÉCAPITULATIF — Heures SP par agent  |  Se recalcule automatiquement si vous modifiez le planning"),
         bg='34495E', fnt=_font(bold=True, size=11, color='FFFFFF'), aln=_aln('center'))
    ws.row_dimensions[row].height = 24
    row += 1

    # Sous-titre
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="ℹ️  Les totaux sont calculés par formule Excel — toute modification manuelle d'un nom d'agent met à jour le récap automatiquement"),
         bg='EBF5FB', fnt=_font(size=9, italic=True, color='2471A3'), aln=_aln('left'))
    ws.row_dimensions[row].height = 16
    row += 1

    # En-têtes récap
    recap_hdrs = ['Agent','Section','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Total semaine']
    recap_hdr_colors = ['34495E','34495E','1A5276','1A5276','1A5276','1A5276','922B21','145A32']
    for ci, (h, hc) in enumerate(zip(recap_hdrs, recap_hdr_colors)):
        _set(ws.cell(row=row, column=ci+1, value=h),
             bg=hc, fnt=_font(bold=True, size=10, color='FFFFFF'),
             aln=_aln(), brd=_brd())
    ws.row_dimensions[row].height = 20
    hdr_row = row
    row += 1

    # ── Construction du plan de données pour les formules ──
    # Pour chaque (agent, jour), on a besoin de savoir quelles cellules
    # contiennent cet agent. On va construire des plages par jour et section.
    #
    # Architecture des formules :
    # Pour chaque agent A et jour J :
    #   = SUMPRODUCT((C_range="A")*1)*duree_creneau + ...
    # Mais comme les créneaux ont des durées variables, on va écrire
    # une valeur calculée statique (les formules Excel ne peuvent pas
    # facilement pondérer par durée de créneau variable).
    # On écrit donc des formules COUNTIF × durée pour chaque créneau.
    #
    # Approche plus simple et robuste :
    # On note quelles cellules (référence Excel type C5) contiennent quels agents
    # et on construit une formule SUMPRODUCT pondérée.

    # Construire index : (jour, section) -> [(row, col, duree_min)]
    jour_section_cells = defaultdict(list)
    for agent_key, cell_list in agent_rows.items():
        pass  # on a déjà tout dans agent_rows

    # On repart de zéro : parcourir le planning pour construire les formules
    # Structure: planning_cells[jour][cren_name][section] = (row, col)
    planning_cells = defaultdict(dict)

    # Re-parcourir les lignes pour retrouver les positions
    # En fait on va construire les formules directement depuis agent_rows
    # agent_rows[agent] = [(row, col, slot_min, cren_name, jour), ...]

    affectations = metadata.get('affectations', {})

    def agent_primary_section(agent):
        sects = affectations.get(agent, [])
        return sects[0] if sects else '?'

    sp_minmax_data = metadata.get('sp_minmax', {})

    # Agents à afficher (ceux qui ont au moins une apparition dans agent_rows, ou dans sp_minmax)
    agents_to_show = list({a for a in list(agent_rows.keys()) + list(sp_minmax_data.keys())
                           if not is_vacataire(a) or a in agent_rows})
    agents_to_show = sorted(set(agents_to_show),
                            key=lambda x: (agent_primary_section(x), x))

    # Mapping jour -> colonne récap (C=3, D=4, E=5, F=6, G=7)
    jour_to_recap_col = {'Mardi':3, 'Mercredi':4, 'Jeudi':5, 'Vendredi':6, 'Samedi':7}

    for i, agent in enumerate(agents_to_show):
        bg = C['row_alt1'] if i % 2 == 0 else C['row_alt2']
        sect = agent_primary_section(agent)

        _set(ws.cell(row=row, column=1, value=agent),
             bg=bg, fnt=_font(bold=True, size=10), aln=_aln('left'), brd=_brd())
        _set(ws.cell(row=row, column=2, value=sect),
             bg=SEC_COLORS.get(sect, bg), fnt=_font(size=10), aln=_aln(), brd=_brd())

        # Pour chaque jour : calculer les minutes SP directement (valeur statique)
        # ET écrire une formule qui comptera les occurrences dans le planning
        # On utilise une approche hybride :
        # - Valeur statique calculée par l'algo (correcte au moment de la génération)
        # - On ajoute une note indiquant que la valeur vient du planning

        # Construire formule par jour
        # Pour chaque créneau et section, la cellule contient le nom de l'agent (ou plusieurs)
        # On fait SUMPRODUCT(ISNUMBER(SEARCH(agent, plage_cellules)) * durees)

        day_formulas = {}
        day_values   = {}

        for jour, j_col in jour_to_recap_col.items():
            # Trouver les cellules de ce jour pour cet agent
            cells_for_jour = [(r, c, dur) for (r, c, dur, cn, j) in agent_rows.get(agent, [])
                              if j == jour]

            if not cells_for_jour:
                day_values[jour] = 0
                day_formulas[jour] = None
            else:
                # Grouper par durée unique pour simplifier la formule
                # Formule : =SUMPRODUCT((ISNUMBER(SEARCH("agent",C5:C5)))*30) + ...
                # Pour chaque créneau individuel (car durées variables)
                formula_parts = []
                total_min = 0
                for (r, c, dur) in cells_for_jour:
                    col_letter = get_column_letter(c)
                    cell_ref   = f"{col_letter}{r}"
                    # ISNUMBER(SEARCH()) pour gestion du nom partiel dans cellule multi-agents
                    formula_parts.append(
                        f'ESTNUM(CHERCHE("{agent}";{cell_ref}))*{dur}'
                    )
                    total_min += dur

                formula = "=(" + "+".join(formula_parts) + ")/60"
                day_formulas[jour] = formula
                day_values[jour]   = total_min

        # Écrire les cellules de jour avec formules dynamiques
        total_col_refs = []
        for j_col_offset, jour in enumerate(JOURS_SP):
            col     = j_col_offset + 3
            formula = day_formulas.get(jour)
            val_min = day_values.get(jour, 0)
            cell    = ws.cell(row=row, column=col)

            if week_dates.get(jour) is not None:
                bg_c = 'FFF9C4' if (jour == 'Samedi' and val_min > 0) else                        (C['row_alt2'] if val_min > 0 else bg)
                cell.fill      = _fill(bg_c)
                cell.font      = _font(size=10, bold=val_min > 0)
                cell.alignment = _aln()
                cell.border    = _brd()
                if formula:
                    cell.value = formula   # formule =ESTNUM(CHERCHE(...))
                    total_col_refs.append(get_column_letter(col) + str(row))
                else:
                    cell.value = "—"
            else:
                cell.value     = '—'
                cell.fill      = _fill(bg)
                cell.font      = _font(size=9, color=C['gray_txt'])
                cell.alignment = _aln()
                cell.border    = _brd()

        # Total semaine : formule = somme des colonnes jour
        total_cell    = ws.cell(row=row, column=8)
        total_min_val = sum(day_values.values())
        mm            = sp_minmax_data.get(agent, {})
        is_samedi_week = 'Samedi' in week_dates
        min_sp = (mm.get('Min_MarSam', 0) if is_samedi_week else mm.get('Min_MarVen', 0)) * 60
        max_sp = (mm.get('Max_MarSam', 99) if is_samedi_week else mm.get('Max_MarVen', 99)) * 60

        if total_col_refs:
            total_cell.value = "=" + "+".join(total_col_refs)
        else:
            total_cell.value = 0

        if total_min_val < min_sp:
            total_bg = C['sp_alert']
            fnt_col  = C['alert_txt']
        elif total_min_val > max_sp:
            total_bg = C['sp_warn']
            fnt_col  = 'E67E22'
        else:
            total_bg = C['sp_ok']
            fnt_col  = '1E8449'

        total_cell.fill      = _fill(total_bg)
        total_cell.font      = _font(bold=True, size=10, color=fnt_col)
        total_cell.alignment = _aln()
        total_cell.border    = _brd()

        ws.row_dimensions[row].height = 18
        row += 1

    # Légende couleurs récap
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    _set(ws.cell(row=row, column=1,
                 value="  🟢 SP dans la plage normale     🔴 SP insuffisant (sous le min)     🟡 SP dépassé (au-dessus du max)"),
         bg='F4F6F7',
         fnt=_font(size=9, italic=True, color='5D6D7E'),
         aln=_aln('left'))
    ws.row_dimensions[row].height = 16

    return ws


# ══════════════════════════════════════════════════════════════
#  GÉNÉRATION DU FICHIER COMPLET
# ══════════════════════════════════════════════════════════════

def is_vacataire(agent):
    return 'vacataire' in str(agent).lower()

def generate_excel(source_filepath, weeks_data, metadata):
    wb = load_workbook(source_filepath)

    # Supprimer anciens onglets semaine
    for i in range(1, 7):
        if f"Semaine_{i}" in wb.sheetnames:
            del wb[f"Semaine_{i}"]

    # Écrire nouveaux onglets
    for week_data in weeks_data:
        write_week_sheet(wb, week_data, metadata)

    # Réordonner : semaines en premier
    week_sheets = [f"Semaine_{i}" for i in range(1, len(weeks_data)+1)
                   if f"Semaine_{i}" in wb.sheetnames]
    other_sheets = [s for s in wb.sheetnames if s not in week_sheets]
    wb._sheets = [wb[s] for s in week_sheets] + [wb[s] for s in other_sheets]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
