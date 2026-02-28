"""
Planning Engine v3 — Médiathèque
Logique : planning type comme base → ajustements uniquement là où les contraintes l'imposent
Règles :
  - Blocs de 2h ou 2h30 privilégiés (pas de changement toutes les heures)
  - Max 3 agents différents/section/jour (Mar/Jeu/Ven) — max 4 (Mer/Sam)
  - Pas de vacataire seul en Jeunesse sur 10h-12h et 14h-19h (exception 12h-14h)
  - Pas de vacataire seul (toutes sections) plus de 2h consécutives
  - Partir du planning type, modifier seulement les créneaux impactés
"""

import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict


# ─────────────────────────────────────────────
#  UTILITAIRES HORAIRES
# ─────────────────────────────────────────────

def hm_to_min(val):
    """
    Convertit en minutes depuis minuit.
    Gère : datetime.time, timedelta, '9h', '9h30', '9:30', '09:30', '9h30m', etc.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if hasattr(val, 'hour'):
        return val.hour * 60 + val.minute
    if isinstance(val, timedelta):
        return int(val.total_seconds()) // 60
    if isinstance(val, str):
        val = val.strip()
        if not val or val.lower() == 'nan':
            return None
        # Normaliser : remplacer H/h par :, supprimer les 'm' finaux
        normalized = val.lower().replace('h', ':').replace('m', '').rstrip(':')
        parts = normalized.split(':')
        try:
            h = int(parts[0]) if parts[0] else 0
            m = int(parts[1]) if len(parts) >= 2 and parts[1] else 0
            return h * 60 + m
        except:
            return None
    return None

def min_to_hhmm(minutes):
    if minutes is None:
        return ""
    return f"{minutes//60:02d}:{minutes%60:02d}"

def overlap(s1, e1, s2, e2):
    return s1 < e2 and e1 > s2

def is_vacataire(agent):
    return 'vacataire' in str(agent).lower()


# ─────────────────────────────────────────────
#  LECTURE DU FICHIER EXCEL
# ─────────────────────────────────────────────

def load_excel_data(filepath):
    xl = pd.ExcelFile(filepath)
    data = {}
    for sheet in xl.sheet_names:
        data[sheet] = pd.read_excel(filepath, sheet_name=sheet, header=None)
    return data

def parse_parametres(data):
    df = data['Paramètres']
    params = {}
    for _, row in df.iterrows():
        key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        val = row.iloc[1] if pd.notna(row.iloc[1]) else None
        if key and key != 'Paramètre':
            params[key] = val
    return params

def parse_horaires_ouverture(data):
    df = data['Horaire_ouverture_mediatheque']
    result = {}
    for _, row in df.iterrows():
        jour = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not jour or jour in ('Jour', 'nan'):
            continue
        slots = []
        for i in range(2, len(row) - 1, 2):
            s = hm_to_min(row.iloc[i])   if pd.notna(row.iloc[i])   else None
            e = hm_to_min(row.iloc[i+1]) if pd.notna(row.iloc[i+1]) else None
            if s is not None and e is not None and s != e:
                slots.append((s, e))
        if slots:
            result[jour] = slots
    return result

def parse_affectations(data):
    df = data['Affectations']
    result = {}
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not agent or agent == 'Agent':
            continue
        sections = [str(row.iloc[i]).strip() for i in range(1, len(row))
                    if pd.notna(row.iloc[i]) and str(row.iloc[i]).strip() not in ('nan', '')]
        result[agent] = sections
    return result

def parse_horaires_agents(data):
    df = data['Horaires_Des_Agents']
    result = defaultdict(dict)
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        jour  = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not agent or not jour or agent == 'Agent':
            continue
        sm = hm_to_min(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        em = hm_to_min(row.iloc[3]) if pd.notna(row.iloc[3]) else None
        sa = hm_to_min(row.iloc[4]) if pd.notna(row.iloc[4]) else None
        ea = hm_to_min(row.iloc[5]) if pd.notna(row.iloc[5]) else None
        result[agent][jour] = (sm, em, sa, ea)
    return dict(result)

def parse_besoins_jeunesse(data):
    df = data['Besoins_Jeunesse']
    result = {}
    cols = [str(c).strip() for c in df.iloc[0]]
    for _, row in df.iterrows():
        cren = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not cren or cren == 'Créneau':
            continue
        entry = {}
        for i, col in enumerate(cols[1:], 1):
            v = row.iloc[i]
            entry[col] = int(v) if pd.notna(v) else 0
        result[cren] = entry
    return result

def parse_sp_minmax(data):
    df = data['SP_MinMax']
    result = {}
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not agent or agent == 'Agent':
            continue
        result[agent] = {
            'Min_MarVen':     float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
            'Max_MarVen':     float(row.iloc[2]) if pd.notna(row.iloc[2]) else 99,
            'Min_MarSam':     float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
            'Max_MarSam':     float(row.iloc[4]) if pd.notna(row.iloc[4]) else 99,
            'SP_Samedi_Type': float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
        }
    return result

def parse_roulement_samedi(data):
    df = data['Roulement_Samedi']
    result = {}
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        val   = str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else None
        if not agent or agent == 'Agent':
            continue
        if val in ('BLEU', 'ROUGE'):
            result[agent] = val
    return result

# Dictionnaire mois texte → numéro
MOIS_TEXTE_MAP = {
    'janvier':1,'février':2,'fevrier':2,'mars':3,'avril':4,'mai':5,'juin':6,
    'juillet':7,'août':8,'aout':8,'septembre':9,'octobre':10,'novembre':11,
    'décembre':12,'decembre':12
}
JOURS_TEXTE = {'lundi','mardi','mercredi','jeudi','vendredi','samedi','dimanche'}

def parse_date_flexible(val, annee):
    """Parse une date sous toutes ses formes, y compris 'Samedi 7 mars'."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val
    try:
        ts = pd.Timestamp(val)
        if not pd.isna(ts):
            return ts.to_pydatetime()
    except:
        pass
    s = str(val).strip().lower()
    if not s or s == 'nan':
        return None
    # Retirer le jour de la semaine ("samedi 7 mars" → "7 mars")
    for j in JOURS_TEXTE:
        if s.startswith(j):
            s = s[len(j):].strip()
            break
    # "7 mars", "07 mars 2026"
    parts = s.replace(',', ' ').split()
    if len(parts) >= 2:
        try:
            jour_num = int(parts[0])
            mois_nom = parts[1].rstrip('.')
            mois_num = MOIS_TEXTE_MAP.get(mois_nom)
            if mois_num:
                an = int(parts[2]) if len(parts) >= 3 else annee
                return datetime(an, mois_num, jour_num)
        except:
            pass
    # Formats numériques
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d/%m'):
        try:
            d = datetime.strptime(s, fmt)
            if d.year == 1900:
                d = d.replace(year=annee)
            return d
        except:
            pass
    return None

def split_agents_cell(cell_value):
    """
    Découpe une cellule agents en liste.
    Gère les séparateurs ; et les espaces autour.
    Ex: 'Anne-Françoise ; Guillaume ; Macha' → ['Anne-Françoise', 'Guillaume', 'Macha']
    """
    if not cell_value or str(cell_value).strip() in ('nan', '', '-'):
        return []
    raw = str(cell_value).strip()
    parts = raw.split(';')
    return [p.strip() for p in parts if p.strip() and p.strip().lower() != 'nan']

def parse_evenements(data, mois, annee):
    df = data['Événements']
    result = defaultdict(list)
    for _, row in df.iterrows():
        date_val = row.iloc[0]
        if not pd.notna(date_val):
            continue
        d = parse_date_flexible(date_val, annee)
        if d is None or d.month != mois or d.year != annee:
            continue
        debut = hm_to_min(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        fin   = hm_to_min(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        nom   = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None
        if not nom or nom == 'nan' or debut is None:
            continue
        # Chaque cellule agents peut contenir plusieurs noms séparés par ;
        agents = []
        for i in range(4, min(12, len(row))):
            if pd.notna(row.iloc[i]):
                agents.extend(split_agents_cell(row.iloc[i]))
        result[d.strftime('%Y-%m-%d')].append({
            'debut':  debut,
            'fin':    fin if fin else debut + 60,
            'nom':    nom,
            'agents': agents,
        })
    return dict(result)

def parse_creneaux(params):
    raw = params.get('Liste_des_créneaux', '')
    if not raw:
        return []
    creneaux = []
    for c in str(raw).split(';'):
        c = c.strip()
        if '-' in c:
            parts = c.split('-')
            s = hm_to_min(parts[0].strip())
            e = hm_to_min(parts[1].strip())
            if s is not None and e is not None:
                creneaux.append((c, s, e))
    return creneaux


# ─────────────────────────────────────────────
#  PARSING DU PLANNING TYPE
# ─────────────────────────────────────────────

def parse_planning_type(data):
    """
    Lit l'onglet planning_type.
    Retourne { 'Mardi'|'Mercredi'|...|'Samedi_ROUGE'|'Samedi_BLEU' ->
               [ (bloc_label, {section: [agents]}) ] }
    """
    df = data['planning_type']

    SECTION_COLS = {
        'RDC':      [2],
        'Adulte':   [3],
        'MF':       [5, 6],
        'Jeunesse': [7, 8, 9],
    }
    JOUR_MAP = {
        'MARDI': 'Mardi', 'MERCREDI': 'Mercredi', 'JEUDI': 'Jeudi',
        'VENDREDI': 'Vendredi',
    }

    raw_blocs    = {}
    current_jour = None
    samedi_count = 0

    for _, row in df.iterrows():
        cell0 = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ''
        cell1 = str(row.iloc[1]).strip()          if pd.notna(row.iloc[1]) else ''

        if cell0 in JOUR_MAP:
            current_jour = JOUR_MAP[cell0]
            raw_blocs[current_jour] = []
            continue

        if cell0 == 'SAMEDI':
            samedi_count += 1
            current_jour = 'Samedi_ROUGE' if samedi_count == 1 else 'Samedi_BLEU'
            raw_blocs[current_jour] = []
            continue

        if not current_jour:
            continue

        skip_cells = {'nan', '', 'R D C', 'Adulte', 'Musique & Films',
                      'Jeunesse', '12H30', '15H30'}
        if not cell1 or cell1 in skip_cells:
            continue

        bloc = {s: [] for s in ['RDC', 'Adulte', 'MF', 'Jeunesse']}
        for section, col_indices in SECTION_COLS.items():
            for ci in col_indices:
                if ci < len(row):
                    v = str(row.iloc[ci]).strip() if pd.notna(row.iloc[ci]) else ''
                    if v and v not in ('nan', '', '-', 'NaN'):
                        agent = v.split(' à ')[0].split(' à partir')[0].strip()
                        if agent:
                            bloc[section].append(agent)

        raw_blocs[current_jour].append((cell1, bloc))

    return raw_blocs


def explode_planning_type_to_creneaux(raw_blocs, creneaux_def):
    """
    Convertit les blocs du planning type (ex: "10H-12H30")
    en créneaux unitaires (ex: "10:00-11:00", "11:00-12:00", "12:00-12:30").
    Retourne { jour_key -> { cren_name -> { section -> [agents] } } }
    """
    def parse_bloc_time(label):
        label = label.strip().upper()
        label = label.replace('H', ':')
        if '-' not in label:
            return None, None
        parts = label.split('-')
        def fix(p):
            p = p.strip().rstrip(':')
            if ':' not in p:
                p += ':00'
            return p
        s = hm_to_min(fix(parts[0]))
        e = hm_to_min(fix(parts[1]))
        return s, e

    result = {}
    for jour_key, blocs in raw_blocs.items():
        result[jour_key] = {}
        for cren_label, bloc_assignment in blocs:
            bs, be = parse_bloc_time(cren_label)
            if bs is None or be is None:
                continue
            for cren_name, cs, ce in creneaux_def:
                if cs >= bs and ce <= be:
                    result[jour_key][cren_name] = {
                        s: list(agents) for s, agents in bloc_assignment.items()
                    }
    return result


# ─────────────────────────────────────────────
#  DISPONIBILITÉ
# ─────────────────────────────────────────────

JOURS_SP = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
SECTIONS  = ['RDC', 'Adulte', 'MF', 'Jeunesse']
MOIS_MAP  = {
    'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
    'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,'décembre':12
}

def agent_available_for_sp(agent, jour, cs, ce, date_str, horaires_agents, evenements):
    h = horaires_agents.get(agent, {}).get(jour)
    if h is None:
        return False
    sm, em, sa, ea = h
    in_m = sm is not None and em is not None and cs >= sm and ce <= em
    in_a = sa is not None and ea is not None and cs >= sa and ce <= ea
    if not (in_m or in_a):
        return False
    for ev in evenements.get(date_str, []):
        if agent in ev['agents'] and overlap(cs, ce, ev['debut'], ev['fin']):
            return False
    return True

def creneau_is_open(cs, ce, jour, horaires_ouverture):
    for os, oe in horaires_ouverture.get(jour, []):
        if cs >= os and ce <= oe:
            return True
    return False

def exception_vacataire_seul(cs, ce, params):
    exc = params.get('Exception_Vacataire_seul', '12:00-14:00')
    if not exc or str(exc) == 'nan':
        return False
    parts = str(exc).split('-')
    if len(parts) == 2:
        es = hm_to_min(parts[0].strip())
        ee = hm_to_min(parts[1].strip())
        if es is not None and ee is not None:
            return cs >= es and ce <= ee
    return False


# ─────────────────────────────────────────────
#  SEMAINES ET SAMEDIS
# ─────────────────────────────────────────────

def get_weeks_of_month(mois, annee):
    first_day = datetime(annee, mois, 1)
    days_to_tuesday = (1 - first_day.weekday()) % 7
    first_tuesday = first_day + timedelta(days=days_to_tuesday)
    weeks = []
    current = first_tuesday
    offsets = {'Mardi': 0, 'Mercredi': 1, 'Jeudi': 2, 'Vendredi': 3, 'Samedi': 4}
    while current.month == mois:
        week = {j: current + timedelta(days=o) for j, o in offsets.items()
                if (current + timedelta(days=o)).month == mois}
        if week:
            weeks.append(week)
        current += timedelta(weeks=1)
    return weeks

def compute_samedi_type(week_num):
    return 'ROUGE' if week_num % 2 == 1 else 'BLEU'

def get_agents_for_samedi(roulement_samedi, samedi_type, affectations, horaires_agents):
    agents = [a for a, r in roulement_samedi.items()
              if r == samedi_type and horaires_agents.get(a, {}).get('Samedi') is not None]
    for agent in affectations:
        if is_vacataire(agent) and agent not in agents \
           and horaires_agents.get(agent, {}).get('Samedi') is not None:
            agents.append(agent)
    return agents


# ─────────────────────────────────────────────
#  REMPLACEMENT D'AGENT
# ─────────────────────────────────────────────

def find_replacement(section, jour, cs, ce, date_str,
                     eligible, affectations, horaires_agents, evenements,
                     agents_used_today_section, exclude=None):
    """
    Cherche le meilleur remplaçant disponible pour une section/créneau.
    Priorité :
      1. Agent déjà utilisé dans cette section aujourd'hui (continuité de bloc)
      2. Agent régulier (non vacataire)
      3. Agent avec le moins de créneaux utilisés aujourd'hui
    """
    exclude = exclude or set()
    candidates = []
    for agent in eligible:
        if agent in exclude:
            continue
        if section not in affectations.get(agent, []):
            continue
        if not agent_available_for_sp(agent, jour, cs, ce, date_str,
                                       horaires_agents, evenements):
            continue
        already_today = agent in agents_used_today_section
        candidates.append((
            not already_today,   # priorité aux agents déjà utilisés
            is_vacataire(agent), # réguliers avant vacataires
            agent
        ))

    if not candidates:
        return None
    candidates.sort()
    return candidates[0][2]


# ─────────────────────────────────────────────
#  CONTRAINTES VACATAIRE
# ─────────────────────────────────────────────

def check_vacataire_constraints(assignment, section, jour, cs, ce, date_str,
                                 creneaux, day_assignments,
                                 eligible, affectations, horaires_agents,
                                 evenements, params):
    """
    Pour une section donnée, vérifie et corrige si nécessaire :
    1. Jeunesse : pas de vacataire seul hors fenêtre exception (10h-12h, 14h-19h)
    2. Toutes sections : pas de vacataire seul plus de 2h consécutives
    Retourne la liste d'agents corrigée.
    """
    agents = assignment.get(section, [])
    if not agents or not all(is_vacataire(a) for a in agents):
        return agents  # au moins un régulier, pas de problème

    is_exc = exception_vacataire_seul(cs, ce, params)
    MAX_VAC_ALONE_MIN = 120

    need_regular = False

    # Règle 1 : Jeunesse hors exception
    if section == 'Jeunesse' and not is_exc:
        need_regular = True

    # Règle 2 : 2h consécutives vacataire seul (toutes sections)
    if not need_regular:
        cumul = 0
        for cn, c_cs, c_ce in creneaux:
            if c_cs == cs and c_ce == ce:
                break
            prev = day_assignments.get(cn)
            if prev is None:
                cumul = 0
                continue
            prev_agents = prev.get(section, [])
            if prev_agents and all(is_vacataire(a) for a in prev_agents):
                cumul += (c_ce - c_cs)
            else:
                cumul = 0
        if cumul >= MAX_VAC_ALONE_MIN:
            need_regular = True

    if not need_regular:
        return agents

    # Chercher un régulier disponible
    used_today = [a for cn, _, _ in creneaux
                  for a in day_assignments.get(cn, {}).get(section, [])
                  if not is_vacataire(a)]
    exclude = set(a for s in SECTIONS for a in assignment.get(s, [])
                  if a not in agents)

    repl = find_replacement(
        section=section, jour=jour, cs=cs, ce=ce, date_str=date_str,
        eligible=eligible, affectations=affectations,
        horaires_agents=horaires_agents, evenements=evenements,
        agents_used_today_section=used_today,
        exclude=exclude,
    )
    if repl:
        return [repl] + agents
    return agents


# ─────────────────────────────────────────────
#  ALGORITHME PRINCIPAL
# ─────────────────────────────────────────────

def plan_week(week_num, week_dates, planning_type_base,
              affectations, horaires_agents, horaires_ouverture,
              besoins_jeunesse, sp_minmax, roulement_samedi,
              evenements, params, creneaux):

    samedi_type   = compute_samedi_type(week_num)
    samedi_agents = get_agents_for_samedi(roulement_samedi, samedi_type,
                                          affectations, horaires_agents)
    all_agents    = list(affectations.keys())
    vac_days      = str(params.get('Mode_vacataires', 'mercredi,samedi')).lower()

    MAX_AGENTS_PER_SECTION = {
        'Mardi': 3, 'Mercredi': 4, 'Jeudi': 3, 'Vendredi': 3, 'Samedi': 4
    }

    result = {}

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        date_str = date.strftime('%Y-%m-%d')

        # Agents éligibles ce jour
        if jour == 'Samedi':
            eligible = samedi_agents
        elif jour.lower() in vac_days:
            eligible = [a for a in all_agents
                       if horaires_agents.get(a, {}).get(jour) is not None]
        else:
            eligible = [a for a in all_agents if not is_vacataire(a)
                       and horaires_agents.get(a, {}).get(jour) is not None]

        # Clé du planning type
        pt_key = f"Samedi_{samedi_type}" if jour == 'Samedi' else jour
        base   = planning_type_base.get(pt_key, {})
        max_rot = MAX_AGENTS_PER_SECTION.get(jour, 4)

        # Suivi des agents utilisés par section aujourd'hui (pour limiter la rotation)
        agents_used_today = {s: [] for s in SECTIONS}
        day_assignments   = {}

        result[jour] = {'_samedi_type': samedi_type if jour == 'Samedi' else None}

        for cren_name, cs, ce in creneaux:
            if not creneau_is_open(cs, ce, jour, horaires_ouverture):
                result[jour][cren_name] = None
                continue

            # Point de départ : planning type
            base_slot  = base.get(cren_name, {s: [] for s in SECTIONS})
            assignment = {s: list(base_slot.get(s, [])) for s in SECTIONS}

            # Vérifier chaque agent, remplacer si bloqué
            for section in SECTIONS:
                final_agents = []
                for agent in assignment[section]:
                    # Normaliser "VACATAIRES" générique
                    if agent.upper() in ('VACATAIRES', 'VACATAIRE'):
                        vac_avail = [a for a in eligible if is_vacataire(a)
                                    and agent_available_for_sp(a, jour, cs, ce, date_str,
                                                                horaires_agents, evenements)
                                    and a not in final_agents]
                        if vac_avail:
                            final_agents.append(vac_avail[0])
                        continue

                    if agent_available_for_sp(agent, jour, cs, ce, date_str,
                                              horaires_agents, evenements):
                        final_agents.append(agent)
                    else:
                        # Chercher remplaçant — privilégier continuité de bloc
                        already_in = agents_used_today[section]
                        excl = set(final_agents) | set(
                            a for s2 in SECTIONS for a in assignment.get(s2, [])
                            if s2 != section)
                        repl = find_replacement(
                            section=section, jour=jour, cs=cs, ce=ce,
                            date_str=date_str, eligible=eligible,
                            affectations=affectations,
                            horaires_agents=horaires_agents,
                            evenements=evenements,
                            agents_used_today_section=already_in,
                            exclude=excl,
                        )
                        if repl:
                            final_agents.append(repl)

                # Respecter la limite de rotation (max N agents différents/section/jour)
                adjusted = []
                for agent in final_agents:
                    unique_reg = [a for a in agents_used_today[section]
                                 if not is_vacataire(a)]
                    if not is_vacataire(agent) and agent not in agents_used_today[section] \
                       and len(unique_reg) >= max_rot:
                        # Limite atteinte → réutiliser le dernier agent régulier dispo
                        last = next(
                            (a for a in reversed(agents_used_today[section])
                             if not is_vacataire(a)
                             and agent_available_for_sp(a, jour, cs, ce, date_str,
                                                         horaires_agents, evenements)),
                            None
                        )
                        if last and last not in adjusted:
                            adjusted.append(last)
                        else:
                            adjusted.append(agent)  # on garde quand même si pas d'autre choix
                    else:
                        adjusted.append(agent)

                assignment[section] = adjusted

                # Mémoriser agents utilisés
                for a in adjusted:
                    if a not in agents_used_today[section]:
                        agents_used_today[section].append(a)

            # Compléter Jeunesse selon besoins
            bj = besoins_jeunesse.get(cren_name, {})
            if jour == 'Samedi':
                j_needs = int(bj.get(f'Samedi_{samedi_type.lower()}',
                                      bj.get('Samedi_rouge', 1)))
            else:
                j_needs = int(bj.get(jour, 1))
            j_needs = max(j_needs, 1)

            if len(assignment['Jeunesse']) < j_needs:
                occupied = set(a for s in SECTIONS for a in assignment[s])
                for agent in eligible:
                    if len(assignment['Jeunesse']) >= j_needs:
                        break
                    if agent in occupied:
                        continue
                    if 'Jeunesse' not in affectations.get(agent, []):
                        continue
                    if not agent_available_for_sp(agent, jour, cs, ce, date_str,
                                                   horaires_agents, evenements):
                        continue
                    assignment['Jeunesse'].append(agent)
                    if agent not in agents_used_today['Jeunesse']:
                        agents_used_today['Jeunesse'].append(agent)

            # Appliquer contraintes vacataire section par section
            for section in SECTIONS:
                assignment[section] = check_vacataire_constraints(
                    assignment, section, jour, cs, ce, date_str,
                    creneaux, day_assignments,
                    eligible, affectations, horaires_agents, evenements, params
                )

            # Événements du créneau
            slot_events = [ev for ev in evenements.get(date_str, [])
                          if overlap(cs, ce, ev['debut'], ev['fin'])]

            day_assignments[cren_name] = assignment
            result[jour][cren_name] = {
                'assignment': assignment,
                'events':     slot_events,
                'open':       True,
            }

    # Compteur SP
    sp_count = defaultdict(int)
    for jour, jour_plan in result.items():
        for cren_name, cs, ce in creneaux:
            slot = jour_plan.get(cren_name)
            if not slot or not isinstance(slot, dict) or not slot.get('open'):
                continue
            mins = ce - cs
            for agents in slot.get('assignment', {}).values():
                for agent in agents:
                    sp_count[agent] += mins

    return result, samedi_type, dict(sp_count)


# ─────────────────────────────────────────────
#  POINT D'ENTRÉE
# ─────────────────────────────────────────────

def compute_full_planning(filepath):
    raw             = load_excel_data(filepath)
    params          = parse_parametres(raw)
    horaires_ouv    = parse_horaires_ouverture(raw)
    affectations    = parse_affectations(raw)
    horaires_agents = parse_horaires_agents(raw)
    besoins_j       = parse_besoins_jeunesse(raw)
    sp_minmax       = parse_sp_minmax(raw)
    roulement       = parse_roulement_samedi(raw)
    creneaux        = parse_creneaux(params)

    mois_str  = str(params.get('Mois', 'janvier')).strip().lower()
    annee     = int(params.get('Année', 2026))
    mois_num  = MOIS_MAP.get(mois_str, 1)

    evenements          = parse_evenements(raw, mois_num, annee)
    raw_blocs           = parse_planning_type(raw)
    planning_type_base  = explode_planning_type_to_creneaux(raw_blocs, creneaux)

    weeks   = get_weeks_of_month(mois_num, annee)
    results = []

    for i, week_dates in enumerate(weeks):
        week_plan, samedi_type, sp_count = plan_week(
            week_num           = i + 1,
            week_dates         = week_dates,
            planning_type_base = planning_type_base,
            affectations       = affectations,
            horaires_agents    = horaires_agents,
            horaires_ouverture = horaires_ouv,
            besoins_jeunesse   = besoins_j,
            sp_minmax          = sp_minmax,
            roulement_samedi   = roulement,
            evenements         = evenements,
            params             = params,
            creneaux           = creneaux,
        )
        results.append({
            'week_num':    i + 1,
            'week_dates':  week_dates,
            'plan':        week_plan,
            'samedi_type': samedi_type,
            'sp_count':    sp_count,
        })

    metadata = {
        'mois':         mois_str,
        'annee':        annee,
        'params':       params,
        'affectations': affectations,
        'creneaux':     creneaux,
        'horaires_ouv': horaires_ouv,
        'roulement':    roulement,
    }

    return results, metadata
