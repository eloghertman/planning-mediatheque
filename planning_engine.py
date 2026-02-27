"""
Planning Engine — Médiathèque
Calcule automatiquement le planning SP à partir des variables Excel
"""

import pandas as pd
from datetime import datetime, timedelta
import calendar
from collections import defaultdict


# ─────────────────────────────────────────────
#  UTILITAIRES HORAIRES
# ─────────────────────────────────────────────

def hm_to_min(val):
    """Convertit un objet time, timedelta, string 'HH:MM' ou 'HH:MM:SS' en minutes."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if hasattr(val, 'hour'):          # datetime.time
        return val.hour * 60 + val.minute
    if isinstance(val, timedelta):    # timedelta (Excel stocke parfois comme ça)
        total = int(val.total_seconds())
        return total // 60
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return None
        parts = val.split(':')
        return int(parts[0]) * 60 + int(parts[1])
    return None

def min_to_hhmm(minutes):
    if minutes is None:
        return ""
    return f"{minutes//60:02d}:{minutes%60:02d}"

def overlap(s1, e1, s2, e2):
    """True si les deux intervalles se chevauchent."""
    return s1 < e2 and e1 > s2


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
    """Retourne dict: jour -> [(start_min, end_min), ...]"""
    df = data['Horaire_ouverture_mediatheque']
    result = {}
    for _, row in df.iterrows():
        jour = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not jour or jour in ('Jour', 'nan'):
            continue
        slots = []
        for i in range(2, len(row)-1, 2):
            s = hm_to_min(row.iloc[i]) if pd.notna(row.iloc[i]) else None
            e = hm_to_min(row.iloc[i+1]) if pd.notna(row.iloc[i+1]) else None
            if s is not None and e is not None and s != e:
                slots.append((s, e))
        if slots:
            result[jour] = slots
    return result

def parse_affectations(data):
    """Retourne dict: agent -> [section1, section2, ...]"""
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
    """Retourne dict: agent -> jour -> (start_matin, end_matin, start_apm, end_apm)"""
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
    """Retourne dict: creneau -> {'Mardi': n, 'Mercredi': n, ..., 'Samedi_rouge': n, 'samedi bleu': n}"""
    df = data['Besoins_Jeunesse']
    result = {}
    cols = [str(c).strip() for c in df.iloc[0]]  # header row
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
    """Retourne dict: agent -> {Min_MarVen, Max_MarVen, Min_MarSam, Max_MarSam, SP_Samedi_Type}"""
    df = data['SP_MinMax']
    result = {}
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not agent or agent == 'Agent':
            continue
        result[agent] = {
            'Min_MarVen':      row.iloc[1] if pd.notna(row.iloc[1]) else 0,
            'Max_MarVen':      row.iloc[2] if pd.notna(row.iloc[2]) else 99,
            'Min_MarSam':      row.iloc[3] if pd.notna(row.iloc[3]) else 0,
            'Max_MarSam':      row.iloc[4] if pd.notna(row.iloc[4]) else 99,
            'SP_Samedi_Type':  row.iloc[5] if pd.notna(row.iloc[5]) else 0,
        }
    return result

def parse_roulement_samedi(data):
    """Retourne dict: agent -> 'BLEU' | 'ROUGE'"""
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

def parse_evenements(data, mois, annee):
    """
    Retourne dict: date_str ('2026-01-27') -> list of
      {'debut': min, 'fin': min, 'nom': str, 'agents': [str]}
    """
    df = data['Événements']
    result = defaultdict(list)
    agent_cols = [i for i, c in enumerate(df.iloc[0]) if str(c).strip().lower() in ('agents','agent','agents evenemtns','agents evenements') or (i >= 4 and i <= 9)]

    for _, row in df.iterrows():
        date_val = row.iloc[0]
        if not pd.notna(date_val):
            continue
        try:
            if isinstance(date_val, (datetime,)):
                d = date_val
            else:
                d = pd.to_datetime(date_val)
            if d.month != mois or d.year != annee:
                continue
        except:
            continue

        debut = hm_to_min(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        fin   = hm_to_min(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        nom   = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None

        if not nom or nom == 'nan' or debut is None:
            continue

        agents = []
        for i in range(4, min(11, len(row))):
            v = row.iloc[i]
            if pd.notna(v) and str(v).strip() not in ('nan', ''):
                agents.append(str(v).strip())

        date_str = d.strftime('%Y-%m-%d')
        result[date_str].append({
            'debut': debut,
            'fin':   fin if fin else debut + 60,
            'nom':   nom,
            'agents': agents,
        })
    return dict(result)

def parse_creneaux(params):
    """Parse la liste des créneaux depuis les paramètres."""
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
#  CALCUL DES SEMAINES DU MOIS
# ─────────────────────────────────────────────

JOURS_FR = ['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
JOURS_SP = ['Mardi','Mercredi','Jeudi','Vendredi','Samedi']

def get_weeks_of_month(mois, annee):
    """
    Retourne une liste de semaines, chaque semaine = dict:
      { 'Mardi': date, 'Mercredi': date, 'Jeudi': date, 'Vendredi': date, 'Samedi': date }
    """
    # Premier jour du mois
    first_day = datetime(annee, mois, 1)
    # Trouver le premier mardi
    weekday = first_day.weekday()  # 0=lundi, 1=mardi
    days_to_tuesday = (1 - weekday) % 7
    first_tuesday = first_day + timedelta(days=days_to_tuesday)

    weeks = []
    current = first_tuesday
    while current.month == mois:
        week = {}
        for i, jour in enumerate(JOURS_SP):
            # Mardi=0, Mercredi=1, Jeudi=2, Vendredi=3, Samedi=4
            offsets = {'Mardi':0,'Mercredi':1,'Jeudi':2,'Vendredi':3,'Samedi':4}
            d = current + timedelta(days=offsets[jour])
            if d.month == mois:
                week[jour] = d
        if week:
            weeks.append(week)
        current += timedelta(weeks=1)
    return weeks


# ─────────────────────────────────────────────
#  MOTEUR DE PLANIFICATION
# ─────────────────────────────────────────────

SECTIONS = ['RDC', 'Adulte', 'MF', 'Jeunesse']

def is_vacataire(agent):
    return 'vacataire' in agent.lower()

def agent_working(agent, jour, horaires_agents):
    """Retourne (sm, em, sa, ea) ou None si l'agent ne travaille pas ce jour."""
    return horaires_agents.get(agent, {}).get(jour)

def agent_blocked_by_event(agent, cs, ce, date_str, evenements):
    """True si l'agent est bloqué par un événement sur ce créneau."""
    evts = evenements.get(date_str, [])
    for ev in evts:
        if agent in ev['agents']:
            if overlap(cs, ce, ev['debut'], ev['fin']):
                return True, ev['nom']
    return False, None

def agent_available_for_sp(agent, jour, cs, ce, date_str, horaires_agents, evenements):
    """Vérifie si l'agent peut faire du SP sur ce créneau."""
    h = agent_working(agent, jour, horaires_agents)
    if h is None:
        return False
    sm, em, sa, ea = h
    in_matin = sm is not None and em is not None and cs >= sm and ce <= em
    in_apm   = sa is not None and ea is not None and cs >= sa and ce <= ea
    if not (in_matin or in_apm):
        return False
    blocked, _ = agent_blocked_by_event(agent, cs, ce, date_str, evenements)
    return not blocked

def creneau_is_open(cs, ce, jour, horaires_ouverture):
    for (os, oe) in horaires_ouverture.get(jour, []):
        if cs >= os and ce <= oe:
            return True
    return False

def get_samedi_type(week_dates, roulement_samedi):
    """
    Détermine si le samedi de cette semaine est ROUGE ou BLEU.
    On compte le nombre de samedis depuis le début de l'année et on alterne.
    Convention : semaine 1 = ROUGE si le premier samedi du mois est ROUGE pour les agents ROUGE.
    On regarde juste la parité de la semaine dans le mois.
    """
    # On retourne ROUGE pour les semaines impaires, BLEU pour les paires
    # (par défaut — l'utilisateur peut overrider dans les paramètres)
    return 'ROUGE'  # sera calculé dynamiquement par numéro de semaine

def compute_samedi_type_by_week(week_num):
    """Semaine 1 et 3 = ROUGE, semaine 2 et 4 = BLEU (convention standard)."""
    return 'ROUGE' if week_num % 2 == 1 else 'BLEU'

def get_agents_for_samedi(roulement_samedi, samedi_type, affectations):
    """Retourne les agents qui travaillent ce samedi (selon ROUGE/BLEU)."""
    agents = []
    for agent, roul in roulement_samedi.items():
        if roul == samedi_type:
            agents.append(agent)
    # Ajouter les vacataires (ils travaillent tous les samedis)
    for agent in affectations:
        if is_vacataire(agent) and agent not in agents:
            agents.append(agent)
    return agents

def exception_vacataire_seul(cs, ce, params):
    """True si ce créneau est dans la fenêtre d'exception vacataire seul (ex: 12h-14h)."""
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

def has_regular_agent(agents_assigned, affectations):
    """Vérifie qu'il y a au moins un agent régulier (non vacataire) parmi les assignés."""
    for a in agents_assigned:
        if not is_vacataire(a):
            return True
    return False

def compute_jeunesse_needs(cren_name, jour, samedi_type, besoins_jeunesse):
    """Retourne le nombre d'agents Jeunesse nécessaires pour ce créneau."""
    bj = besoins_jeunesse.get(cren_name, {})
    if jour == 'Samedi':
        key = 'Samedi_rouge' if samedi_type == 'ROUGE' else 'samedi bleu'
        return bj.get(key, bj.get('Samedi_rouge', 1))
    return bj.get(jour, 1)

def minutes_of_cren(cs, ce):
    return ce - cs


# ─────────────────────────────────────────────
#  ALGORITHME PRINCIPAL DE PLANIFICATION
# ─────────────────────────────────────────────

def plan_week(
    week_num, week_dates,
    affectations, horaires_agents, horaires_ouverture,
    besoins_jeunesse, sp_minmax, roulement_samedi,
    evenements, params, creneaux
):
    """
    Calcule le planning d'une semaine.
    Retourne dict: jour -> creneau_name -> { section -> [agents], 'events': [...] }
    """
    samedi_type = compute_samedi_type_by_week(week_num)
    samedi_agents = get_agents_for_samedi(roulement_samedi, samedi_type, affectations)

    all_agents = list(affectations.keys())
    vacataires = [a for a in all_agents if is_vacataire(a)]
    regular_agents = [a for a in all_agents if not is_vacataire(a)]

    # Compteur SP par agent (en minutes)
    sp_count = defaultdict(int)

    result = {}

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        date_str = date.strftime('%Y-%m-%d')

        # Agents disponibles ce jour
        if jour == 'Samedi':
            eligible = samedi_agents
        else:
            eligible = regular_agents  # pas de vacataires Mardi-Vendredi

        # Vacataires uniquement mercredi et samedi
        vac_days = str(params.get('Mode_vacataires', 'mercredi,samedi')).lower()
        if jour.lower() in vac_days:
            eligible = all_agents
        else:
            eligible = regular_agents

        result[jour] = {'_samedi_type': samedi_type if jour == 'Samedi' else None}

        for cren_name, cs, ce in creneaux:
            # Vérifier si ouvert
            if not creneau_is_open(cs, ce, jour, horaires_ouverture):
                result[jour][cren_name] = None  # fermé
                continue

            # Récupérer les événements du créneau
            slot_events = []
            for ev in evenements.get(date_str, []):
                if overlap(cs, ce, ev['debut'], ev['fin']):
                    slot_events.append(ev)

            # Besoins Jeunesse
            j_needs = compute_jeunesse_needs(cren_name, jour, samedi_type, besoins_jeunesse)

            # Assignation par section
            assignment = {s: [] for s in SECTIONS}

            # Pour chaque section, trouver les agents disponibles
            for section in SECTIONS:
                if section == 'Jeunesse':
                    continue  # traité séparément

                # Agents qui peuvent couvrir cette section ce créneau
                candidates = []
                for agent in eligible:
                    if section not in affectations.get(agent, []):
                        continue
                    if not agent_available_for_sp(agent, jour, cs, ce, date_str,
                                                   horaires_agents, evenements):
                        continue
                    # Ne pas mettre le même agent dans 2 sections
                    already_assigned = any(agent in assignment[s] for s in SECTIONS)
                    if already_assigned:
                        continue
                    candidates.append(agent)

                if candidates:
                    # Prioriser : agent avec moins de SP d'abord, non vacataire en premier
                    candidates.sort(key=lambda a: (is_vacataire(a), sp_count[a]))
                    chosen = candidates[0]
                    assignment[section].append(chosen)

            # Section Jeunesse
            j_candidates = []
            for agent in eligible:
                if 'Jeunesse' not in affectations.get(agent, []):
                    continue
                if not agent_available_for_sp(agent, jour, cs, ce, date_str,
                                               horaires_agents, evenements):
                    continue
                already_assigned = any(agent in assignment[s] for s in ['RDC', 'Adulte', 'MF'])
                if already_assigned:
                    continue
                j_candidates.append(agent)

            j_candidates.sort(key=lambda a: (is_vacataire(a), sp_count[a]))

            for agent in j_candidates[:max(j_needs, 1)]:
                assignment['Jeunesse'].append(agent)

            # Règle vacataire seul
            is_exception = exception_vacataire_seul(cs, ce, params)
            for section in SECTIONS:
                agents_in_section = assignment[section]
                if not agents_in_section:
                    continue
                all_vac = all(is_vacataire(a) for a in agents_in_section)
                if all_vac and not is_exception:
                    # Chercher un agent régulier pour accompagner
                    for agent in eligible:
                        if is_vacataire(agent):
                            continue
                        if section not in affectations.get(agent, []):
                            continue
                        if not agent_available_for_sp(agent, jour, cs, ce, date_str,
                                                       horaires_agents, evenements):
                            continue
                        already = any(agent in assignment[s] for s in SECTIONS)
                        if not already:
                            assignment[section].insert(0, agent)
                            break

            # Mettre à jour le compteur SP
            for section, agents in assignment.items():
                for agent in agents:
                    sp_count[agent] += minutes_of_cren(cs, ce)

            result[jour][cren_name] = {
                'assignment': assignment,
                'events': slot_events,
                'open': True,
            }

    return result, samedi_type, sp_count


def compute_full_planning(filepath):
    """
    Point d'entrée principal.
    Lit le fichier Excel, calcule le planning pour chaque semaine du mois.
    Retourne: weeks_data list, params, metadata
    """
    raw = load_excel_data(filepath)

    params           = parse_parametres(raw)
    horaires_ouv     = parse_horaires_ouverture(raw)
    affectations     = parse_affectations(raw)
    horaires_agents  = parse_horaires_agents(raw)
    besoins_j        = parse_besoins_jeunesse(raw)
    sp_minmax        = parse_sp_minmax(raw)
    roulement        = parse_roulement_samedi(raw)
    creneaux         = parse_creneaux(params)

    mois_str  = str(params.get('Mois', 'janvier')).strip().lower()
    annee     = int(params.get('Année', 2026))

    MOIS_MAP = {
        'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
        'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,'décembre':12
    }
    mois_num = MOIS_MAP.get(mois_str, 1)

    evenements = parse_evenements(raw, mois_num, annee)
    weeks      = get_weeks_of_month(mois_num, annee)

    results = []
    for i, week_dates in enumerate(weeks):
        week_plan, samedi_type, sp_count = plan_week(
            week_num        = i + 1,
            week_dates      = week_dates,
            affectations    = affectations,
            horaires_agents = horaires_agents,
            horaires_ouverture = horaires_ouv,
            besoins_jeunesse   = besoins_j,
            sp_minmax       = sp_minmax,
            roulement_samedi   = roulement,
            evenements      = evenements,
            params          = params,
            creneaux        = creneaux,
        )
        results.append({
            'week_num':    i + 1,
            'week_dates':  week_dates,
            'plan':        week_plan,
            'samedi_type': samedi_type,
            'sp_count':    dict(sp_count),
        })

    metadata = {
        'mois':      mois_str,
        'annee':     annee,
        'params':    params,
        'affectations': affectations,
        'creneaux':  creneaux,
        'horaires_ouv': horaires_ouv,
        'roulement': roulement,
    }

    return results, metadata
