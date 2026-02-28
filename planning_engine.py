"""
Planning Engine v5 — Médiathèque
═══════════════════════════════════════════════════════════════
PRINCIPE FONDAMENTAL : le planning type EST la référence.
L'algo NE TOUCHE à un créneau QUE si :
  1. L'agent du planning type est absent (congé/événement/hors horaires)
  2. L'agent dépasse son max SP hebdo
  3. Une contrainte dure est violée (2h30 consécutif, etc.)
Dans TOUS les autres cas → on garde exactement le planning type.

CONTRAINTES DURES (jamais violées) :
  - Présence réelle de l'agent (horaires)
  - Événements / indisponibilités
  - Sections autorisées par agent
  - 1 agent obligatoire par section RDC/Adulte/MF (sinon ALERTE)
  - Besoins Jeunesse exacts (sinon ALERTE)
  - Max 2h30 consécutif Mar/Jeu/Ven
  - Max 4h consécutif Mer/Sam + pause 1h
  - Roulement samedi ROUGE/BLEU (depuis Paramètres Samedi_S1..S5)
  - Vacataires uniquement Mer/Sam
  - Pas de vacataire seul en Jeunesse 10h-12h et 14h-19h
  - Max SP hebdo STRICT (si dépassé → retirer créneaux non critiques)

CONTRAINTES OPTIMISÉES (satisfaction maximale) :
  - Planning type comme base — minimum de changements
  - Blocs idéaux depuis Paramètres
  - Éviter créneaux ≤1h hors méridienne
  - Max 3 agents différents/section Mar/Jeu/Ven, max 4 Mer/Sam
  - Vacataire = demi-journée minimum
  - ≥2 fermetures Mar-Ven par agent
  - Si min SP non atteint → ALERTE dans récap (pas de recalage auto)
═══════════════════════════════════════════════════════════════
"""

import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
import copy


# ══════════════════════════════════════════════════════════════
#  SECTION 1 — UTILITAIRES
# ══════════════════════════════════════════════════════════════

SECTIONS  = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SP  = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
MOIS_MAP  = {
    'janvier':1,'février':2,'fevrier':2,'mars':3,'avril':4,'mai':5,'juin':6,
    'juillet':7,'août':8,'aout':8,'septembre':9,'octobre':10,'novembre':11,
    'décembre':12,'decembre':12
}
JOURS_TEXTE = {'lundi','mardi','mercredi','jeudi','vendredi','samedi','dimanche'}

def hm_to_min(val):
    """Convertit en minutes depuis minuit. Gère 9h, 9h30, 09:30, timedelta, time."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if hasattr(val, 'hour'):
        return val.hour * 60 + val.minute
    if isinstance(val, timedelta):
        return int(val.total_seconds()) // 60
    if isinstance(val, str):
        v = val.strip()
        if not v or v.lower() == 'nan':
            return None
        norm = v.lower().replace('h', ':').replace('m', '').rstrip(':')
        parts = norm.split(':')
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
    return f"{int(minutes)//60:02d}:{int(minutes)%60:02d}"

def overlap(s1, e1, s2, e2):
    return s1 < e2 and e1 > s2

def is_vacataire(agent):
    return 'vacataire' in str(agent).lower()

def parse_date_flexible(val, annee):
    """Parse date sous toutes ses formes, y compris 'Samedi 7 mars'."""
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
    for j in JOURS_TEXTE:
        if s.startswith(j):
            s = s[len(j):].strip()
            break
    parts = s.replace(',', ' ').split()
    if len(parts) >= 2:
        try:
            jour_num = int(parts[0])
            mois_nom = parts[1].rstrip('.')
            mois_num = MOIS_MAP.get(mois_nom)
            if mois_num:
                an = int(parts[2]) if len(parts) >= 3 else annee
                return datetime(an, mois_num, jour_num)
        except:
            pass
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
    """Découpe 'Anne-Françoise ; Guillaume ; Macha' → ['Anne-Françoise', 'Guillaume', 'Macha']."""
    if not cell_value or str(cell_value).strip() in ('nan', '', '-'):
        return []
    raw = str(cell_value).strip()
    parts = raw.split(';')
    return [p.strip() for p in parts if p.strip() and p.strip().lower() != 'nan']

def parse_duration_param(val, default_min=150):
    """Parse '2h30' ou '4h' en minutes."""
    if not val or str(val).strip() == 'nan':
        return default_min
    s = str(val).strip().lower()
    m = hm_to_min(s)
    return m if m else default_min

def parse_bloc_intervals(raw_str):
    """Parse '10:00-12:30;15:30-17:00;17:00-19:00' → [(600,750),(930,1020),(1020,1140)]."""
    if not raw_str or str(raw_str).strip() == 'nan':
        return []
    result = []
    for part in str(raw_str).split(';'):
        part = part.strip()
        if '-' in part:
            halves = part.split('-')
            s = hm_to_min(halves[0].strip())
            e = hm_to_min(halves[1].strip())
            if s is not None and e is not None:
                result.append((s, e))
    return result


# ══════════════════════════════════════════════════════════════
#  SECTION 2 — LECTURE DU FICHIER EXCEL
# ══════════════════════════════════════════════════════════════

def load_excel_data(filepath):
    xl = pd.ExcelFile(filepath)
    return {sheet: pd.read_excel(filepath, sheet_name=sheet, header=None)
            for sheet in xl.sheet_names}

def parse_parametres(data):
    df = data['Paramètres']
    params = {}
    for _, row in df.iterrows():
        key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        val = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else None
        if key and key not in ('Paramètre', 'nan'):
            params[key] = val
    return params

def parse_horaires_ouverture(data):
    """Retourne dict: jour -> [(start_min, end_min), ...]"""
    df = data['Horaire_ouverture_mediatheque']
    result = {}
    for _, row in df.iterrows():
        jour = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not jour or jour in ('Jour', 'nan', 'WEEKDAY(date,2)'):
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
    """Retourne dict: agent -> [section1, section2, ...]"""
    df = data['Affectations']
    result = {}
    for _, row in df.iterrows():
        agent = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not agent or agent in ('Agent', 'nan'):
            continue
        sections = [str(row.iloc[i]).strip() for i in range(1, len(row))
                    if pd.notna(row.iloc[i]) and str(row.iloc[i]).strip() not in ('nan', '')]
        result[agent] = sections
    return result

def parse_horaires_agents(data):
    """Retourne dict: agent -> jour -> (sm, em, sa, ea) en minutes."""
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
    """Retourne dict: cren_name -> {jour_key: n_agents}"""
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
    """Retourne dict: agent -> {Min_MarVen, Max_MarVen, Min_MarSam, Max_MarSam, SP_Samedi_Type}"""
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
    """Retourne dict: date_str -> [{debut, fin, nom, agents}]"""
    df = data['Événements']
    result = defaultdict(list)
    for _, row in df.iterrows():
        d = parse_date_flexible(row.iloc[0], annee)
        if d is None or d.month != mois or d.year != annee:
            continue
        debut = hm_to_min(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        fin   = hm_to_min(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        nom   = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None
        if not nom or nom == 'nan' or debut is None:
            continue
        agents = []
        for i in range(4, min(12, len(row))):
            if pd.notna(row.iloc[i]):
                agents.extend(split_agents_cell(row.iloc[i]))
        result[d.strftime('%Y-%m-%d')].append({
            'debut': debut,
            'fin':   fin if fin else debut + 60,
            'nom':   nom,
            'agents': agents,
        })
    return dict(result)

def parse_creneaux(params):
    """Retourne [(cren_name, cs, ce), ...]"""
    raw = params.get('Liste_des_créneaux', '')
    if not raw:
        return []
    result = []
    for c in str(raw).split(';'):
        c = c.strip()
        if '-' in c:
            parts = c.split('-')
            s = hm_to_min(parts[0].strip())
            e = hm_to_min(parts[1].strip())
            if s is not None and e is not None:
                result.append((c, s, e))
    return result

def parse_samedi_types(params):
    """
    Lit Samedi_S1..Samedi_S5 dans les paramètres.
    Retourne dict: {1: 'ROUGE', 2: 'BLEU', ...}
    """
    result = {}
    for i in range(1, 6):
        key = f'Samedi_S{i}'
        val = params.get(key)
        if val and str(val).strip().upper() in ('ROUGE', 'BLEU'):
            result[i] = str(val).strip().upper()
    # Fallback : alternance ROUGE/BLEU si non renseigné
    for i in range(1, 6):
        if i not in result:
            result[i] = 'ROUGE' if i % 2 == 1 else 'BLEU'
    return result


# ══════════════════════════════════════════════════════════════
#  SECTION 3 — PARSING DU PLANNING TYPE
# ══════════════════════════════════════════════════════════════

def parse_planning_type(data):
    """
    Lit l'onglet planning_type.
    Retourne dict: jour_key -> [(bloc_label, {section: [agents]})]
    jour_key: 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi_ROUGE', 'Samedi_BLEU'
    """
    df = data['planning_type']
    SECTION_COLS = {
        'RDC':      [2],
        'Adulte':   [3],
        'MF':       [5, 6],
        'Jeunesse': [7, 8, 9],
    }
    JOUR_MAP = {
        'MARDI': 'Mardi', 'MERCREDI': 'Mercredi',
        'JEUDI': 'Jeudi', 'VENDREDI': 'Vendredi',
    }
    SKIP = {'nan', '', 'R D C', 'Adulte', 'Musique & Films',
            'Jeunesse', '12H30', '15H30', 'NaN'}

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
        if not current_jour or not cell1 or cell1 in SKIP:
            continue

        bloc = {s: [] for s in SECTIONS}
        for section, col_indices in SECTION_COLS.items():
            for ci in col_indices:
                if ci < len(row):
                    v = str(row.iloc[ci]).strip() if pd.notna(row.iloc[ci]) else ''
                    if v and v not in SKIP:
                        agent = v.split(' à ')[0].split(' à partir')[0].strip()
                        if agent:
                            bloc[section].append(agent)
        raw_blocs[current_jour].append((cell1, bloc))

    return raw_blocs

def explode_planning_type(raw_blocs, creneaux_def):
    """
    Convertit les blocs du planning type (ex: '10H-12H30')
    en créneaux unitaires.
    Retourne dict: jour_key -> {cren_name -> {section -> [agents]}}
    """
    def parse_bloc_time(label):
        label = label.strip().upper().replace('H', ':')
        if '-' not in label:
            return None, None
        parts = label.split('-')
        def fix(p):
            p = p.strip().rstrip(':')
            return p + ':00' if ':' not in p else p
        return hm_to_min(fix(parts[0])), hm_to_min(fix(parts[1]))

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


# ══════════════════════════════════════════════════════════════
#  SECTION 4 — VÉRIFICATIONS DE DISPONIBILITÉ
# ══════════════════════════════════════════════════════════════

def agent_in_schedule(agent, jour, horaires_agents):
    """Retourne (sm, em, sa, ea) ou None si l'agent ne travaille pas ce jour."""
    return horaires_agents.get(agent, {}).get(jour)

def agent_covers_slot(agent, jour, cs, ce, horaires_agents):
    """True si l'agent est présent (selon ses horaires) sur le créneau cs-ce."""
    h = agent_in_schedule(agent, jour, horaires_agents)
    if h is None:
        return False
    sm, em, sa, ea = h
    in_m = sm is not None and em is not None and cs >= sm and ce <= em
    in_a = sa is not None and ea is not None and cs >= sa and ce <= ea
    return in_m or in_a

def agent_blocked_by_event(agent, cs, ce, date_str, evenements):
    """True si l'agent a un événement qui chevauche le créneau."""
    for ev in evenements.get(date_str, []):
        if agent in ev['agents'] and overlap(cs, ce, ev['debut'], ev['fin']):
            return True
    return False

def agent_available(agent, jour, cs, ce, date_str, horaires_agents, evenements):
    """True si l'agent est présent ET non bloqué par événement."""
    if not agent_covers_slot(agent, jour, cs, ce, horaires_agents):
        return False
    if agent_blocked_by_event(agent, cs, ce, date_str, evenements):
        return False
    return True

def creneau_is_open(cs, ce, jour, horaires_ouverture):
    """True si le créneau est dans les heures d'ouverture."""
    for os, oe in horaires_ouverture.get(jour, []):
        if cs >= os and ce <= oe:
            return True
    return False

def is_meridienne(cs, ce, params):
    """True si le créneau est dans la fenêtre méridienne (exception vacataire/blocs courts)."""
    exc = params.get('Exception_Vacataire_seul', '12:00-14:00')
    parts = str(exc).split('-')
    if len(parts) == 2:
        es = hm_to_min(parts[0].strip())
        ee = hm_to_min(parts[1].strip())
        if es and ee:
            return cs >= es and ce <= ee
    return False


# ══════════════════════════════════════════════════════════════
#  SECTION 5 — SÉLECTION DU REMPLAÇANT
# ══════════════════════════════════════════════════════════════

def find_best_replacement(section, jour, cs, ce, date_str,
                          eligible, affectations, horaires_agents, evenements,
                          agents_used_today_section, sp_count, max_sp_min,
                          exclude=None):
    """
    Cherche le meilleur remplaçant disponible pour une section/créneau.
    Priorité :
      1. Agent déjà utilisé dans cette section aujourd'hui (continuité de bloc)
      2. Agent régulier (non vacataire)
      3. Agent avec le plus de marge SP restante
    Ne retourne jamais un agent qui dépasserait son max SP.
    """
    exclude = set(exclude or [])
    candidates = []
    slot_min = ce - cs

    for agent in eligible:
        if agent in exclude:
            continue
        if section not in affectations.get(agent, []):
            continue
        if not agent_available(agent, jour, cs, ce, date_str, horaires_agents, evenements):
            continue
        # Vérifier marge SP
        current_sp = sp_count.get(agent, 0)
        max_sp = max_sp_min.get(agent, 99 * 60)
        if current_sp + slot_min > max_sp:
            continue
        already_today = agent in agents_used_today_section
        candidates.append((
            not already_today,   # priorité aux agents déjà utilisés (continuité)
            is_vacataire(agent), # réguliers avant vacataires
            current_sp,          # moins de SP = plus disponible
            agent
        ))

    if not candidates:
        return None
    candidates.sort()
    return candidates[0][3]


# ══════════════════════════════════════════════════════════════
#  SECTION 6 — CONTRAINTES CONSÉCUTIVES ET PAUSE
# ══════════════════════════════════════════════════════════════

def get_consecutive_sp(agent, section, cren_idx, creneaux, day_assignments):
    """
    Calcule les minutes SP consécutives de l'agent dans une section
    juste AVANT le créneau d'index cren_idx (sans interruption).
    """
    cumul = 0
    for i in range(cren_idx - 1, -1, -1):
        cn, c_cs, c_ce = creneaux[i]
        slot = day_assignments.get(cn)
        if slot is None:
            break  # créneau fermé = interruption
        assigned = slot.get('assignment', {}).get(section, [])
        if agent in assigned:
            cumul += (c_ce - c_cs)
        else:
            break
    return cumul

def agent_has_pause(agent, cren_idx, creneaux, day_assignments, pause_min=60):
    """
    Vérifie si l'agent a eu une pause d'au moins pause_min minutes
    à un moment quelconque dans la journée (créneau sans SP).
    """
    for i in range(cren_idx):
        cn, c_cs, c_ce = creneaux[i]
        slot = day_assignments.get(cn)
        if slot is None:
            return True  # créneau fermé = pause
        all_assigned = [a for s in SECTIONS
                        for a in slot.get('assignment', {}).get(s, [])]
        if agent not in all_assigned:
            # Calculer la durée de non-SP
            gap = c_ce - c_cs
            if gap >= pause_min:
                return True
    return False


# ══════════════════════════════════════════════════════════════
#  SECTION 7 — PLANIFICATION D'UNE SEMAINE
# ══════════════════════════════════════════════════════════════

def get_samedi_agents(roulement_samedi, samedi_type, affectations, horaires_agents):
    """Agents éligibles pour un samedi ROUGE ou BLEU (+ vacataires)."""
    agents = [a for a, r in roulement_samedi.items()
              if r == samedi_type and horaires_agents.get(a, {}).get('Samedi') is not None]
    for agent in affectations:
        if is_vacataire(agent) and agent not in agents \
           and horaires_agents.get(agent, {}).get('Samedi') is not None:
            agents.append(agent)
    return agents

def plan_week(week_num, week_dates, planning_type_base, samedi_type,
              affectations, horaires_agents, horaires_ouverture,
              besoins_jeunesse, sp_minmax, roulement_samedi,
              evenements, params, creneaux):
    """
    Calcule le planning d'une semaine.
    PRINCIPE : planning type = référence. On ne modifie que si nécessaire.
    """
    # ── Paramètres ──
    max_consec_court = parse_duration_param(params.get('Durée_SP_max_idéale', '2h30'), 150)
    max_consec_long  = parse_duration_param(params.get('Durée_SP_max_tolérée', '4h'), 240)
    vac_days         = str(params.get('Mode_vacataires', 'mercredi,samedi')).lower()
    exc_vac_start    = hm_to_min(str(params.get('Exception_Vacataire_seul', '12:00')).split('-')[0]) or 720
    exc_vac_end      = hm_to_min(str(params.get('Exception_Vacataire_seul', '14:00')).split('-')[-1]) or 840

    max_agents_day = {'Mardi': 3, 'Mercredi': 4, 'Jeudi': 3, 'Vendredi': 3, 'Samedi': 4}
    all_agents = list(affectations.keys())

    # ── Convertir max SP en minutes ──
    sp_max_min = {}
    sp_min_min = {}
    for agent, mm in sp_minmax.items():
        is_samedi_week = 'Samedi' in week_dates
        if is_samedi_week:
            sp_max_min[agent] = mm['Max_MarSam'] * 60
            sp_min_min[agent] = mm['Min_MarSam'] * 60
        else:
            sp_max_min[agent] = mm['Max_MarVen'] * 60
            sp_min_min[agent] = mm['Min_MarVen'] * 60

    result = {}
    # sp_count en minutes sur la semaine
    sp_count = defaultdict(int)

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        date_str = date.strftime('%Y-%m-%d')

        # ── Agents éligibles ce jour ──
        if jour == 'Samedi':
            eligible = get_samedi_agents(roulement_samedi, samedi_type,
                                          affectations, horaires_agents)
        elif jour.lower() in vac_days:
            eligible = [a for a in all_agents
                        if horaires_agents.get(a, {}).get(jour) is not None]
        else:
            eligible = [a for a in all_agents if not is_vacataire(a)
                        and horaires_agents.get(a, {}).get(jour) is not None]

        # ── Base : planning type pour ce jour ──
        pt_key = f"Samedi_{samedi_type}" if jour == 'Samedi' else jour
        base   = planning_type_base.get(pt_key, {})
        max_rot = max_agents_day.get(jour, 4)

        agents_used_today  = {s: [] for s in SECTIONS}
        day_assignments    = {}
        alerts             = {}   # cren_name -> [messages d'alerte]

        result[jour] = {'_samedi_type': samedi_type if jour == 'Samedi' else None}

        for cren_idx, (cren_name, cs, ce) in enumerate(creneaux):
            slot_min = ce - cs

            # ── Créneau fermé ? ──
            if not creneau_is_open(cs, ce, jour, horaires_ouverture):
                result[jour][cren_name] = None
                continue

            # ── Partir du planning type ──
            base_slot  = base.get(cren_name, {s: [] for s in SECTIONS})
            assignment = {s: list(base_slot.get(s, [])) for s in SECTIONS}

            slot_alerts = []

            # ══ PASSE A : Vérifier chaque agent du planning type ══
            for section in SECTIONS:
                final_agents = []
                for agent in assignment[section]:

                    # Normaliser vacataires génériques
                    if agent.upper() in ('VACATAIRES', 'VACATAIRE'):
                        vac_found = next(
                            (a for a in eligible if is_vacataire(a)
                             and agent_available(a, jour, cs, ce, date_str,
                                                  horaires_agents, evenements)
                             and a not in final_agents
                             and sp_count.get(a, 0) + slot_min <= sp_max_min.get(a, 99*60)),
                            None
                        )
                        if vac_found:
                            final_agents.append(vac_found)
                        continue

                    # Raisons de toucher au planning type :
                    must_replace = False
                    reason = ""

                    # 1. Agent absent (hors horaires ou événement)
                    if not agent_available(agent, jour, cs, ce, date_str,
                                           horaires_agents, evenements):
                        must_replace = True
                        reason = "absent/événement"

                    # 2. Agent dépasse son max SP
                    elif sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
                        must_replace = True
                        reason = "max SP atteint"

                    # 3. Contrainte consécutif
                    elif jour in ('Mardi', 'Jeudi', 'Vendredi'):
                        consec = get_consecutive_sp(agent, section, cren_idx,
                                                     creneaux, day_assignments)
                        if consec + slot_min > max_consec_court:
                            must_replace = True
                            reason = f"max {max_consec_court//60}h{max_consec_court%60:02d} consécutif"
                    else:  # Mercredi / Samedi
                        consec = get_consecutive_sp(agent, section, cren_idx,
                                                     creneaux, day_assignments)
                        # Pause obligatoire 1h après 4h
                        if consec >= max_consec_long and not agent_has_pause(
                                agent, cren_idx, creneaux, day_assignments):
                            must_replace = True
                            reason = "pause 1h obligatoire"
                        elif consec + slot_min > max_consec_long:
                            must_replace = True
                            reason = f"max {max_consec_long//60}h consécutif"

                    if not must_replace:
                        # Garder l'agent du planning type — aucune raison de le changer
                        final_agents.append(agent)
                    else:
                        # Chercher remplaçant
                        excl = set(final_agents) | set(
                            a for s2 in SECTIONS
                            for a in assignment.get(s2, []) if s2 != section)
                        repl = find_best_replacement(
                            section=section, jour=jour, cs=cs, ce=ce,
                            date_str=date_str, eligible=eligible,
                            affectations=affectations,
                            horaires_agents=horaires_agents,
                            evenements=evenements,
                            agents_used_today_section=agents_used_today[section],
                            sp_count=sp_count,
                            max_sp_min=sp_max_min,
                            exclude=excl,
                        )
                        if repl:
                            final_agents.append(repl)
                        # Si pas de remplaçant → sera traité en passe B

                assignment[section] = final_agents

            # ══ PASSE B : S'assurer que chaque section a 1 agent (contrainte dure) ══
            for section in ['RDC', 'Adulte', 'MF']:
                if not assignment[section]:
                    excl = set(a for s in SECTIONS for a in assignment[s])
                    repl = find_best_replacement(
                        section=section, jour=jour, cs=cs, ce=ce,
                        date_str=date_str, eligible=eligible,
                        affectations=affectations,
                        horaires_agents=horaires_agents,
                        evenements=evenements,
                        agents_used_today_section=agents_used_today[section],
                        sp_count=sp_count,
                        max_sp_min=sp_max_min,
                        exclude=excl,
                    )
                    if repl:
                        assignment[section] = [repl]
                    else:
                        slot_alerts.append(f"ALERTE — {section} : aucun agent disponible")

            # ══ PASSE C : Besoins Jeunesse ══
            bj = besoins_jeunesse.get(cren_name, {})
            if jour == 'Samedi':
                j_key = f'Samedi_{samedi_type.lower()}'
                j_needs = int(bj.get(j_key, bj.get('Samedi_rouge', 1)))
            else:
                j_needs = int(bj.get(jour, 1))
            j_needs = max(j_needs, 0)

            current_j_count = len(assignment.get('Jeunesse', []))
            if current_j_count < j_needs:
                occupied = set(a for s in SECTIONS for a in assignment[s])
                for agent in eligible:
                    if len(assignment['Jeunesse']) >= j_needs:
                        break
                    if agent in occupied:
                        continue
                    if 'Jeunesse' not in affectations.get(agent, []):
                        continue
                    if not agent_available(agent, jour, cs, ce, date_str,
                                           horaires_agents, evenements):
                        continue
                    if sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
                        continue
                    assignment['Jeunesse'].append(agent)
                    if agent not in agents_used_today['Jeunesse']:
                        agents_used_today['Jeunesse'].append(agent)

            if len(assignment['Jeunesse']) < j_needs:
                slot_alerts.append(
                    f"ALERTE — Jeunesse : {len(assignment['Jeunesse'])}/{j_needs} agents")

            # ══ PASSE D : Contrainte vacataire seul en Jeunesse ══
            j_agents = assignment.get('Jeunesse', [])
            if j_agents and all(is_vacataire(a) for a in j_agents):
                # Hors fenêtre méridienne : interdit
                if not (cs >= exc_vac_start and ce <= exc_vac_end):
                    occupied = set(a for s in SECTIONS for a in assignment[s]
                                   if a not in j_agents)
                    repl = find_best_replacement(
                        section='Jeunesse', jour=jour, cs=cs, ce=ce,
                        date_str=date_str, eligible=eligible,
                        affectations=affectations,
                        horaires_agents=horaires_agents,
                        evenements=evenements,
                        agents_used_today_section=agents_used_today['Jeunesse'],
                        sp_count=sp_count,
                        max_sp_min=sp_max_min,
                        exclude=occupied,
                    )
                    if repl and not is_vacataire(repl):
                        assignment['Jeunesse'].insert(0, repl)
                    elif not repl:
                        slot_alerts.append("ALERTE — Jeunesse : vacataire seul hors exception")

            # ══ PASSE E : Limite rotation agents par section ══
            for section in SECTIONS:
                adjusted = []
                for agent in assignment[section]:
                    unique_reg = [a for a in agents_used_today[section]
                                 if not is_vacataire(a)]
                    if (not is_vacataire(agent)
                            and agent not in agents_used_today[section]
                            and len(unique_reg) >= max_rot):
                        # Réutiliser dernier agent régulier dispo
                        last = next(
                            (a for a in reversed(agents_used_today[section])
                             if not is_vacataire(a)
                             and agent_available(a, jour, cs, ce, date_str,
                                                  horaires_agents, evenements)
                             and sp_count.get(a, 0) + slot_min
                                <= sp_max_min.get(a, 99*60)),
                            None
                        )
                        adjusted.append(last if last else agent)
                    else:
                        adjusted.append(agent)
                assignment[section] = adjusted

            # ══ Mettre à jour les compteurs ══
            for section in SECTIONS:
                for agent in assignment[section]:
                    sp_count[agent] += slot_min
                    if agent not in agents_used_today[section]:
                        agents_used_today[section].append(agent)

            # Événements du créneau
            slot_events = [ev for ev in evenements.get(date_str, [])
                           if overlap(cs, ce, ev['debut'], ev['fin'])]

            day_assignments[cren_name] = {
                'assignment': assignment,
                'events':     slot_events,
                'alerts':     slot_alerts,
                'open':       True,
            }
            result[jour][cren_name] = day_assignments[cren_name]

    # ── Vérification min SP en fin de semaine ──
    sp_alerts = {}
    for agent, mm in sp_minmax.items():
        is_samedi_week = 'Samedi' in week_dates
        min_sp = (mm['Min_MarSam'] if is_samedi_week else mm['Min_MarVen']) * 60
        actual = sp_count.get(agent, 0)
        if actual < min_sp:
            sp_alerts[agent] = (
                f"SP insuffisant : {actual//60}h{actual%60:02d} "
                f"/ min {int(min_sp)//60}h{int(min_sp)%60:02d}"
            )

    return result, samedi_type, dict(sp_count), sp_alerts


# ══════════════════════════════════════════════════════════════
#  SECTION 8 — CALENDRIER
# ══════════════════════════════════════════════════════════════

def get_weeks_of_month(mois, annee):
    """Retourne liste de dicts {jour: date} pour chaque semaine du mois."""
    first_day = datetime(annee, mois, 1)
    days_to_tuesday = (1 - first_day.weekday()) % 7
    first_tuesday   = first_day + timedelta(days=days_to_tuesday)
    offsets = {'Mardi': 0, 'Mercredi': 1, 'Jeudi': 2, 'Vendredi': 3, 'Samedi': 4}
    weeks = []
    current = first_tuesday
    while current.month == mois:
        week = {}
        for jour, off in offsets.items():
            d = current + timedelta(days=off)
            if d.month == mois:
                week[jour] = d
        if week:
            weeks.append(week)
        current += timedelta(weeks=1)
    return weeks


# ══════════════════════════════════════════════════════════════
#  SECTION 9 — POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════

def compute_full_planning(filepath):
    """
    Point d'entrée principal.
    Lit le fichier Excel, calcule le planning pour chaque semaine.
    Retourne (weeks_data, metadata).
    """
    raw             = load_excel_data(filepath)
    params          = parse_parametres(raw)
    horaires_ouv    = parse_horaires_ouverture(raw)
    affectations    = parse_affectations(raw)
    horaires_agents = parse_horaires_agents(raw)
    besoins_j       = parse_besoins_jeunesse(raw)
    sp_minmax       = parse_sp_minmax(raw)
    roulement       = parse_roulement_samedi(raw)
    creneaux        = parse_creneaux(params)
    samedi_types    = parse_samedi_types(params)

    mois_str  = str(params.get('Mois', 'janvier')).strip().lower()
    annee     = int(params.get('Année', 2026))
    mois_num  = MOIS_MAP.get(mois_str, 1)

    evenements         = parse_evenements(raw, mois_num, annee)
    raw_blocs          = parse_planning_type(raw)
    planning_type_base = explode_planning_type(raw_blocs, creneaux)

    weeks   = get_weeks_of_month(mois_num, annee)
    results = []

    for i, week_dates in enumerate(weeks):
        week_num    = i + 1
        samedi_type = samedi_types.get(week_num, 'ROUGE' if week_num % 2 == 1 else 'BLEU')

        week_plan, sam_type, sp_count, sp_alerts = plan_week(
            week_num           = week_num,
            week_dates         = week_dates,
            planning_type_base = planning_type_base,
            samedi_type        = samedi_type,
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
            'week_num':    week_num,
            'week_dates':  week_dates,
            'plan':        week_plan,
            'samedi_type': sam_type,
            'sp_count':    sp_count,
            'sp_alerts':   sp_alerts,
        })

    metadata = {
        'mois':         mois_str,
        'annee':        annee,
        'params':       params,
        'affectations': affectations,
        'creneaux':     creneaux,
        'horaires_ouv': horaires_ouv,
        'roulement':    roulement,
        'sp_minmax':    sp_minmax,
    }

    return results, metadata
