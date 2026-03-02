"""
Planning Engine v10 — Médiathèque
════════════════════════════════════════════════════════════════
CONTRAINTES DURES (D1-D16) :
  D1.  Présence réelle (Horaires_Des_Agents)
  D2.  Événements / indisponibilités
  D3.  Sections autorisées + max 1 agent/section/créneau (sauf Jeunesse)
  D4.  Besoins Jeunesse exacts (ni plus ni moins) — tableau Vacances/HorsVacances
  D5.  1 agent min RDC/Adulte/MF par créneau ouvert → D16 si impossible
  D6.  Min/Max SP hebdo (SP_MinMax par semaine)
  D7.  Consécutif Mar/Jeu/Ven : 2h30 max, 4h si congé dans la section
  D8.  Consécutif Mer/Sam : 4h max + pause 1h obligatoire après 4h
  D9.  Roulement Samedi par semaine (Roulement_Samedi SEMAINE_1…)
  D10. Vacataires uniquement jours autorisés (Mode_vacataires)
  D11. Pas de vacataire seul dans une section hors fenêtre 12h-14h
  D12. Max 3 agents réguliers différents/section/jour Mar/Jeu/Ven
       Max 4 agents réguliers différents/section/jour Mer/Sam
  D13. Un agent = une seule section par créneau
  D14. Pause méridienne ≥1h continue entre 12h-14h — réguliers uniquement
  D15. (fusionné dans O6) Vacataire min 6h si matin+après-midi
  D16. Agent hors section habituelle en dernier recours absolu (cellule rouge)

CONTRAINTES OPTIMISÉES (O1-O9) :
  O1.  Planning type = référence (minimum de changements)
  O2.  Section primaire prioritaire; exception vacataires = comblent les trous
  O3.  Répartition équitable SP sur la semaine
  O4.  Répartir heures SP dans la journée
  O5.  Éviter créneaux ≤1h sauf méridienne
  O6.  Règles vacataires : Vac1 prioritaire, relais dès 2h30 régulier,
       pause méridienne, min 6h si matin+après-midi, max 7h
  O7.  Éviter >2h30 consécutif (déprioritiser)
  O8.  Recalage final min SP
  O9.  Équilibrage section/jour (priorité très basse)
════════════════════════════════════════════════════════════════
"""

import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict

SECTIONS    = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SP    = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
JOURS_COURT = {'Mardi': 3, 'Mercredi': 4, 'Jeudi': 3, 'Vendredi': 3, 'Samedi': 4}
MOIS_MAP    = {
    'janvier':1,'février':2,'fevrier':2,'mars':3,'avril':4,'mai':5,'juin':6,
    'juillet':7,'août':8,'aout':8,'septembre':9,'octobre':10,'novembre':11,
    'décembre':12,'decembre':12
}
JOURS_TEXTE = {'lundi','mardi','mercredi','jeudi','vendredi','samedi','dimanche'}


# ══════════════════════════════════════════════════════════════
#  UTILITAIRES
# ══════════════════════════════════════════════════════════════

def hm_to_min(val):
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
        except Exception:
            return None
    return None

def overlap(s1, e1, s2, e2):
    return s1 < e2 and e1 > s2

def is_vacataire(agent):
    return 'vacataire' in str(agent).lower()

def normalize_agent_name(name):
    """Normalise les noms d'agents (ex: Vacataire2 → Vacataire 2)."""
    if not name or str(name).strip() in ('nan', ''):
        return None
    n = str(name).strip()
    import re
    n = re.sub(r'(?i)vacataire(\d)', r'Vacataire \1', n)
    return n

def parse_date_flexible(val, annee):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val
    try:
        ts = pd.Timestamp(val)
        if not pd.isna(ts):
            return ts.to_pydatetime()
    except Exception:
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
        except Exception:
            pass
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d/%m'):
        try:
            d = datetime.strptime(s, fmt)
            if d.year == 1900:
                d = d.replace(year=annee)
            return d
        except Exception:
            pass
    return None

def split_agents_cell(cell_value):
    if not cell_value or str(cell_value).strip() in ('nan', '', '-'):
        return []
    parts = str(cell_value).strip().split(';')
    result = []
    for p in parts:
        n = normalize_agent_name(p)
        if n:
            result.append(n)
    return result

def parse_duration_param(val, default_min=150):
    if not val or str(val).strip() == 'nan':
        return default_min
    m = hm_to_min(str(val).strip().lower())
    return m if m else default_min


# ══════════════════════════════════════════════════════════════
#  LECTURE EXCEL
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

def parse_semaines_type(params):
    """
    Retourne dict: {1: 'Vacances Scolaires', 2: 'Hors Vacances scolaires', ...}
    Clés: Semaine_1 à Semaine_5 dans Paramètres.
    """
    result = {}
    for i in range(1, 7):
        val = params.get(f'Semaine_{i}')
        if val:
            v = str(val).strip().lower()
            if 'vacances' in v and 'hors' not in v:
                result[i] = 'Vacances Scolaires'
            else:
                result[i] = 'Hors Vacances scolaires'
    return result

def parse_horaires_ouverture(data):
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
    """Retourne dict: agent -> [section1, section2, ...] (ordre = priorité)."""
    df = data['Affectations']
    result = {}
    for _, row in df.iterrows():
        agent = normalize_agent_name(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        if not agent or agent in ('Agent', 'nan'):
            continue
        sections = []
        for i in range(1, len(row)):
            v = str(row.iloc[i]).strip() if pd.notna(row.iloc[i]) else ''
            if v and v not in ('nan', ''):
                sections.append(v)
        result[agent] = sections
    return result

def parse_horaires_agents(data):
    df = data['Horaires_Des_Agents']
    result = defaultdict(dict)
    for _, row in df.iterrows():
        agent = normalize_agent_name(row.iloc[0]) if pd.notna(row.iloc[0]) else None
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
    """
    Lit les 2 tableaux (Hors Vacances scolaires / Vacances Scolaires).
    La ligne '17:00-19:00' est explosée en '17:00-18:00' et '18:00-19:00'.
    Retourne dict: {'Hors Vacances scolaires': {cren: {jour: nb}},
                    'Vacances Scolaires':       {cren: {jour: nb}}}
    """
    df = data['Besoins_Jeunesse']
    result = {}
    current_table = None
    headers = []

    for _, row in df.iterrows():
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''

        # Détection entête de tableau
        if 'hors vacances' in c0.lower():
            current_table = 'Hors Vacances scolaires'
            result[current_table] = {}
            headers = []
            continue
        if 'vacances scolaires' in c0.lower() and 'hors' not in c0.lower():
            current_table = 'Vacances Scolaires'
            result[current_table] = {}
            headers = []
            continue

        if current_table is None:
            continue

        # Ligne d'en-têtes colonnes
        if c1.lower() == 'créneau' or c1.lower() == 'creneau':
            headers = []
            for i in range(2, len(row)):
                h = str(row.iloc[i]).strip() if pd.notna(row.iloc[i]) else ''
                headers.append(h if h and h != 'nan' else '')
            continue

        # Ligne de données
        if not c1 or c1 == 'nan' or not headers:
            continue
        # Nettoyer le label créneau
        cren_raw = c1.strip()
        if '-' not in cren_raw:
            continue

        # Lire les valeurs
        entry = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            col_idx = i + 2
            if col_idx < len(row):
                v = row.iloc[col_idx]
                entry[h] = int(v) if pd.notna(v) and str(v) not in ('nan', '') else 0

        # Exploser 17:00-19:00 en deux créneaux
        def _add_cren(cren_label, entry_dict):
            result[current_table][cren_label] = entry_dict

        parts = cren_raw.split('-', 1)
        try:
            cs = hm_to_min(parts[0].strip())
            ce = hm_to_min(parts[1].strip())
        except Exception:
            continue
        if cs is None or ce is None:
            continue

        if ce - cs > 60:  # Créneau > 1h : exploser
            mid = cs + 60
            while mid < ce:
                cren1 = f"{cs//60:02d}:{cs%60:02d}-{mid//60:02d}:{mid%60:02d}"
                _add_cren(cren1, dict(entry))
                cs = mid
                mid = cs + 60
            cren_last = f"{cs//60:02d}:{cs%60:02d}-{ce//60:02d}:{ce%60:02d}"
            _add_cren(cren_last, dict(entry))
        else:
            norm = f"{cs//60:02d}:{cs%60:02d}-{ce//60:02d}:{ce%60:02d}"
            _add_cren(norm, dict(entry))

    return result

def parse_sp_minmax(data):
    """
    Lit SP_MinMax par semaine.
    Retourne dict: {week_num -> {agent -> {Min_MarVen, Max_MarVen, Min_MarSam, Max_MarSam}}}
    """
    df = data['SP_MinMax']
    result = {}
    current_week = None
    for _, row in df.iterrows():
        c0 = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ''
        c1 = normalize_agent_name(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
        if c0.startswith('SEMAINE'):
            parts = c0.split()
            if len(parts) >= 2:
                try:
                    current_week = int(parts[1])
                    if current_week not in result:
                        result[current_week] = {}
                except ValueError:
                    pass
            continue
        if not c1 or c1 in ('Agent', 'nan', '') or current_week is None:
            continue
        try:
            result[current_week][c1] = {
                'Min_MarVen':     float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
                'Max_MarVen':     float(row.iloc[3]) if pd.notna(row.iloc[3]) else 99,
                'Min_MarSam':     float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
                'Max_MarSam':     float(row.iloc[5]) if pd.notna(row.iloc[5]) else 99,
            }
        except (IndexError, ValueError):
            pass
    fallback = result.get(2) or (list(result.values())[0] if result else {})
    for i in range(1, 7):
        if i not in result:
            result[i] = dict(fallback)
    return result

def parse_roulement_samedi(data):
    """
    Nouveau format : blocs SEMAINE_1, SEMAINE_2...
    Retourne dict: {week_num -> {agent -> 'ROUGE'|'BLEU'}}
    Vacataires : peuvent avoir 2 valeurs (ROUGE + BLEU) → ils travaillent toujours.
    """
    df = data['Roulement_Samedi']
    result = {}
    current_week = None

    for _, row in df.iterrows():
        c0 = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ''
        c1 = normalize_agent_name(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        c2 = str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else ''

        if c0.startswith('SEMAINE'):
            parts = c0.split('_')
            if len(parts) >= 2:
                try:
                    current_week = int(parts[1])
                    if current_week not in result:
                        result[current_week] = {}
                except ValueError:
                    pass
            continue

        if not c1 or c1 in ('Agent', 'nan') or current_week is None:
            continue
        if c2 in ('ROUGE', 'BLEU'):
            # Vacataire avec 2 colonnes → travaille les deux types
            c3 = str(row.iloc[3]).strip().upper() if len(row) > 3 and pd.notna(row.iloc[3]) else ''
            if c3 in ('ROUGE', 'BLEU'):
                result[current_week][c1] = 'BOTH'
            else:
                result[current_week][c1] = c2

    # Fallback semaines manquantes
    fallback = result.get(1) or result.get(2) or {}
    for i in range(1, 7):
        if i not in result:
            result[i] = dict(fallback)
    return result

def parse_evenements(data, mois_debut, annee_debut, date_fin=None):
    df = data['Événements']
    result = defaultdict(list)
    for _, row in df.iterrows():
        d = parse_date_flexible(row.iloc[0], annee_debut)
        if d is None:
            continue
        in_main = (d.year == annee_debut and d.month == mois_debut)
        in_ext  = date_fin is not None and not in_main and d <= date_fin
        if not in_main and not in_ext:
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
    raw = params.get('Liste_des_créneaux', '')
    if not raw:
        return []
    result = []
    for c in str(raw).split(';'):
        c = c.strip()
        if '-' in c:
            parts = c.split('-', 1)
            s = hm_to_min(parts[0].strip())
            e = hm_to_min(parts[1].strip())
            if s is not None and e is not None:
                result.append((c, s, e))
    return result

def parse_samedi_types(params):
    result = {}
    for i in range(1, 7):
        val = params.get(f'Samedi_{i}') or params.get(f'Samedi_S{i}')
        if val and str(val).strip().upper() in ('ROUGE', 'BLEU'):
            result[i] = str(val).strip().upper()
    for i in range(1, 7):
        if i not in result:
            result[i] = 'ROUGE' if i % 2 == 1 else 'BLEU'
    return result


# ══════════════════════════════════════════════════════════════
#  PLANNING TYPE
# ══════════════════════════════════════════════════════════════

def parse_planning_type(data):
    df = data['planning_type']
    SECTION_COLS = {
        'RDC':      [2],
        'Adulte':   [3],
        'MF':       [5, 6],
        'Jeunesse': [7, 8, 9],
    }
    JOUR_MAP = {'MARDI':'Mardi','MERCREDI':'Mercredi','JEUDI':'Jeudi','VENDREDI':'Vendredi'}
    SKIP = {'nan','','R D C','Adulte','Musique & Films','Jeunesse','12H30','15H30','NaN'}

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
                        # Extraire juste le nom (ignorer "à partir de Xh")
                        agent_raw = v.split(' à ')[0].split(' à partir')[0].strip()
                        agent = normalize_agent_name(agent_raw)
                        if agent:
                            if section == 'Jeunesse':
                                if agent not in bloc[section]:
                                    bloc[section].append(agent)
                            else:
                                if not bloc[section]:
                                    bloc[section].append(agent)
        raw_blocs[current_jour].append((cell1, bloc))
    return raw_blocs

def explode_planning_type(raw_blocs, creneaux_def):
    def parse_bloc_time(label):
        label = label.strip().upper().replace('H', ':')
        if '-' not in label:
            return None, None
        parts = label.split('-', 1)
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
                    if cren_name not in result[jour_key]:
                        result[jour_key][cren_name] = {
                            s: list(agents) for s, agents in bloc_assignment.items()
                        }
    return result


# ══════════════════════════════════════════════════════════════
#  DISPONIBILITÉ
# ══════════════════════════════════════════════════════════════

def agent_covers_slot(agent, jour, cs, ce, horaires_agents):
    h = horaires_agents.get(agent, {}).get(jour)
    if h is None:
        return False
    sm, em, sa, ea = h
    in_m = sm is not None and em is not None and cs >= sm and ce <= em
    in_a = sa is not None and ea is not None and cs >= sa and ce <= ea
    return in_m or in_a

def agent_blocked_by_event(agent, cs, ce, date_str, evenements):
    for ev in evenements.get(date_str, []):
        if agent in ev['agents'] and overlap(cs, ce, ev['debut'], ev['fin']):
            return True
    return False

def agent_available(agent, jour, cs, ce, date_str, horaires_agents, evenements):
    if not agent_covers_slot(agent, jour, cs, ce, horaires_agents):
        return False
    if agent_blocked_by_event(agent, cs, ce, date_str, evenements):
        return False
    return True

def creneau_is_open(cs, ce, jour, horaires_ouverture):
    for os, oe in horaires_ouverture.get(jour, []):
        if cs >= os and ce <= oe:
            return True
    return False


# ══════════════════════════════════════════════════════════════
#  CONSÉCUTIF ET PAUSES
# ══════════════════════════════════════════════════════════════

def get_consecutive_sp_before(agent, cren_idx, creneaux, day_assignments):
    """SP consécutifs juste avant cren_idx."""
    cumul = 0
    for i in range(cren_idx - 1, -1, -1):
        cn, c_cs, c_ce = creneaux[i]
        slot = day_assignments.get(cn)
        if slot is None:
            break
        all_a = [a for s in SECTIONS for a in slot.get('assignment', {}).get(s, [])]
        if agent in all_a:
            cumul += (c_ce - c_cs)
        else:
            break
    return cumul

def get_sp_today_before(agent, cren_idx, creneaux, day_assignments):
    total = 0
    for i in range(cren_idx):
        cn, c_cs, c_ce = creneaux[i]
        slot = day_assignments.get(cn)
        if slot is None:
            continue
        all_a = [a for s in SECTIONS for a in slot.get('assignment', {}).get(s, [])]
        if agent in all_a:
            total += (c_ce - c_cs)
    return total

def agent_has_pause_before(agent, cren_idx, creneaux, day_assignments, pause_min=60):
    """True si l'agent a eu ≥ pause_min minutes d'affilée hors SP avant cren_idx."""
    cumul_pause = 0
    for i in range(cren_idx):
        cn, c_cs, c_ce = creneaux[i]
        slot = day_assignments.get(cn)
        if slot is None:
            cumul_pause += (c_ce - c_cs)
        else:
            all_a = [a for s in SECTIONS for a in slot.get('assignment', {}).get(s, [])]
            if agent not in all_a:
                cumul_pause += (c_ce - c_cs)
            else:
                cumul_pause = 0
        if cumul_pause >= pause_min:
            return True
    return False

def agent_meridienne_sp_total(agent, cren_idx, creneaux, day_assignments):
    """SP déjà effectué dans la fenêtre 12h-14h avant cren_idx."""
    MER_S, MER_E = 720, 840
    total = 0
    for i in range(cren_idx):
        cn, c_cs, c_ce = creneaux[i]
        if c_ce <= MER_S or c_cs >= MER_E:
            continue
        inter_s = max(c_cs, MER_S)
        inter_e = min(c_ce, MER_E)
        if inter_e <= inter_s:
            continue
        slot = day_assignments.get(cn)
        if slot is None:
            continue
        all_a = [a for s in SECTIONS for a in slot.get('assignment', {}).get(s, [])]
        if agent in all_a:
            total += (inter_e - inter_s)
    return total

def agent_has_meridienne_pause(agent, cren_idx, creneaux, day_assignments,
                                horaires_agents, jour):
    """
    D14 : réguliers uniquement.
    Vérifie qu'il existe au moins 1h continue sans SP dans la fenêtre 12h-14h.
    """
    if is_vacataire(agent):
        return True  # D14 ne s'applique pas aux vacataires

    MER_S, MER_E = 720, 840
    _, cs, ce = creneaux[cren_idx]

    if ce <= MER_S or cs >= MER_E:
        return True  # Créneau hors fenêtre méridienne

    inter_s = max(cs, MER_S)
    inter_e = min(ce, MER_E)
    slot_in_mer = inter_e - inter_s if inter_e > inter_s else 0

    sp_already = agent_meridienne_sp_total(agent, cren_idx, creneaux, day_assignments)

    # SP total dans la fenêtre après ce créneau
    total_sp = sp_already + slot_in_mer
    # La pause = 120 - total_sp doit être ≥ 60
    if total_sp > 60:
        return False
    return True

def count_congés_in_section(section, jour, date_str, affectations, evenements, horaires_agents):
    count = 0
    for agent, sects in affectations.items():
        if is_vacataire(agent):
            continue
        if section not in sects:
            continue
        h = horaires_agents.get(agent, {}).get(jour)
        if h is None:
            continue
        for ev in evenements.get(date_str, []):
            if agent in ev['agents'] and ev['nom'].lower() in ('congé', 'conge', 'vacation', 'rtt'):
                count += 1
                break
    return count

def violates_consec_hard(agent, jour, cs, ce, cren_idx, creneaux,
                          day_assignments, max_court, max_long,
                          section, date_str, affectations, evenements, horaires_agents):
    slot_min = ce - cs
    consec   = get_consecutive_sp_before(agent, cren_idx, creneaux, day_assignments)

    if jour in ('Mardi', 'Jeudi', 'Vendredi'):
        nb_conges = count_congés_in_section(section, jour, date_str, affectations,
                                             evenements, horaires_agents)
        limit = max_long if nb_conges > 0 else max_court
        return consec + slot_min > limit
    else:
        has_pause = agent_has_pause_before(agent, cren_idx, creneaux, day_assignments)
        return (consec + slot_min > max_long) and not has_pause

def over_ideal_consec(agent, cren_idx, creneaux, day_assignments, max_court):
    consec = get_consecutive_sp_before(agent, cren_idx, creneaux, day_assignments)
    return consec >= max_court


# ══════════════════════════════════════════════════════════════
#  VÉRIFICATION PAUSE VACATAIRE MÉRIDIENNE (O6)
# ══════════════════════════════════════════════════════════════

def vacataire_meridienne_ok(agent, cren_idx, creneaux, day_assignments, horaires_agents, jour):
    """
    O6 : pour les vacataires.
    Préférence : 1h pause entre 12h-14h.
    Sinon acceptable : pause 11h-12h OU 14h-15h.
    30min de pause minimum acceptable.
    Retourne (ok: bool, preference_score: int) — lower is better.
    """
    if not is_vacataire(agent):
        return True, 0

    MER_S, MER_E = 720, 840
    _, cs, ce = creneaux[cren_idx]

    if ce <= MER_S or cs >= MER_E:
        return True, 0  # hors fenêtre méridienne

    # Calculer SP total dans 12h-14h après ce créneau
    inter_s = max(cs, MER_S)
    inter_e = min(ce, MER_E)
    slot_in_mer = inter_e - inter_s if inter_e > inter_s else 0
    sp_already  = agent_meridienne_sp_total(agent, cren_idx, creneaux, day_assignments)
    total_sp    = sp_already + slot_in_mer
    pause_in_mer = 120 - total_sp

    if pause_in_mer >= 60:
        return True, 0   # Pause idéale ≥ 1h dans 12h-14h
    if pause_in_mer >= 30:
        # Acceptable si pause 11h-12h ou 14h-15h compensent
        has_pause_before = agent_has_pause_before(
            agent, cren_idx, creneaux, day_assignments, pause_min=30)
        return True, 1   # Acceptable mais sous-optimal
    # Pause < 30min dans méridienne : vérifier compensation externe
    return True, 2  # Dernier recours


# ══════════════════════════════════════════════════════════════
#  SÉLECTION DU REMPLAÇANT
# ══════════════════════════════════════════════════════════════

def find_replacement(section, jour, cs, ce, date_str,
                     eligible, affectations, horaires_agents, evenements,
                     sp_count, sp_max_min,
                     sp_week_count,
                     creneaux, cren_idx, day_assignments,
                     max_court, max_long,
                     exclude=None, vac_day_sp=None,
                     force_vacataire=False,
                     allow_any_section=False):  # D16
    """
    Cherche le meilleur agent pour une section/créneau.
    allow_any_section=True : D16, accepter tout agent disponible même non habilité.
    """
    exclude    = set(exclude or [])
    vac_day_sp = vac_day_sp or {}
    slot_min   = ce - cs
    VAC_MAX    = 420  # 7h/jour

    candidates = []
    for agent in eligible:
        if agent in exclude:
            continue

        # Vérifier habilitation (sauf D16)
        agent_sects = affectations.get(agent, [])
        if not allow_any_section and section not in agent_sects:
            continue

        if not agent_available(agent, jour, cs, ce, date_str, horaires_agents, evenements):
            continue

        # D6 : max SP hebdo
        if sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
            continue

        # Vacataire max journalier
        if is_vacataire(agent) and vac_day_sp.get(agent, 0) + slot_min > VAC_MAX:
            continue

        # D7/D8 : limite absolue consécutif
        if violates_consec_hard(agent, jour, cs, ce, cren_idx, creneaux,
                                 day_assignments, max_court, max_long,
                                 section, date_str, affectations, evenements, horaires_agents):
            continue

        # D14 : pause méridienne (réguliers uniquement)
        if not agent_has_meridienne_pause(agent, cren_idx, creneaux, day_assignments,
                                           horaires_agents, jour):
            continue

        # Critères de tri
        out_of_section = 0 if (allow_any_section is False or section in agent_sects) else 1

        if is_vacataire(agent):
            # O2 exception vacataires : pas de priorité section, priorité = combler les trous
            prim = 0
        else:
            prim = 0 if (agent_sects and agent_sects[0] == section) else 1  # O2

        vac = 1 if is_vacataire(agent) else 0

        if force_vacataire:
            vac_score = 0 if is_vacataire(agent) else 1
        else:
            vac_score = vac

        # Priorité vacataire : Vacataire 1 avant Vacataire 2
        vac_order = 0
        if is_vacataire(agent):
            vac_order = 0 if '1' in agent else 1

        over_ideal  = 1 if over_ideal_consec(agent, cren_idx, creneaux,
                                              day_assignments, max_court) else 0
        sp_semaine  = sp_week_count.get(agent, 0)
        sp_jour     = get_sp_today_before(agent, cren_idx, creneaux, day_assignments)
        court_cren  = 1 if slot_min <= 60 else 0

        # Score pause méridienne vacataire
        _, mer_score = vacataire_meridienne_ok(agent, cren_idx, creneaux,
                                                day_assignments, horaires_agents, jour)

        candidates.append((
            out_of_section,  # D16 : section habituelle avant hors-section
            prim,            # O2
            vac_score,       # O6/défaut
            vac_order,       # Vac1 avant Vac2
            over_ideal,      # O7
            mer_score,       # O6 pause méridienne
            sp_semaine,      # O3
            sp_jour,         # O4
            court_cren,      # O5
            agent
        ))

    if not candidates:
        return None, False
    candidates.sort()
    best = candidates[0]
    out_of_section_flag = best[0] == 1
    return best[-1], out_of_section_flag


# ══════════════════════════════════════════════════════════════
#  PLANIFICATION D'UNE SEMAINE
# ══════════════════════════════════════════════════════════════

def get_samedi_eligible(roulement_semaine, samedi_type, affectations, horaires_agents):
    """Agents éligibles pour un samedi donné."""
    agents = []
    for agent, roul in roulement_semaine.items():
        if roul == 'BOTH' or roul == samedi_type:
            if horaires_agents.get(agent, {}).get('Samedi') is not None:
                agents.append(agent)
    # Ajouter vacataires si absents
    for agent in affectations:
        if is_vacataire(agent) and agent not in agents:
            if horaires_agents.get(agent, {}).get('Samedi') is not None:
                agents.append(agent)
    return agents

def get_besoins_jeunesse_slot(besoins_j, cren_name, jour, samedi_type, semaine_type):
    """Retourne le besoin Jeunesse pour un créneau/jour/type de semaine."""
    table = besoins_j.get(semaine_type, besoins_j.get('Hors Vacances scolaires', {}))
    entry = table.get(cren_name, {})
    if not entry:
        return 0

    if jour == 'Samedi':
        key_rouge = next((k for k in entry if 'rouge' in k.lower()), None)
        key_bleu  = next((k for k in entry if 'bleu' in k.lower()), None)
        if samedi_type == 'ROUGE' and key_rouge:
            return int(entry.get(key_rouge, 0))
        elif samedi_type == 'BLEU' and key_bleu:
            return int(entry.get(key_bleu, 0))
        return 0

    # Chercher la colonne correspondant au jour
    for k, v in entry.items():
        if k.lower().startswith(jour.lower()[:4]):
            return int(v) if v else 0
    return 0

def compute_vac_min_sp(agent, jour, horaires_agents):
    """
    O6/D15 : calcule le min SP attendu pour un vacataire.
    Si matin ET après-midi → 6h. Si demi-journée → proportionnel.
    """
    h = horaires_agents.get(agent, {}).get(jour)
    if h is None:
        return 0
    sm, em, sa, ea = h
    has_matin = sm is not None and em is not None
    has_apm   = sa is not None and ea is not None
    if has_matin and has_apm:
        return 360  # 6h
    elif has_matin:
        dispo = em - sm
        return max(0, int(dispo * 0.8))
    elif has_apm:
        dispo = ea - sa
        return max(0, int(dispo * 0.8))
    return 0


def plan_week(week_num, week_dates, planning_type_base, samedi_type,
              affectations, horaires_agents, horaires_ouverture,
              besoins_jeunesse, sp_minmax_week, roulement_semaine,
              evenements, params, creneaux, semaine_type):

    max_court = parse_duration_param(params.get('Durée_SP_max_idéale', '2h30'), 150)
    max_long  = parse_duration_param(params.get('Durée_SP_max_tolérée', '4h'), 240)
    vac_days  = str(params.get('Mode_vacataires', 'mercredi,samedi')).lower()
    exc_vac_s = hm_to_min(str(params.get('Exception_Vacataire_seul', '12:00-14:00')).split('-')[0]) or 720
    exc_vac_e = hm_to_min(str(params.get('Exception_Vacataire_seul', '12:00-14:00')).split('-')[-1]) or 840
    all_agents = list(affectations.keys())
    is_sam_week = 'Samedi' in week_dates

    # D6 : SP min/max en minutes
    sp_max_min = {}
    sp_min_min = {}
    for agent, mm in sp_minmax_week.items():
        if is_sam_week:
            sp_max_min[agent] = mm['Max_MarSam'] * 60
            sp_min_min[agent] = mm['Min_MarSam'] * 60
        else:
            sp_max_min[agent] = mm['Max_MarVen'] * 60
            sp_min_min[agent] = mm['Min_MarVen'] * 60
    for agent in all_agents:
        if is_vacataire(agent) and agent not in sp_max_min:
            sp_max_min[agent] = 420 * 2
            sp_min_min[agent] = 0

    result    = {}
    sp_count  = defaultdict(int)
    sp_jour_cum = defaultdict(lambda: defaultdict(int))

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        date_str   = date.strftime('%Y-%m-%d')
        is_vac_day = jour.lower() in vac_days

        # Agents éligibles ce jour
        if jour == 'Samedi':
            eligible = get_samedi_eligible(roulement_semaine, samedi_type,
                                            affectations, horaires_agents)
        elif is_vac_day:
            eligible = [a for a in all_agents
                        if horaires_agents.get(a, {}).get(jour) is not None]
        else:
            eligible = [a for a in all_agents if not is_vacataire(a)
                        and horaires_agents.get(a, {}).get(jour) is not None]

        pt_key            = f"Samedi_{samedi_type}" if jour == 'Samedi' else jour
        base              = planning_type_base.get(pt_key, {})
        max_rot           = JOURS_COURT.get(jour, 4)
        agents_used_today = {s: [] for s in SECTIONS}
        day_assignments   = {}
        vac_day_sp        = defaultdict(int)

        result[jour] = {'_samedi_type': samedi_type if jour == 'Samedi' else None}

        for cren_idx, (cren_name, cs, ce) in enumerate(creneaux):
            slot_min = ce - cs

            if not creneau_is_open(cs, ce, jour, horaires_ouverture):
                result[jour][cren_name] = None
                continue

            base_slot  = base.get(cren_name, {s: [] for s in SECTIONS})
            assignment = {s: list(base_slot.get(s, [])) for s in SECTIONS}
            slot_alerts      = []
            slot_out_section = {}  # section -> True si agent hors habilitation (D16)
            assigned_this    = set()

            # Besoin Jeunesse pour ce créneau
            j_needs = get_besoins_jeunesse_slot(
                besoins_jeunesse, cren_name, jour, samedi_type, semaine_type)

            # ══ PASSE A : valider/remplacer agents du planning type ══
            for section in SECTIONS:
                final = []
                pt_agents_orig = list(base_slot.get(section, []))

                # D4 : tronquer Jeunesse au besoin AVANT traitement
                if section == 'Jeunesse':
                    pt_agents_orig = pt_agents_orig[:j_needs]

                for agent in pt_agents_orig:
                    # Résoudre générique "VACATAIRE(S)"
                    if agent.upper() in ('VACATAIRES', 'VACATAIRE', 'VACATAIRE 1',
                                         'VACATAIRE 2') and section != 'Jeunesse':
                        agent_resolved = agent
                    elif agent.upper().startswith('VACATAIRE'):
                        agent_resolved = agent
                    else:
                        agent_resolved = agent

                    if agent_resolved in assigned_this:
                        continue

                    must_replace = False
                    force_vac    = False

                    if not agent_available(agent_resolved, jour, cs, ce, date_str,
                                           horaires_agents, evenements):
                        must_replace = True
                    elif sp_count.get(agent_resolved, 0) + slot_min > sp_max_min.get(agent_resolved, 99*60):
                        must_replace = True
                    elif violates_consec_hard(agent_resolved, jour, cs, ce, cren_idx, creneaux,
                                              day_assignments, max_court, max_long,
                                              section, date_str, affectations,
                                              evenements, horaires_agents):
                        must_replace = True
                    elif not agent_has_meridienne_pause(agent_resolved, cren_idx, creneaux,
                                                         day_assignments, horaires_agents, jour):
                        must_replace = True
                    elif is_vac_day and not is_vacataire(agent_resolved):
                        consec = get_consecutive_sp_before(
                            agent_resolved, cren_idx, creneaux, day_assignments)
                        if consec >= max_court:
                            force_vac    = True
                            must_replace = True

                    if not must_replace:
                        final.append(agent_resolved)
                        assigned_this.add(agent_resolved)
                    else:
                        repl, oos = find_replacement(
                            section=section, jour=jour, cs=cs, ce=ce,
                            date_str=date_str, eligible=eligible,
                            affectations=affectations,
                            horaires_agents=horaires_agents,
                            evenements=evenements,
                            sp_count=sp_count, sp_max_min=sp_max_min,
                            sp_week_count=dict(sp_count),
                            creneaux=creneaux, cren_idx=cren_idx,
                            day_assignments=day_assignments,
                            max_court=max_court, max_long=max_long,
                            exclude=assigned_this,
                            vac_day_sp=vac_day_sp,
                            force_vacataire=force_vac,
                        )
                        if repl:
                            final.append(repl)
                            assigned_this.add(repl)

                assignment[section] = final

            # ══ PASSE B : sections RDC/Adulte/MF obligatoires (D5) ══
            # Traiter en ordre "most constrained first" : section avec le moins de candidats en premier
            def count_candidates(sect):
                count = 0
                for ag in eligible:
                    if ag in assigned_this: continue
                    if sect not in affectations.get(ag, []): continue
                    if not agent_available(ag, jour, cs, ce, date_str, horaires_agents, evenements): continue
                    if sp_count.get(ag, 0) + slot_min > sp_max_min.get(ag, 99*60): continue
                    count += 1
                return count

            empty_sections = [s for s in ['RDC', 'Adulte', 'MF'] if not assignment[s]]
            empty_sections.sort(key=count_candidates)

            for section in empty_sections:
                repl, oos = find_replacement(
                    section=section, jour=jour, cs=cs, ce=ce,
                    date_str=date_str, eligible=eligible,
                    affectations=affectations,
                    horaires_agents=horaires_agents,
                    evenements=evenements,
                    sp_count=sp_count, sp_max_min=sp_max_min,
                    sp_week_count=dict(sp_count),
                    creneaux=creneaux, cren_idx=cren_idx,
                    day_assignments=day_assignments,
                    max_court=max_court, max_long=max_long,
                    exclude=assigned_this,
                    vac_day_sp=vac_day_sp,
                )
                if repl:
                    assignment[section] = [repl]
                    assigned_this.add(repl)
                    if oos:
                        slot_out_section[section] = True

            # ══ PASSE C : Besoins Jeunesse exacts (D4) ══
            # Dédoublonner et tronquer au besoin max
            seen_j = []
            for a in assignment['Jeunesse']:
                if a not in seen_j and len(seen_j) < j_needs:
                    seen_j.append(a)
            assignment['Jeunesse'] = seen_j

            # Compléter si insuffisant
            if len(assignment['Jeunesse']) < j_needs:
                for agent in eligible:
                    if len(assignment['Jeunesse']) >= j_needs:
                        break
                    if agent in assigned_this:
                        continue
                    if 'Jeunesse' not in affectations.get(agent, []):
                        continue
                    if not agent_available(agent, jour, cs, ce, date_str,
                                            horaires_agents, evenements):
                        continue
                    if sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
                        continue
                    if is_vacataire(agent) and vac_day_sp.get(agent, 0) + slot_min > 420:
                        continue
                    if violates_consec_hard(agent, jour, cs, ce, cren_idx, creneaux,
                                            day_assignments, max_court, max_long,
                                            'Jeunesse', date_str, affectations,
                                            evenements, horaires_agents):
                        continue
                    assignment['Jeunesse'].append(agent)
                    assigned_this.add(agent)

            # D16 si encore insuffisant (dernier recours)
            if len(assignment['Jeunesse']) < j_needs:
                needed = j_needs - len(assignment['Jeunesse'])
                for agent in eligible:
                    if needed <= 0:
                        break
                    if agent in assigned_this:
                        continue
                    if not agent_available(agent, jour, cs, ce, date_str,
                                            horaires_agents, evenements):
                        continue
                    if sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
                        continue
                    assignment['Jeunesse'].append(agent)
                    assigned_this.add(agent)
                    slot_out_section['Jeunesse'] = True
                    needed -= 1

            if len(assignment['Jeunesse']) < j_needs:
                slot_alerts.append(
                    f"ALERTE — Jeunesse : {len(assignment['Jeunesse'])}/{j_needs} agents")

            # ══ PASSE D : pas de vacataire seul (D11) — vérification finale ══
            for section in SECTIONS:
                ag_list = assignment.get(section, [])
                if not ag_list:
                    continue
                if all(is_vacataire(a) for a in ag_list):
                    if not (cs >= exc_vac_s and ce <= exc_vac_e):
                        # Chercher un régulier parmi TOUS les éligibles (pas seulement non-assignés)
                        # Si trouvé parmi les déjà-assignés à une autre section → SWAP
                        found_regular = None

                        # D'abord chercher parmi les non-assignés
                        repl, _ = find_replacement(
                            section=section, jour=jour, cs=cs, ce=ce,
                            date_str=date_str, eligible=eligible,
                            affectations=affectations,
                            horaires_agents=horaires_agents,
                            evenements=evenements,
                            sp_count=sp_count, sp_max_min=sp_max_min,
                            sp_week_count=dict(sp_count),
                            creneaux=creneaux, cren_idx=cren_idx,
                            day_assignments=day_assignments,
                            max_court=max_court, max_long=max_long,
                            exclude=assigned_this,
                            vac_day_sp=vac_day_sp,
                        )
                        if repl and not is_vacataire(repl):
                            found_regular = repl
                        else:
                            # Chercher parmi les déjà assignés à d'autres sections
                            for other_sect in SECTIONS:
                                if other_sect == section:
                                    continue
                                for ag in list(assignment.get(other_sect, [])):
                                    if is_vacataire(ag):
                                        continue
                                    if section not in affectations.get(ag, []):
                                        continue
                                    if not agent_available(ag, jour, cs, ce, date_str,
                                                            horaires_agents, evenements):
                                        continue
                                    # SWAP : mettre le vacataire dans l'autre section
                                    # et l'agent régulier ici
                                    vac_agent = ag_list[0]
                                    if other_sect in affectations.get(vac_agent, []):
                                        assignment[other_sect].remove(ag)
                                        assignment[other_sect].append(vac_agent)
                                        # Retirer le vacataire de la section courante
                                        for v in list(ag_list):
                                            if is_vacataire(v):
                                                assignment[section].remove(v)
                                        found_regular = ag
                                        break
                                if found_regular:
                                    break

                        if found_regular:
                            if section == 'Jeunesse':
                                if len(assignment['Jeunesse']) >= j_needs:
                                    assignment['Jeunesse'].pop()
                                assignment['Jeunesse'].insert(0, found_regular)
                            else:
                                assignment[section] = [found_regular]
                            assigned_this.add(found_regular)
                        else:
                            slot_alerts.append(
                                f"ALERTE — {section} : vacataire seul hors 12h-14h")

            # ══ PASSE E : D16 final pour RDC/Adulte/MF encore vides ══
            for section in ['RDC', 'Adulte', 'MF']:
                if not assignment[section]:
                    repl, oos = find_replacement(
                        section=section, jour=jour, cs=cs, ce=ce,
                        date_str=date_str, eligible=eligible,
                        affectations=affectations,
                        horaires_agents=horaires_agents,
                        evenements=evenements,
                        sp_count=sp_count, sp_max_min=sp_max_min,
                        sp_week_count=dict(sp_count),
                        creneaux=creneaux, cren_idx=cren_idx,
                        day_assignments=day_assignments,
                        max_court=max_court, max_long=max_long,
                        exclude=assigned_this,
                        vac_day_sp=vac_day_sp,
                        allow_any_section=True,  # D16
                    )
                    if repl:
                        assignment[section] = [repl]
                        assigned_this.add(repl)
                        slot_out_section[section] = True
                    else:
                        slot_alerts.append(f"ALERTE — {section} : aucun agent disponible")

            # ══ Nettoyage alertes obsolètes Jeunesse ══
            # Recalculer après toutes les passes (Passe C peut avoir alerté prématurément)
            j_ag_final = assignment.get('Jeunesse', [])
            slot_alerts = [al for al in slot_alerts if 'Jeunesse' not in al]
            if len(j_ag_final) < j_needs:
                slot_alerts.append(
                    f"ALERTE — Jeunesse : {len(j_ag_final)}/{j_needs} agents")

            # ══ Compteurs ══
            for section in SECTIONS:
                for agent in assignment[section]:
                    sp_count[agent]       += slot_min
                    sp_jour_cum[jour][agent] += slot_min
                    if is_vacataire(agent):
                        vac_day_sp[agent] += slot_min
                    if agent not in agents_used_today[section]:
                        agents_used_today[section].append(agent)

            slot_events = [ev for ev in evenements.get(date_str, [])
                           if overlap(cs, ce, ev['debut'], ev['fin'])]

            day_assignments[cren_name] = {
                'assignment':    assignment,
                'events':        slot_events,
                'alerts':        slot_alerts,
                'out_of_section': slot_out_section,
                'open':          True,
            }
            result[jour][cren_name] = day_assignments[cren_name]

    # ══════════════════════════════════════════════════════════
    # O8 : RECALAGE FINAL — Min SP
    # ══════════════════════════════════════════════════════════
    sp_alerts = {}
    for agent in list(sp_minmax_week.keys()):
        if is_vacataire(agent):
            continue
        mm     = sp_minmax_week.get(agent, {})
        min_sp = (mm['Min_MarSam'] if is_sam_week else mm['Min_MarVen']) * 60
        max_sp = (mm['Max_MarSam'] if is_sam_week else mm['Max_MarVen']) * 60
        actual = sp_count.get(agent, 0)
        if actual >= min_sp:
            continue

        for jour in JOURS_SP:
            if actual >= min_sp:
                break
            date = week_dates.get(jour)
            if date is None:
                continue
            date_str = date.strftime('%Y-%m-%d')
            for cren_idx, (cren_name, cs, ce) in enumerate(creneaux):
                if actual >= min_sp:
                    break
                slot = result[jour].get(cren_name)
                if not slot or not isinstance(slot, dict) or not slot.get('open'):
                    continue
                assignment = slot.get('assignment', {})
                already = any(agent in assignment.get(s, []) for s in SECTIONS)
                if already:
                    continue
                if not agent_available(agent, jour, cs, ce, date_str,
                                        horaires_agents, evenements):
                    continue
                if sp_count.get(agent, 0) + (ce - cs) > max_sp:
                    continue
                for section in ['Jeunesse', 'RDC', 'Adulte', 'MF']:
                    if section not in affectations.get(agent, []):
                        continue
                    # D4 : ne pas dépasser le besoin Jeunesse
                    if section == 'Jeunesse':
                        j_needs_slot = get_besoins_jeunesse_slot(
                            besoins_jeunesse, cren_name, jour,
                            result[jour].get('_samedi_type', 'ROUGE'),
                            semaine_type)
                        if len(assignment['Jeunesse']) >= j_needs_slot:
                            continue
                    else:
                        # D3 : max 1 agent
                        if len(assignment.get(section, [])) >= 1:
                            continue
                    added = ce - cs
                    assignment[section].append(agent)
                    sp_count[agent] += added
                    actual += added
                    break

        actual = sp_count.get(agent, 0)
        if actual < min_sp:
            sp_alerts[agent] = (
                f"SP insuffisant : {int(actual)//60}h{int(actual)%60:02d} "
                f"/ min {int(min_sp)//60}h{int(min_sp)%60:02d}"
            )

    # ══ Nettoyage post-O8 : alertes Jeunesse devenues obsolètes ══
    for jour in JOURS_SP:
        if jour not in result: continue
        jp = result[jour]
        for cn in list(jp.keys()):
            slot = jp.get(cn)
            if not slot or not isinstance(slot, dict) or not slot.get('open'): continue
            j_ag = slot.get('assignment', {}).get('Jeunesse', [])
            sam_t = jp.get('_samedi_type', 'ROUGE')
            j_needs_post = get_besoins_jeunesse_slot(
                besoins_jeunesse, cn, jour, sam_t, semaine_type)
            old_alerts = slot.get('alerts', [])
            new_alerts = [al for al in old_alerts if 'Jeunesse' not in al]
            if len(j_ag) < j_needs_post:
                new_alerts.append(
                    f"ALERTE — Jeunesse : {len(j_ag)}/{j_needs_post} agents")
            slot['alerts'] = new_alerts

    # Vérification min SP vacataires (O6/D15)
    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        is_vac_day = jour.lower() in str(params.get('Mode_vacataires', '')).lower()
        if not is_vac_day:
            continue
        for agent in all_agents:
            if not is_vacataire(agent):
                continue
            vac_min = compute_vac_min_sp(agent, jour, horaires_agents)
            actual  = sp_jour_cum[jour].get(agent, 0)
            if vac_min > 0 and actual < vac_min and actual > 0:
                sp_alerts[f"{agent}_{jour}"] = (
                    f"{agent} {jour} SP insuffisant : "
                    f"{actual//60}h{actual%60:02d} / min {vac_min//60}h{vac_min%60:02d}"
                )

    return result, samedi_type, dict(sp_count), sp_alerts, dict(sp_jour_cum)


# ══════════════════════════════════════════════════════════════
#  CALENDRIER
# ══════════════════════════════════════════════════════════════

def get_weeks_of_month(mois, annee):
    first_day       = datetime(annee, mois, 1)
    days_to_tuesday = (1 - first_day.weekday()) % 7
    first_tuesday   = first_day + timedelta(days=days_to_tuesday)
    offsets = {'Mardi':0,'Mercredi':1,'Jeudi':2,'Vendredi':3,'Samedi':4}

    all_weeks = []
    current   = first_tuesday
    while True:
        week_full = {j: current + timedelta(days=off) for j, off in offsets.items()}
        has_mois  = any(d.month == mois and d.year == annee for d in week_full.values())
        if not has_mois:
            break
        all_weeks.append(week_full)
        current += timedelta(weeks=1)

    result = []
    for i, week in enumerate(all_weeks):
        if i < len(all_weeks) - 1:
            result.append({j: d for j, d in week.items()
                           if d.month == mois and d.year == annee})
        else:
            result.append(week)
    return result


# ══════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════

def compute_full_planning(filepath):
    raw              = load_excel_data(filepath)
    params           = parse_parametres(raw)
    horaires_ouv     = parse_horaires_ouverture(raw)
    affectations     = parse_affectations(raw)
    horaires_agents  = parse_horaires_agents(raw)
    besoins_j        = parse_besoins_jeunesse(raw)
    sp_minmax_all    = parse_sp_minmax(raw)
    roulement_all    = parse_roulement_samedi(raw)
    creneaux         = parse_creneaux(params)
    samedi_types     = parse_samedi_types(params)
    semaines_type    = parse_semaines_type(params)

    mois_str = str(params.get('Mois', 'janvier')).strip().lower()
    annee    = int(params.get('Année', 2026))
    mois_num = MOIS_MAP.get(mois_str, 1)

    weeks    = get_weeks_of_month(mois_num, annee)
    date_fin = weeks[-1].get('Samedi') if weeks else None
    evenements = parse_evenements(raw, mois_num, annee, date_fin=date_fin)

    raw_blocs          = parse_planning_type(raw)
    planning_type_base = explode_planning_type(raw_blocs, creneaux)

    results = []
    for i, week_dates in enumerate(weeks):
        week_num       = i + 1
        samedi_type    = samedi_types.get(week_num, 'ROUGE' if week_num % 2 == 1 else 'BLEU')
        sp_minmax_week = sp_minmax_all.get(week_num, sp_minmax_all.get(2, {}))
        roulement_sem  = roulement_all.get(week_num, roulement_all.get(1, {}))
        semaine_type   = semaines_type.get(week_num, 'Hors Vacances scolaires')

        week_plan, sam_type, sp_count, sp_alerts, sp_jour_cum = plan_week(
            week_num           = week_num,
            week_dates         = week_dates,
            planning_type_base = planning_type_base,
            samedi_type        = samedi_type,
            affectations       = affectations,
            horaires_agents    = horaires_agents,
            horaires_ouverture = horaires_ouv,
            besoins_jeunesse   = besoins_j,
            sp_minmax_week     = sp_minmax_week,
            roulement_semaine  = roulement_sem,
            evenements         = evenements,
            params             = params,
            creneaux           = creneaux,
            semaine_type       = semaine_type,
        )
        results.append({
            'week_num':    week_num,
            'week_dates':  week_dates,
            'plan':        week_plan,
            'samedi_type': sam_type,
            'sp_count':    sp_count,
            'sp_alerts':   sp_alerts,
            'sp_jour_cum': sp_jour_cum,
            'semaine_type': semaine_type,
        })

    metadata = {
        'mois':            mois_str,
        'annee':           annee,
        'params':          params,
        'affectations':    affectations,
        'creneaux':        creneaux,
        'horaires_ouv':    horaires_ouv,
        'horaires_agents': horaires_agents,
        'roulement_all':   roulement_all,
        'sp_minmax_all':   sp_minmax_all,
        'evenements':      evenements,
        'besoins_j':       besoins_j,
        'semaines_type':   semaines_type,
    }
    return results, metadata
