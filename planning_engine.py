"""
Planning Engine v11 — Médiathèque
════════════════════════════════════════════════════════════════
CONTRAINTES DURES (D1-D17) :
  D1.  Présence réelle (Horaires_Des_Agents)
  D2.  Événements / indisponibilités
  D3.  Sections autorisées + max 1 agent/section/créneau (sauf Jeunesse)
       Les vacataires n'ont pas RDC dans leurs affectations → exclus de RDC
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
  D15. SP vacataire : journée complète → exactement 8h SP (min ET max)
       Demi-journée matin (10h-13h) → 3h SP continu, sans pause
       Demi-journée après-midi (14h-19h) → 5h SP continu, sans pause
  D16. Agent hors section habituelle en dernier recours absolu (cellule rouge)
  D17. Stéphane strictement verrouillé sur sa section — jamais affecté via D16

CONTRAINTES OPTIMISÉES (O1-O9) :
  O1.  Planning type = référence (minimum de changements, réf D1-D17)
  O2.  Section primaire prioritaire; exception vacataires = comblent les trous
  O3.  Répartition équitable SP par jour ET par semaine
  O4.  Répartir heures SP dans la journée
  O5.  Éviter créneaux ≤1h sauf méridienne
  O6.  Règles vacataires :
       - Samedi : toujours ≥1 vacataire obligatoire
       - Mercredi : vacataire uniquement si planning non réalisable sans eux
       - Vac1 prioritaire sur Vac2
       - Dès qu'un vacataire est présent : prioritaire sur tous réguliers
       - Relais dès 2h30 consécutif régulier
       - Pause méridienne préférence 1h entre 12h-14h (blocs 30min)
       - Min/max 8h journée complète, 3h matin, 5h après-midi
  O7.  Éviter >2h30 consécutif (déprioritiser)
  O8.  Recalage final min SP réguliers — uniquement si aucun vacataire dispo
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

# E : Blocs horaires de continuité par type de jour
# Jours avec coupure (Mardi, Jeudi, Vendredi) : 3 blocs
# Jours continus (Mercredi, Samedi) : 6 blocs
# F : Exception continuité 10h-14h pour Lydie et Delphine
BLOCS_COUPURE = [
    (600, 750),   # 10h-12h30
    (930, 1020),  # 15h30-17h
    (1020, 1140), # 17h-19h
]
BLOCS_CONTINU = [
    (600, 720),   # 10h-12h
    (720, 780),   # 12h-13h
    (780, 840),   # 13h-14h
    (840, 900),   # 14h-15h
    (900, 1020),  # 15h-17h
    (1020, 1140), # 17h-19h
]
JOURS_CONTINU = {'Mercredi', 'Samedi'}
# F : agents pouvant travailler en continu 10h-14h sans pause méridienne
AGENTS_EXCEPTION_MERIDIENNE = {'lydie', 'delphine'}

def get_bloc_id(jour, cs):
    """Retourne l'identifiant du bloc horaire (int) pour ce jour/heure début."""
    blocs = BLOCS_CONTINU if jour in JOURS_CONTINU else BLOCS_COUPURE
    for i, (bs, be) in enumerate(blocs):
        if bs <= cs < be:
            return i
    return -1  # hors bloc (pause ou fermé)


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
    """
    Retourne (affectations, categories) :
      affectations : {agent: [section1, section2, ...]}  (ordre = priorité)
      categories   : {agent: str|None}  (A, B, C, C2, ou None pour vacataires)

    Format Excel attendu : Agent | Catégorie | Section1 | Section2 | ...
    Si la colonne Catégorie est absente (ancien format à 1 col d'agent + sections),
    on retourne categories={} et on lit depuis col 1.
    Les colonnes de notes (contenant ":", "→", ou >20 chars) sont ignorées.
    """
    df = data['Affectations']
    result = {}
    categories = {}

    # Détecter la présence de la colonne Catégorie en lisant la 1ère ligne de données
    has_cat = False
    for _, row in df.iterrows():
        agent_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        if agent_val in ('', 'nan', 'Agent'):
            continue
        # Si la 2ème valeur est une lettre unique ou C2 → c'est une catégorie
        if len(row) > 1:
            v2 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            if v2 in ('A', 'B', 'C', 'C2', ''):
                has_cat = True
        break

    sect_start = 2 if has_cat else 1

    for _, row in df.iterrows():
        agent = normalize_agent_name(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        if not agent or agent in ('Agent', 'nan'):
            continue
        # Catégorie
        if has_cat and len(row) > 1:
            v = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            categories[agent] = v if v not in ('nan', '') else None
        else:
            categories[agent] = None
        # Sections (ignorer colonnes de notes)
        sections = []
        for i in range(sect_start, len(row)):
            v = str(row.iloc[i]).strip() if pd.notna(row.iloc[i]) else ''
            if v and v not in ('nan',) and len(v) <= 20 and ':' not in v and '→' not in v:
                sections.append(v)
        if sections:
            result[agent] = sections
    return result, categories


def get_sections_sans_alerte(agent, affectations, categories):
    """
    Sections où l'agent peut aller sans alerte rouge, selon sa catégorie :
      A  → section primaire uniquement (Jeunesse)
      B  → section primaire uniquement (MF)
      C  → Adulte + RDC (équivalents)
      C2 → section primaire uniquement (Adulte)
      None (vacataires) → toutes leurs sections sauf RDC
    """
    cat = categories.get(agent)
    sects = affectations.get(agent, [])
    if cat == 'A':
        return sects[:1]
    elif cat == 'B':
        return sects[:1]
    elif cat == 'C':
        return [s for s in sects if s in ('Adulte', 'RDC')]
    elif cat == 'C2':
        return [sects[0]] if sects else []
    else:
        # Vacataires : toutes sections sauf RDC
        return [s for s in sects if s != 'RDC']


def is_section_rouge(agent, section, affectations, categories):
    """True si affecter cet agent à cette section déclenche une alerte rouge (D16)."""
    return section not in get_sections_sans_alerte(agent, affectations, categories)

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
        SKIP_C1 = ('Agent', 'nan', '', 'Min_MarVen', 'Max_MarVen',
                   'Min_MarSam', 'Max_MarSam', 'SP_Samedi_Type', 'Créneau')
        if not c1 or c1 in SKIP_C1 or current_week is None:
            continue
        # Vérifier que les valeurs numériques sont bien des nombres
        try:
            v2 = row.iloc[2] if len(row) > 2 else None
            v3 = row.iloc[3] if len(row) > 3 else None
            v4 = row.iloc[4] if len(row) > 4 else None
            v5 = row.iloc[5] if len(row) > 5 else None
            # Skip si les valeurs sont des chaînes (ligne d'en-tête)
            if any(isinstance(v, str) and not v.replace('.','').replace('-','').isdigit()
                   for v in [v2, v3, v4, v5] if v is not None and pd.notna(v)):
                continue
            result[current_week][c1] = {
                'Min_MarVen': float(v2) if v2 is not None and pd.notna(v2) else 0,
                'Max_MarVen': float(v3) if v3 is not None and pd.notna(v3) else 99,
                'Min_MarSam': float(v4) if v4 is not None and pd.notna(v4) else 0,
                'Max_MarSam': float(v5) if v5 is not None and pd.notna(v5) else 99,
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
    D14 (réguliers) + O6 (vacataires) :
    Au moins 1h sans SP dans la fenêtre 12h-14h.
    Pour les vacataires : découpable en créneaux de 30min (12h30-13h30 possible).
    F : Lydie et Delphine sont exemptées — peuvent travailler en continu 10h-14h.
    """
    # F : exemption Lydie et Delphine
    if agent.lower() in AGENTS_EXCEPTION_MERIDIENNE:
        return True

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
    # F : Lydie et Delphine peuvent travailler 10h-14h sans limite (4h exactement)
    if agent.lower() in AGENTS_EXCEPTION_MERIDIENNE and cs < 840:
        return False
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
                     eligible, affectations, categories, horaires_agents, evenements,
                     sp_count, sp_max_min,
                     sp_week_count,
                     creneaux, cren_idx, day_assignments,
                     max_court, max_long,
                     exclude=None, vac_day_sp=None,
                     force_vacataire=False,
                     allow_any_section=False,   # D16
                     vac_prioritaire=False):    # O6 : vacataire prioritaire sur réguliers
    """
    Cherche le meilleur agent pour une section/créneau.
    allow_any_section=True : D16, accepter tout agent disponible même non habilité.
    vac_prioritaire=True : O6, vacataire avant tous les réguliers.
    """
    exclude    = set(exclude or [])
    vac_day_sp = vac_day_sp or {}
    slot_min   = ce - cs
    VAC_MAX    = 480  # D15 : 8h/jour max

    candidates = []
    for agent in eligible:
        if agent in exclude:
            continue

        # D17 : Stéphane jamais affecté hors section via D16
        if allow_any_section and agent.lower() == 'stéphane':
            continue

        # Vérifier habilitation : l'agent doit avoir la section dans ses affectations
        agent_sects = affectations.get(agent, [])
        if section not in agent_sects:
            continue  # jamais habilité physiquement, même en D16

        # B : vacataires interdits en RDC (jamais, même en D16)
        if is_vacataire(agent) and section == 'RDC':
            continue

        # Catégorie : rouge si section hors catégorie sans alerte
        rouge = is_section_rouge(agent, section, affectations, categories)
        # En mode normal (non D16) : les agents A/B ne vont jamais en section rouge
        # → on les exclut sauf si allow_any_section (D16 forcé)
        if rouge and not allow_any_section:
            continue

        if not agent_available(agent, jour, cs, ce, date_str, horaires_agents, evenements):
            continue

        # D6 : max SP hebdo
        if sp_count.get(agent, 0) + slot_min > sp_max_min.get(agent, 99*60):
            continue

        # D15 : Vacataire max journalier (8h)
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
        # (out_of_section remplacé par rouge_score basé sur la catégorie)

        if is_vacataire(agent):
            prim = 0  # O2 exception vacataires : pas de priorité section
        else:
            prim = 0 if (agent_sects and agent_sects[0] == section) else 1  # O2

        # Rouge : agent hors section sans alerte → pénalisé fort (D16)
        rouge_score = 1 if is_section_rouge(agent, section, affectations, categories) else 0

        # O6 : score vacataire
        # - vac_prioritaire=True → vacataire avant régulier (score 0 vs 1)
        # - force_vacataire → idem
        # - sinon → régulier avant vacataire (score 0 vs 1)
        if vac_prioritaire or force_vacataire:
            vac_score = 0 if is_vacataire(agent) else 1
        else:
            vac_score = 1 if is_vacataire(agent) else 0

        # Priorité vacataire : Vacataire 1 avant Vacataire 2
        vac_order = 0
        if is_vacataire(agent):
            vac_order = 0 if '1' in agent else 1

        over_ideal = 1 if over_ideal_consec(agent, cren_idx, creneaux,
                                             day_assignments, max_court) else 0
        # O3 : SP semaine ET SP jour
        sp_semaine = sp_week_count.get(agent, 0)
        sp_jour    = get_sp_today_before(agent, cren_idx, creneaux, day_assignments)
        court_cren = 1 if slot_min <= 60 else 0

        # Score pause méridienne vacataire
        _, mer_score = vacataire_meridienne_ok(agent, cren_idx, creneaux,
                                                day_assignments, horaires_agents, jour)

        # C+D+E : Continuité par blocs horaires
        # Bloc courant pour ce créneau
        bloc_courant = get_bloc_id(jour, cs)
        continuity = 0.0
        if cren_idx > 0:
            prev_cren_name = creneaux[cren_idx - 1][0]
            prev_cs = creneaux[cren_idx - 1][1]
            prev_slot = day_assignments.get(prev_cren_name)
            if prev_slot and isinstance(prev_slot, dict):
                prev_agents = prev_slot.get('assignment', {}).get(section, [])
                if agent in prev_agents:
                    bloc_prev = get_bloc_id(jour, prev_cs)
                    if bloc_prev == bloc_courant:
                        continuity = -1.0  # bonus fort : même bloc → maintenir
                    else:
                        continuity = +0.5  # malus léger : bloc différent → préférer rotation

        candidates.append((
            rouge_score,     # Catégorie : sans-alerte avant rouge (D16)
            vac_score,       # O6 : vacataire prioritaire ou non
            vac_order,       # Vac1 avant Vac2
            prim,            # O2 : section primaire
            sp_semaine,      # E/O3 : équité semaine (prime sur continuité)
            continuity,      # E : intra-bloc = bonus, inter-bloc = malus léger
            over_ideal,      # O7
            mer_score,       # O6 pause méridienne
            sp_jour,         # O3/O4 : équité journée
            court_cren,      # O5
            agent
        ))

    if not candidates:
        return None, False
    candidates.sort()
    best = candidates[0]
    out_of_section_flag = best[0] == 1  # rouge_score==1 → alerte rouge
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
    D15 : calcule le SP attendu pour un vacataire.
    - Journée complète (matin + après-midi) : exactement 8h = 480 min
    - Demi-journée matin (10h-13h) : exactement 3h = 180 min, SP continu sans pause
    - Demi-journée après-midi (14h-19h) : exactement 5h = 300 min, SP continu sans pause
    """
    h = horaires_agents.get(agent, {}).get(jour)
    if h is None:
        return 0
    sm, em, sa, ea = h
    has_matin = sm is not None and em is not None
    has_apm   = sa is not None and ea is not None
    if has_matin and has_apm:
        return 480  # D15 : 8h si journée complète
    elif has_matin:
        return 180  # D15 : 3h si demi-journée matin (10h-13h)
    elif has_apm:
        return 300  # D15 : 5h si demi-journée après-midi (14h-19h)
    return 0

def vac_is_demi_journee(agent, jour, horaires_agents):
    """Retourne le type de journée vacataire : 'full', 'matin', 'apm' ou None."""
    h = horaires_agents.get(agent, {}).get(jour)
    if h is None:
        return None
    sm, em, sa, ea = h
    has_matin = sm is not None and em is not None
    has_apm   = sa is not None and ea is not None
    if has_matin and has_apm:
        return 'full'
    elif has_matin:
        return 'matin'
    elif has_apm:
        return 'apm'
    return None


def plan_week(week_num, week_dates, planning_type_base, samedi_type,
              affectations, categories, horaires_agents, horaires_ouverture,
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

    # ── O6 : détermine si le mercredi est réalisable sans vacataires ──
    def mercredi_realisable_sans_vac(date_str_mer, creneaux_list):
        """Vérifie si toutes les sections RDC/Adulte/MF/Jeunesse peuvent être
        couvertes avec les seuls réguliers présents ce mercredi."""
        reguliers = [a for a in all_agents
                     if not is_vacataire(a)
                     and horaires_agents.get(a, {}).get('Mercredi') is not None]
        for cren_name, cs, ce in creneaux_list:
            if not creneau_is_open(cs, ce, 'Mercredi', horaires_ouverture):
                continue
            for sect in ['RDC', 'Adulte', 'MF']:
                dispo = [a for a in reguliers
                         if sect in affectations.get(a, [])
                         and agent_available(a, 'Mercredi', cs, ce, date_str_mer,
                                             horaires_agents, evenements)]
                if not dispo:
                    return False
        return True

    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        date_str   = date.strftime('%Y-%m-%d')
        is_vac_day = jour.lower() in vac_days

        # O6 : mercredi → vacataires uniquement si planning non réalisable sans eux
        if jour == 'Mercredi' and is_vac_day:
            vac_necessaires_mer = not mercredi_realisable_sans_vac(date_str, creneaux)
        else:
            vac_necessaires_mer = True  # Samedi : toujours OK

        # Agents éligibles ce jour
        if jour == 'Samedi':
            eligible = get_samedi_eligible(roulement_semaine, samedi_type,
                                            affectations, horaires_agents)
        elif is_vac_day:
            if vac_necessaires_mer:
                eligible = [a for a in all_agents
                            if horaires_agents.get(a, {}).get(jour) is not None]
            else:
                # Mercredi réalisable sans vac → exclure les vacataires
                eligible = [a for a in all_agents if not is_vacataire(a)
                            and horaires_agents.get(a, {}).get(jour) is not None]
        else:
            eligible = [a for a in all_agents if not is_vacataire(a)
                        and horaires_agents.get(a, {}).get(jour) is not None]

        # O6 : vacataires présents ce jour ?
        vac_present_jour = any(
            is_vacataire(a) and a in eligible
            and horaires_agents.get(a, {}).get(jour) is not None
            for a in eligible
        ) and is_vac_day

        pt_key            = f"Samedi_{samedi_type}" if jour == 'Samedi' else jour
        base              = planning_type_base.get(pt_key, {})
        max_rot           = JOURS_COURT.get(jour, 4)
        agents_used_today = {s: [] for s in SECTIONS}
        day_assignments   = {}
        vac_day_sp        = defaultdict(int)

        result[jour] = {'_samedi_type': samedi_type if jour == 'Samedi' else None}

        # ── ÉTAPE 0 : pré-assigner chaque vacataire à une section pour la journée ──
        # (journée complète uniquement — demi-journée gérée par D15 post-O8)
        # Chaque vacataire se voit attribuer une section exclusive pour la journée.
        # Les passes A/B utilisent cette info pour leur laisser la priorité.
        vac_section_jour = {}   # agent -> section dédiée pour la journée
        vac_sp_obligatoire = {} # agent -> [(cs, ce), ...]  (non utilisé ici, conservé)

        if is_vac_day:
            vacs_full = sorted(
                [a for a in eligible if is_vacataire(a)
                 and vac_is_demi_journee(a, jour, horaires_agents) == 'full'],
                key=lambda a: (0 if '1' in a else 1)  # Vac1 d'abord
            )
            # A : ordre des sections selon affectations Excel de chaque vacataire
            # B : RDC toujours exclu pour les vacataires
            used_sects = set()
            for vac_ag in vacs_full:
                # Sections selon ordre Excel, sans RDC
                vac_sects_ordered = [s for s in affectations.get(vac_ag, [])
                                     if s != 'RDC' and s not in used_sects]
                # Vérifier que le vacataire est dispo au moins quelques créneaux
                dispo_crens = [
                    (cs, ce) for _, cs, ce in creneaux
                    if creneau_is_open(cs, ce, jour, horaires_ouverture)
                    and agent_available(vac_ag, jour, cs, ce, date_str,
                                        horaires_agents, evenements)
                ]
                if not dispo_crens:
                    continue
                if vac_sects_ordered:
                    vac_section_jour[vac_ag] = vac_sects_ordered[0]
                    used_sects.add(vac_sects_ordered[0])

        for ag in eligible:
            if not is_vacataire(ag):
                continue
            dj = vac_is_demi_journee(ag, jour, horaires_agents)
            if dj == 'matin':
                vac_sp_obligatoire[ag] = [(cs, ce) for _, cs, ce in creneaux
                                           if cs >= 600 and ce <= 780]
            elif dj == 'apm':
                vac_sp_obligatoire[ag] = [(cs, ce) for _, cs, ce in creneaux
                                           if cs >= 840 and ce <= 1140]
            else:
                vac_sp_obligatoire[ag] = []

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
                    # A : si c'est un vacataire avec une section dédiée différente de cette section,
                    # on le saute → O6 le placera dans sa section dédiée
                    # Normaliser le nom pour comparer (le planning type peut avoir VACATAIRE 1 en maj)
                    if is_vacataire(agent):
                        # Trouver la clé correspondante dans vac_section_jour
                        _vag_key = next((k for k in vac_section_jour
                                         if k.lower() == agent.lower()), None)
                        if _vag_key and vac_section_jour[_vag_key] != section:
                            continue  # O6 gère ce vacataire dans sa section dédiée
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
                            categories=categories,
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
                            vac_prioritaire=vac_present_jour,
                        )
                        if repl:
                            final.append(repl)
                            assigned_this.add(repl)

                assignment[section] = final

            # ══ PASSE A2 : optimisation section primaire (O2) ══
            # Si un agent est en section non-primaire ET sa section primaire est vide → le déplacer
            for section in ['RDC', 'Adulte', 'MF']:
                agents_here = list(assignment.get(section, []))
                for agent in agents_here:
                    if is_vacataire(agent):
                        continue
                    sects = affectations.get(agent, [])
                    if not sects or sects[0] == section:
                        continue  # déjà en section primaire
                    prim_sect = sects[0]
                    if prim_sect not in ['RDC', 'Adulte', 'MF']:
                        continue  # section primaire = Jeunesse ou autre
                    if assignment.get(prim_sect):
                        continue  # section primaire déjà couverte
                    # Déplacer l'agent vers sa section primaire
                    assignment[section].remove(agent)
                    assignment[prim_sect] = [agent]
                    # La section abandonnée sera couverte par Passe B

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
                            categories=categories,
                    horaires_agents=horaires_agents,
                    evenements=evenements,
                    sp_count=sp_count, sp_max_min=sp_max_min,
                    sp_week_count=dict(sp_count),
                    creneaux=creneaux, cren_idx=cren_idx,
                    day_assignments=day_assignments,
                    max_court=max_court, max_long=max_long,
                    exclude=assigned_this,
                    vac_day_sp=vac_day_sp,
                    vac_prioritaire=vac_present_jour,
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

            # Compléter si insuffisant — O2 (section primaire Jeunesse) + O3 (équité SP) + O5 (≥1h)
            if len(assignment['Jeunesse']) < j_needs:
                # Construire liste triée : primaire Jeunesse d'abord, puis O3 SP semaine
                jeunesse_candidates = []
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
                    if is_vacataire(agent) and vac_day_sp.get(agent, 0) + slot_min > 480:
                        continue
                    if violates_consec_hard(agent, jour, cs, ce, cren_idx, creneaux,
                                            day_assignments, max_court, max_long,
                                            'Jeunesse', date_str, affectations,
                                            evenements, horaires_agents):
                        continue
                    # Pause méridienne obligatoire (D14 réguliers + O6 vacataires)
                    if not agent_has_meridienne_pause(agent, cren_idx, creneaux,
                                                       day_assignments, horaires_agents, jour):
                        continue
                    sects = affectations.get(agent, [])
                    # O2 : section primaire Jeunesse (0=oui, 1=non)
                    prim_j = 0 if (sects and sects[0] == 'Jeunesse') else 1
                    # Catégorie : si Jeunesse est rouge pour cet agent → D16 uniquement
                    # en Passe C (remplissage normal) on n'utilise pas les agents rouge
                    if is_section_rouge(agent, 'Jeunesse', affectations, categories):
                        continue  # sera utilisé seulement si vide en Passe E
                    # B : vacataires interdits en RDC (Jeunesse OK, pas de check ici)
                    # O6 : vacataire avec section dédiée → ne pas le "consommer" en Jeunesse
                    # sauf s'il est lui-même dédié à Jeunesse
                    if is_vacataire(agent) and agent in vac_section_jour:
                        if vac_section_jour[agent] != 'Jeunesse':
                            continue  # réservé pour sa section dédiée, Passe O6 le placera
                    # O3 : SP semaine (prime sur continuité)
                    sp_s = sp_count.get(agent, 0)
                    # C+D+E : continuité par blocs
                    bloc_courant = get_bloc_id(jour, cs)
                    cont = 0.0
                    if cren_idx > 0:
                        prev_cn  = creneaux[cren_idx-1][0]
                        prev_cs2 = creneaux[cren_idx-1][1]
                        prev_s   = day_assignments.get(prev_cn)
                        if prev_s and agent in prev_s.get('assignment', {}).get('Jeunesse', []):
                            bloc_prev = get_bloc_id(jour, prev_cs2)
                            cont = -1.0 if bloc_prev == bloc_courant else +0.5
                    # O5 : pénaliser ≤30min sauf si intra-bloc contigu
                    court = 1 if (slot_min <= 30 and cont >= 0) else 0
                    jeunesse_candidates.append((prim_j, sp_s, cont, court, agent))
                jeunesse_candidates.sort()
                for _, _, _, _, agent in jeunesse_candidates:
                    if len(assignment['Jeunesse']) >= j_needs:
                        break
                    assignment['Jeunesse'].append(agent)
                    assigned_this.add(agent)

            # D16 si encore insuffisant (dernier recours)
            if len(assignment['Jeunesse']) < j_needs:
                needed = j_needs - len(assignment['Jeunesse'])
                for agent in sorted(eligible, key=lambda a: (
                    1 if is_section_rouge(a,'Jeunesse',affectations,categories) else 0,
                    sp_count.get(a,0)
                )):
                    if needed <= 0:
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
                    if not agent_has_meridienne_pause(agent, cren_idx, creneaux,
                                                       day_assignments, horaires_agents, jour):
                        continue
                    is_rouge = is_section_rouge(agent,'Jeunesse',affectations,categories)
                    assignment['Jeunesse'].append(agent)
                    assigned_this.add(agent)
                    if is_rouge:
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
                            categories=categories,
                            horaires_agents=horaires_agents,
                            evenements=evenements,
                            sp_count=sp_count, sp_max_min=sp_max_min,
                            sp_week_count=dict(sp_count),
                            creneaux=creneaux, cren_idx=cren_idx,
                            day_assignments=day_assignments,
                            max_court=max_court, max_long=max_long,
                            exclude=assigned_this,
                            vac_day_sp=vac_day_sp,
                            vac_prioritaire=vac_present_jour,
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
                            categories=categories,
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
                        vac_prioritaire=vac_present_jour,
                    )
                    if repl:
                        assignment[section] = [repl]
                        assigned_this.add(repl)
                        slot_out_section[section] = True
                    else:
                        slot_alerts.append(f"ALERTE — {section} : aucun agent disponible")

            # ══ PASSE O6 : placer vacataire dans sa section dédiée (8h/jour) ══
            # Utilise vac_section_jour défini en ÉTAPE 0.
            # Le vacataire remplace le régulier dans sa section dédiée.
            # Le régulier éjecté est récupéré par O8 si son min SP n'est pas atteint.
            if is_vac_day:
                for vac_ag, sect_ded in vac_section_jour.items():
                    if not agent_available(vac_ag, jour, cs, ce, date_str,
                                            horaires_agents, evenements):
                        continue
                    # D15 : ne pas dépasser 8h
                    if vac_day_sp.get(vac_ag, 0) + slot_min > 480:
                        continue
                    # Pause méridienne obligatoire (au moins 1h dans 12h-14h)
                    if not agent_has_meridienne_pause(vac_ag, cren_idx, creneaux,
                                                       day_assignments, horaires_agents, jour):
                        continue
                    # Déjà assigné ce créneau ?
                    if any(vac_ag in assignment.get(s, []) for s in SECTIONS):
                        continue
                    # Section dédiée : placer ou évincer
                    curr = assignment.get(sect_ded, [])
                    placed = False
                    # Capacité maximale de la section dédiée pour ce créneau
                    if sect_ded == 'Jeunesse':
                        sect_cap = j_needs  # D4 : ne pas dépasser le besoin Jeunesse
                    else:
                        sect_cap = 1        # D3 : max 1 agent pour RDC/Adulte/MF

                    if not curr:
                        # Section vide : ajouter directement
                        if sect_ded == 'Jeunesse' and j_needs == 0:
                            pass  # D4 : pas de besoin Jeunesse ce créneau
                        else:
                            assignment[sect_ded].append(vac_ag)
                            assigned_this.add(vac_ag)
                            placed = True
                    elif sect_ded == 'Jeunesse' and len(curr) < j_needs:
                        # Jeunesse : ajouter si besoin non encore comblé
                        assignment['Jeunesse'].append(vac_ag)
                        assigned_this.add(vac_ag)
                        placed = True
                    elif sect_ded == 'Jeunesse' and len(curr) < j_needs:
                        # Jeunesse incomplète → ajouter le vacataire directement
                        assignment['Jeunesse'].append(vac_ag)
                        assigned_this.add(vac_ag)
                        placed = True
                    elif sect_ded == 'Jeunesse' and len(curr) >= j_needs                             and j_needs >= 1                             and not any(is_vacataire(a) for a in curr):
                        # Jeunesse pleine de réguliers → remplacer 1 régulier par le vacataire
                        # Choisir le régulier le plus facile à déplacer (le plus polyvalent)
                        best_displaced = None
                        for cand in sorted(curr,
                                key=lambda a: -len(affectations.get(a,[]))):
                            d_sects = affectations.get(cand, [])
                            # Peut-il aller ailleurs ?
                            can_go = any(
                                s2 in d_sects and not assignment.get(s2)
                                and agent_available(cand, jour, cs, ce, date_str,
                                                    horaires_agents, evenements)
                                for s2 in ['RDC', 'Adulte', 'MF']
                            )
                            if can_go or True:  # toujours évincer si besoin
                                best_displaced = cand
                                break
                        if best_displaced:
                            d_sects = affectations.get(best_displaced, [])
                            relocated = False
                            for s2 in ['RDC', 'Adulte', 'MF']:
                                if s2 not in d_sects: continue
                                if assignment.get(s2): continue
                                if not agent_available(best_displaced, jour, cs, ce,
                                                        date_str, horaires_agents, evenements):
                                    continue
                                assignment[s2] = [best_displaced]
                                assignment['Jeunesse'] = [a for a in assignment['Jeunesse']
                                                          if a != best_displaced] + [vac_ag]
                                assigned_this.add(vac_ag)
                                placed = True
                                relocated = True
                                break
                            if not placed:
                                # Évincer sans relocalisation (O8 récupère)
                                assignment['Jeunesse'] = [a for a in assignment['Jeunesse']
                                                          if a != best_displaced] + [vac_ag]
                                assigned_this.add(vac_ag)
                                placed = True
                    elif len(curr) == 1 and not is_vacataire(curr[0])                             and sect_ded != 'Jeunesse':
                        displaced = curr[0]
                        if displaced.lower() == 'stéphane':
                            continue  # D17
                        d_sects = affectations.get(displaced, [])
                        # Essai 1 : relocaliser dans une autre section libre
                        for s2 in ['RDC', 'Adulte', 'MF']:
                            if s2 == sect_ded: continue
                            if s2 not in d_sects: continue
                            if assignment.get(s2): continue
                            if not agent_available(displaced, jour, cs, ce, date_str,
                                                    horaires_agents, evenements):
                                continue
                            assignment[s2]       = [displaced]
                            assignment[sect_ded] = [vac_ag]
                            assigned_this.add(vac_ag)
                            placed = True
                            break
                        # Essai 2 : évincer (vacataire prioritaire absolu — O8 récupèrera le régulier)
                        if not placed:
                            assignment[sect_ded] = [vac_ag]
                            assigned_this.add(vac_ag)
                            placed = True

            # ══ Nettoyage alertes obsolètes Jeunesse ══
            # Recalculer après toutes les passes (Passe C peut avoir alerté prématurément)
            j_ag_final = assignment.get('Jeunesse', [])
            slot_alerts = [al for al in slot_alerts if 'Jeunesse' not in al]
            if len(j_ag_final) < j_needs:
                slot_alerts.append(
                    f"ALERTE — Jeunesse : {len(j_ag_final)}/{j_needs} agents")

            # ══ PASSE D2 : re-vérification D11 post-O6 (vacataire seul Jeunesse) ══
            j_ags_final = assignment.get('Jeunesse', [])
            if j_ags_final and all(is_vacataire(a) for a in j_ags_final):
                if not (cs >= exc_vac_s and ce <= exc_vac_e):
                    found_reg = None
                    for cand in sorted(eligible, key=lambda a: sp_count.get(a, 0)):
                        if is_vacataire(cand): continue
                        if cand in assigned_this: continue
                        if 'Jeunesse' not in affectations.get(cand, []): continue
                        if not agent_available(cand, jour, cs, ce, date_str,
                                                horaires_agents, evenements): continue
                        if sp_count.get(cand,0)+slot_min > sp_max_min.get(cand,99*60): continue
                        if violates_consec_hard(cand, jour, cs, ce, cren_idx, creneaux,
                                                day_assignments, max_court, max_long,
                                                'Jeunesse', date_str, affectations,
                                                evenements, horaires_agents): continue
                        if not agent_has_meridienne_pause(cand, cren_idx, creneaux,
                                                           day_assignments, horaires_agents, jour):
                            continue
                        found_reg = cand
                        break
                    if found_reg:
                        if len(assignment['Jeunesse']) >= j_needs:
                            assignment['Jeunesse'] = [found_reg]
                        else:
                            assignment['Jeunesse'].insert(0, found_reg)
                        assigned_this.add(found_reg)
                    else:
                        assignment['Jeunesse'] = []
                        slot_alerts = [al for al in slot_alerts if 'Jeunesse' not in al]
                        slot_alerts.append(
                            "ALERTE — Jeunesse : vacataire seul hors 12h-14h, aucun régulier dispo")

            # ══ RÈGLE : vacataire seul max 3h en Adulte/MF ══
            VAC_SEUL_MAX = 180  # 3h en minutes
            for section in ['Adulte', 'MF']:
                sect_ags = assignment.get(section, [])
                if len(sect_ags) == 1 and is_vacataire(sect_ags[0]):
                    # Calculer durée consécutive vacataire seul dans cette section
                    consec_seul = slot_min
                    for prev_cn, prev_cs_v, prev_ce_v in reversed(creneaux[:cren_idx]):
                        prev_sl = day_assignments.get(prev_cn)
                        if not prev_sl: break
                        if prev_ce_v < cs - 1: break  # gap → stop
                        prev_sect_ags = prev_sl.get('assignment', {}).get(section, [])
                        if len(prev_sect_ags) == 1 and is_vacataire(prev_sect_ags[0]):
                            consec_seul += prev_ce_v - prev_cs_v
                        else:
                            break
                    if consec_seul > VAC_SEUL_MAX:
                        # Chercher un régulier à ajouter
                        repl, _ = find_replacement(
                            section=section, jour=jour, cs=cs, ce=ce,
                            date_str=date_str, eligible=eligible,
                            affectations=affectations, categories=categories,
                            horaires_agents=horaires_agents, evenements=evenements,
                            sp_count=sp_count, sp_max_min=sp_max_min,
                            sp_week_count=dict(sp_count),
                            creneaux=creneaux, cren_idx=cren_idx,
                            day_assignments=day_assignments,
                            max_court=max_court, max_long=max_long,
                            exclude=assigned_this, vac_day_sp=vac_day_sp,
                        )
                        if repl and not is_vacataire(repl):
                            assignment[section] = [repl]
                            assigned_this.add(repl)
                        else:
                            slot_alerts.append(
                                f"ALERTE — {section} : vacataire seul >3h, aucun régulier dispo")

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
    # O8 : RECALAGE FINAL — Min SP réguliers
    # Skip si un vacataire est disponible sur ce créneau (O6 priorité vac)
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
            is_vac_day_o8 = jour.lower() in vac_days
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
                # O6/O8 : si un vacataire est disponible ce créneau, ne pas
                # utiliser un régulier (le vacataire prime)
                if is_vac_day_o8:
                    vac_dispo_o8 = any(
                        is_vacataire(a) and a in eligible
                        and agent_available(a, jour, cs, ce, date_str,
                                            horaires_agents, evenements)
                        for a in all_agents if is_vacataire(a)
                    )
                    if vac_dispo_o8:
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

    # ══════════════════════════════════════════════════════════
    # D15 POST-O8 : forcer créneaux obligatoires demi-journée vacataires
    # Demi-journée matin (10h-13h) ou après-midi (14h-19h) : SP continu, sans pause
    # ══════════════════════════════════════════════════════════
    for jour in JOURS_SP:
        date = week_dates.get(jour)
        if date is None:
            continue
        is_vac_day_d15 = jour.lower() in vac_days
        if not is_vac_day_d15:
            continue
        date_str_d15 = date.strftime('%Y-%m-%d')
        for ag in all_agents:
            if not is_vacataire(ag):
                continue
            dj = vac_is_demi_journee(ag, jour, horaires_agents)
            if dj is None or dj == 'full':
                continue  # journée complète : géré par O6
            # Récupérer les créneaux obligatoires non encore assignés
            for cren_name, cs, ce in creneaux:
                if dj == 'matin' and not (cs >= 600 and ce <= 780):
                    continue
                if dj == 'apm' and not (cs >= 840 and ce <= 1140):
                    continue
                slot = result.get(jour, {}).get(cren_name)
                if not slot or not isinstance(slot, dict) or not slot.get('open'):
                    continue
                if not creneau_is_open(cs, ce, jour, horaires_ouverture):
                    continue
                if not agent_available(ag, jour, cs, ce, date_str_d15,
                                        horaires_agents, evenements):
                    continue
                assignment_d15 = slot.get('assignment', {})
                already = any(ag in assignment_d15.get(s, []) for s in SECTIONS)
                if already:
                    continue
                # Chercher une section disponible
                for sect in ['Adulte', 'MF', 'Jeunesse', 'RDC']:
                    if sect not in affectations.get(ag, []):
                        continue
                    if sect == 'Jeunesse':
                        jn = get_besoins_jeunesse_slot(
                            besoins_jeunesse, cren_name, jour,
                            result[jour].get('_samedi_type', 'ROUGE'), semaine_type)
                        if len(assignment_d15['Jeunesse']) >= jn:
                            continue
                    else:
                        if len(assignment_d15.get(sect, [])) >= 1:
                            continue
                    assignment_d15[sect].append(ag)
                    sp_count[ag] += (ce - cs)
                    sp_jour_cum[jour][ag] += (ce - cs)
                    break

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

    # Vérification min SP vacataires (D15)
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
            vac_target = compute_vac_min_sp(agent, jour, horaires_agents)
            actual     = sp_jour_cum[jour].get(agent, 0)
            if vac_target > 0 and actual > 0:
                if actual < vac_target:
                    sp_alerts[f"{agent}_{jour}"] = (
                        f"{agent} {jour} SP insuffisant : "
                        f"{actual//60}h{actual%60:02d} / cible {vac_target//60}h{vac_target%60:02d}"
                    )
                elif actual > vac_target:
                    sp_alerts[f"{agent}_{jour}_max"] = (
                        f"{agent} {jour} SP dépassé : "
                        f"{actual//60}h{actual%60:02d} / max {vac_target//60}h{vac_target%60:02d}"
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
    affectations, categories = parse_affectations(raw)
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
            categories         = categories,
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
