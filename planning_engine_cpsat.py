"""
planning_engine_cpsat.py
Moteur de calcul planning médiathèque — CP-SAT (Google OR-Tools)
Remplace l'ancien moteur glouton.
"""

import datetime
import re
from collections import defaultdict

import openpyxl
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════
#  CONSTANTES
# ══════════════════════════════════════════════════════════════

SECTIONS     = ['RDC', 'Adulte', 'MF', 'Jeunesse']
JOURS_SEMAINE = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
JOURS_VAC    = {'Mercredi', 'Samedi'}   # jours où les vacataires peuvent travailler

# Poids des contraintes molles (plus = plus prioritaire)
POIDS = {
    'G1_planning_type':        100,   # respecter le planning type
    'G2_meme_section_repl':     50,   # remplaçant même section principale
    'H2_equite_journee':        40,   # équité heures sur la journée
    'H1_equite_semaine':        30,   # équité heures sur la semaine
    'J1_section_principale':    30,   # agent dans sa section principale
    'J3_responsable':           25,   # responsables déprioritisés
    'I1_non_fragmentation':     20,   # blocs continus préférés
    'I2_un_seul_remplacant':    20,   # 1 remplaçant par bloc absent
    'H3_equite_remplacement':   15,   # équité des remplacements
    'K1_vac_dernier_recours':   10,   # vacataires en dernier recours
    'K2_vac1_avant_vac2':        5,   # Vacataire 1 avant Vacataire 2
}


# ══════════════════════════════════════════════════════════════
#  UTILITAIRES TEMPS
# ══════════════════════════════════════════════════════════════

def hhmm_to_min(t):
    """Convertit un objet time, une string 'HH:MM' ou 'HH:MM:SS' en minutes depuis minuit."""
    if t is None or (isinstance(t, str) and t.strip() in ('', '\xa0')):
        return None
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute
    if isinstance(t, str):
        t = t.strip().replace('h', ':').replace('H', ':')
        parts = t.split(':')
        return int(parts[0]) * 60 + int(parts[1]) if len(parts) >= 2 else None
    return None


def parse_creneau(s):
    """Convertit '10:00-12:30' ou '10H-12H30' en (cs, ce) en minutes. None si invalide."""
    if not s:
        return None
    s = str(s).strip()
    if '-' not in s:
        return None

    def to_min(t):
        t = t.strip().upper()
        m = re.match(r'(\d{1,2})H(\d{0,2})', t)
        if m:
            return int(m.group(1)) * 60 + (int(m.group(2)) if m.group(2) else 0)
        if ':' in t:
            p = t.split(':')
            try:
                return int(p[0]) * 60 + int(p[1])
            except (ValueError, IndexError):
                return None
        return None

    parts = s.split('-', 1)
    cs = to_min(parts[0])
    ce = to_min(parts[1])
    if cs is None or ce is None or ce <= cs:
        return None
    return (cs, ce)


def load_excel_data(filepath):
    """Charge le fichier Excel et retourne un dict {nom_onglet: worksheet}."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    return {ws.title: ws for ws in wb.worksheets}


def parse_parametres(raw):
    """
    Retourne un dict avec les paramètres du mois :
    mois, annee, creneaux (liste de (cs,ce)), samedis {1:'ROUGE',...},
    semaines {1:'Hors Vacances scolaires',...}, mode_vacataires (set de jours)
    """
    ws = raw['Paramètres']
    params = {}
    for row in ws.iter_rows(values_only=True):
        if row[0] and row[1] is not None:
            params[str(row[0]).strip()] = row[1]

    # Créneaux : liste de (cs, ce)
    creneaux_str = str(params.get('Liste_des_créneaux', '')).split(';')
    creneaux = []
    for c in creneaux_str:
        parsed = parse_creneau(c.strip())
        if parsed:
            creneaux.append(parsed)

    # Samedis
    samedis = {}
    for i in range(1, 6):
        v = params.get(f'Samedi_{i}')
        if v:
            samedis[i] = str(v).strip().upper()

    # Semaines
    semaines = {}
    for i in range(1, 6):
        v = params.get(f'Semaine_{i}')
        if v:
            semaines[i] = str(v).strip()

    # Mode vacataires
    mode_vac_str = str(params.get('Mode_vacataires', 'samedi')).lower()
    mode_vac = set()
    for j in JOURS_SEMAINE:
        if j.lower() in mode_vac_str:
            mode_vac.add(j)

    return {
        'mois':     str(params.get('Mois', '')).strip(),
        'annee':    int(params.get('Année', 2026)),
        'creneaux': creneaux,
        'samedis':  samedis,
        'semaines': semaines,
        'mode_vac': mode_vac,
    }


def parse_affectations(raw):
    """
    Retourne :
    - affectations  : {agent: [section1, section2, ...]}  (ordre = priorité)
    - categories    : {agent: 'A' | None}  (A = sections 1 et 2 équivalentes)
    - responsables  : set d'agents responsables
    - pause_flexible: set d'agents avec pause flexible
    """
    ws = raw['Affectations']
    affectations = {}
    categories   = {}
    responsables = set()
    pause_flex   = set()

    header_skipped = False
    for row in ws.iter_rows(values_only=True):
        if not row[0]:
            continue
        agent = str(row[0]).strip()
        if agent == 'Agent':
            header_skipped = True
            continue
        if not header_skipped:
            continue

        sections = [str(row[i]).strip() for i in range(2, 6) if row[i] and str(row[i]).strip()]
        affectations[agent]  = sections
        categories[agent]    = str(row[1]).strip() if row[1] else None
        if row[6] and str(row[6]).strip() == 'OUI':
            responsables.add(agent)
        if row[7] and str(row[7]).strip() == 'OUI':
            pause_flex.add(agent)

    return affectations, categories, responsables, pause_flex


def parse_horaires_agents(raw):
    """
    Retourne {agent: {jour: (debut_matin, fin_matin, debut_apm, fin_apm)}}
    Toutes les valeurs en minutes depuis minuit (ou None si absent).
    """
    ws = raw['Horaires_Des_Agents']
    horaires = defaultdict(dict)
    header_skipped = False

    for row in ws.iter_rows(values_only=True):
        if not row[0]:
            continue
        if str(row[0]).strip() == 'Agent':
            header_skipped = True
            continue
        if not header_skipped:
            continue

        agent = str(row[0]).strip()
        jour  = str(row[1]).strip() if row[1] else None
        if not jour:
            continue

        dm = hhmm_to_min(row[2])
        fm = hhmm_to_min(row[3])
        da = hhmm_to_min(row[4])
        fa = hhmm_to_min(row[5])

        horaires[agent][jour] = (dm, fm, da, fa)

    return dict(horaires)


def parse_roulement_samedi(raw):
    """
    Retourne :
    - roulement_type : {agent: 'ROUGE' | 'BLEU'}
    - exceptions     : {semaine_num: {agent: 'ROUGE' | 'BLEU'}}
    """
    ws = raw['Roulement_Samedi']
    roulement_type = {}
    exceptions     = defaultdict(dict)

    mode = None   # 'type' ou 'exceptions'
    current_sem = None

    for row in ws.iter_rows(values_only=True):
        if not any(c for c in row):
            continue
        c0 = str(row[0] or '').strip()
        c1 = str(row[1] or '').strip()
        c2 = str(row[2] or '').strip()

        if c0 == 'Roulement type':
            mode = 'type'
            continue
        if c0 == 'Exceptions par semaine':
            mode = 'exceptions'
            current_sem = None
            continue

        if mode == 'type':
            if c1 and c1 != 'Agent' and c2 and c2 != 'Roulement':
                roulement_type[c1] = c2.upper()

        elif mode == 'exceptions':
            if c1 == 'Semaine':
                continue
            # Numéro de semaine
            if row[1] is not None and str(row[1]).strip().isdigit():
                current_sem = int(str(row[1]).strip())
            # Agent + roulement
            agent = str(row[2] or '').strip().rstrip()
            roul  = str(row[3] or '').strip().upper() if len(row) > 3 else ''
            if agent and roul and current_sem is not None:
                exceptions[current_sem][agent] = roul

    return roulement_type, dict(exceptions)


def parse_besoins_jeunesse(raw):
    """
    Retourne {periode: {jour_cle: {creneau_str: nb_agents}}}
    periode = 'Hors Vacances scolaires' | 'Vacances Scolaires'
    jour_cle = 'Mardi' | 'Mercredi' | 'Jeudi' | 'Vendredi' | 'Samedi_rouge' | 'samedi bleu'
    """
    ws = raw['Besoins_Jeunesse']
    result = {}
    current_periode = None
    headers = []

    for row in ws.iter_rows(values_only=True):
        if not any(c for c in row):
            continue
        c0 = str(row[0] or '').strip()
        c1 = str(row[1] or '').strip()

        # Ligne de période
        if 'Vacances' in c0 or 'Hors' in c0:
            current_periode = c0.rstrip()
            result[current_periode] = {}
            headers = []
            continue

        # Ligne d'en-têtes
        if c1 == 'Créneau' or c1 == 'créneau':
            headers = [str(row[i] or '').strip() for i in range(2, 8)]
            for h in headers:
                if h:
                    result[current_periode][h] = {}
            continue

        # Ligne de données
        if current_periode and headers and c1 and ':' in c1:
            cren = c1.strip()
            for i, h in enumerate(headers):
                if h and row[i + 2] is not None:
                    try:
                        result[current_periode][h][cren] = int(row[i + 2])
                    except (ValueError, TypeError):
                        pass

    return result


def parse_evenements(raw):
    """
    Retourne liste de {date_str, cs, ce, nom, agents: []}
    """
    ws = raw['Événements']
    events = []
    header_skipped = False

    for row in ws.iter_rows(values_only=True):
        if not row[0]:
            continue
        if str(row[0]).strip() in ('Date', 'date'):
            header_skipped = True
            continue
        if not header_skipped:
            continue

        date_val = row[0]
        if isinstance(date_val, datetime.datetime):
            date_str = date_val.strftime('%Y-%m-%d')
        elif isinstance(date_val, datetime.date):
            date_str = date_val.strftime('%Y-%m-%d')
        else:
            continue

        cs = hhmm_to_min(row[1])
        ce = hhmm_to_min(row[2])
        nom = str(row[3] or '').strip()
        agents_str = str(row[4] or '').strip()
        agents = [a.strip() for a in re.split(r'[,;/]', agents_str) if a.strip()] if agents_str else []

        if cs is not None and ce is not None:
            events.append({'date': date_str, 'cs': cs, 'ce': ce, 'nom': nom, 'agents': agents})

    return events


def parse_planning_type(raw):
    """
    Retourne {jour: {creneau_str: {section: [agents]}}}
    jour = 'Mardi' | 'Mercredi' | 'Jeudi' | 'Vendredi' | 'Samedi_ROUGE' | 'Samedi_BLEU'
    """
    ws = raw['planning_type']
    result   = {}
    cur_jour = None
    sections = []

    # Mapping colonnes → sections
    COL_SECTIONS = {2: 'RDC', 3: 'Adulte', 5: 'MF'}  # col index 0-based

    for row in ws.iter_rows(values_only=True):
        if not any(c for c in row):
            continue
        c0 = str(row[0] or '').strip()
        c1 = str(row[1] or '').strip()

        # Détection du jour
        if c0 in ('MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI'):
            jour_map = {'MARDI': 'Mardi', 'MERCREDI': 'Mercredi',
                        'JEUDI': 'Jeudi', 'VENDREDI': 'Vendredi'}
            cur_jour = jour_map[c0]
            result.setdefault(cur_jour, {})
            continue

        # Samedis
        if 'SAMEDI' in c0 and 'ROUGE' in c0.upper():
            cur_jour = 'Samedi_ROUGE'
            result.setdefault(cur_jour, {})
            continue
        if 'SAMEDI' in c0 and ('BLEU' in c0.upper() or 'BLEUE' in c0.upper()):
            cur_jour = 'Samedi_BLEU'
            result.setdefault(cur_jour, {})
            continue

        if cur_jour is None or not c1:
            continue

        # Ignorer lignes de durées (nombres)
        if isinstance(row[1], (int, float)):
            continue

        # Ligne d'agents : c1 contient un créneau style '10H-12H30'
        if '-' in c1 and 'H' in c1.upper():
            cren_str = c1.strip()
            result[cur_jour].setdefault(cren_str, {'RDC': [], 'Adulte': [], 'MF': [], 'Jeunesse': []})
            for col_idx, section in COL_SECTIONS.items():
                val = row[col_idx] if col_idx < len(row) else None
                if val and str(val).strip():
                    agents = [a.strip() for a in re.split(r'[/,]', str(val)) if a.strip()]
                    result[cur_jour][cren_str][section].extend(agents)

            # Jeunesse : colonne non fixe — on cherche dans les colonnes restantes
            # (colonnes 6-9 selon la structure du fichier)
            for col_idx in range(6, min(10, len(row))):
                val = row[col_idx]
                if val and str(val).strip():
                    agents = [a.strip() for a in re.split(r'[/,]', str(val)) if a.strip()]
                    result[cur_jour][cren_str]['Jeunesse'].extend(agents)

    return result


# ══════════════════════════════════════════════════════════════
#  CONSTRUCTION DU CALENDRIER DU MOIS
# ══════════════════════════════════════════════════════════════

def build_calendar(mois_str, annee, samedis_params):
    """
    Retourne une liste de semaines. Chaque semaine est un dict :
    {
      'num': 1,
      'jours': [
        {'date': '2026-05-05', 'jour': 'Mardi', 'samedi_type': None},
        ...
        {'date': '2026-05-09', 'jour': 'Samedi', 'samedi_type': 'ROUGE'},
      ]
    }
    """
    MOIS_FR = {
        'janvier': 1, 'février': 2, 'mars': 3, 'avril': 4,
        'mai': 5, 'juin': 6, 'juillet': 7, 'août': 8,
        'septembre': 9, 'octobre': 10, 'novembre': 11, 'décembre': 12
    }
    mois_num = MOIS_FR.get(mois_str.lower(), 5)
    premier  = datetime.date(annee, mois_num, 1)

    # Trouver le premier mardi du mois
    jours_fr = {0: 'Lundi', 1: 'Mardi', 2: 'Mercredi', 3: 'Jeudi',
                4: 'Vendredi', 5: 'Samedi', 6: 'Dimanche'}
    jours_sp = ['Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

    # Collecter tous les jours SP du mois
    semaines = []
    current_week = []
    week_num = 1
    sam_count = 0

    d = premier
    while d.month == mois_num:
        jour_fr = jours_fr[d.weekday()]
        if jour_fr in jours_sp:
            sam_type = None
            if jour_fr == 'Samedi':
                sam_count += 1
                sam_type = samedis_params.get(sam_count)

            current_week.append({
                'date': d.strftime('%Y-%m-%d'),
                'jour': jour_fr,
                'samedi_type': sam_type,
            })

            # Fin de semaine après le samedi
            if jour_fr == 'Samedi':
                semaines.append({'num': week_num, 'jours': current_week})
                current_week = []
                week_num += 1

        d += datetime.timedelta(days=1)

    # Semaine incomplète sans samedi (rare en début de mois)
    if current_week:
        semaines.append({'num': week_num, 'jours': current_week})

    return semaines


# ══════════════════════════════════════════════════════════════
#  DISPONIBILITÉ DES AGENTS
# ══════════════════════════════════════════════════════════════

def is_vacataire(agent):
    return 'Vacataire' in agent or 'vacataire' in agent


def agent_disponible(agent, jour, cs, ce, horaires_agents, evenements,
                     date_str, pause_flex):
    """
    Retourne True si l'agent peut être placé sur ce créneau (cs, ce) ce jour-là.
    Vérifie : horaires contractuels, pause contractuelle, événements bloquants.
    """
    # Vacataires : pas de contrainte horaire contractuelle individuelle
    if not is_vacataire(agent):
        h = horaires_agents.get(agent, {}).get(jour)
        if not h:
            return False  # pas de contrat ce jour

        dm, fm, da, fa = h

        # Vérifier que le créneau est dans les heures de travail
        dans_matin = (dm is not None and fm is not None and cs >= dm and ce <= fm)
        dans_apm   = (da is not None and fa is not None and cs >= da and ce <= fa)

        if not (dans_matin or dans_apm):
            # Avec pause flexible : le créneau peut chevaucher la pause
            if agent in pause_flex and dm is not None and fa is not None:
                if cs >= dm and ce <= fa:
                    # OK : dans la plage globale mais en pause
                    # Vérifier que la pause sera ≥ 1h (B2)
                    # On accepte pour l'instant — la contrainte 1h sera vérifiée globalement
                    pass
                else:
                    return False
            else:
                return False

        # Pause contractuelle (sans pause flexible) : ne pas placer pendant la pause
        if agent not in pause_flex:
            if dm is not None and fm is not None and da is not None:
                en_pause = (cs >= fm and ce <= da)
                if en_pause:
                    return False

    # Événements bloquants
    for ev in evenements:
        if ev['date'] != date_str:
            continue
        if ev['agents'] and agent not in ev['agents']:
            continue
        # Chevauchement
        if cs < ev['ce'] and ce > ev['cs']:
            return False

    return True


# ══════════════════════════════════════════════════════════════
#  MOTEUR CP-SAT — UNE JOURNÉE
# ══════════════════════════════════════════════════════════════

def solve_day(jour, date_str, creneaux_ouverts, agents_eligibles,
              affectations, categories, responsables, pause_flex,
              horaires_agents, evenements, besoins_jeunesse,
              planning_type_jour, roulement_agents,
              samedi_type=None, periode='Hors Vacances scolaires',
              mode_vac=None):
    """
    Résout le planning d'une journée avec CP-SAT.

    Paramètres :
    - creneaux_ouverts : liste de (cs, ce)
    - agents_eligibles : liste d'agents pouvant travailler ce jour
    - planning_type_jour : {creneau_str: {section: [agents]}}
    - roulement_agents : {agent: 'ROUGE'|'BLEU'} pour ce jour (samedi)

    Retourne : {creneau_idx: {section: [agents]}} ou None si infaisable
    """
    if mode_vac is None:
        mode_vac = JOURS_VAC

    model  = cp_model.CpModel()
    agents = list(agents_eligibles)
    n_cren = len(creneaux_ouverts)

    # ── Variables de décision ──────────────────────────────────
    # x[a][c][s] = 1 si agent a travaille au créneau c en section s
    x = {}
    for a in agents:
        for c in range(n_cren):
            for s in SECTIONS:
                x[a, c, s] = model.new_bool_var(f'x_{a}_{c}_{s}')

    # ══ CONTRAINTES DURES ════════════════════════════════════

    # A2 : vacataires jamais en RDC
    for a in agents:
        if is_vacataire(a):
            for c in range(n_cren):
                model.add(x[a, c, 'RDC'] == 0)

    # A3 : Stéphane uniquement MF
    if 'Stéphane' in agents:
        for c in range(n_cren):
            for s in SECTIONS:
                if s != 'MF':
                    model.add(x['Stéphane', c, s] == 0)

    # A1 : sections habilitées uniquement
    for a in agents:
        sects_ok = set(affectations.get(a, []))
        for c in range(n_cren):
            for s in SECTIONS:
                if s not in sects_ok:
                    model.add(x[a, c, s] == 0)

    # A4 : max 1 agent par section/créneau pour RDC, Adulte, MF
    for c in range(n_cren):
        for s in ['RDC', 'Adulte', 'MF']:
            model.add_at_most_one(x[a, c, s] for a in agents)

    # D13 : 1 agent = 1 section par créneau
    for a in agents:
        for c in range(n_cren):
            model.add_at_most_one(x[a, c, s] for s in SECTIONS)

    # B1/B2 : disponibilité contractuelle
    for a in agents:
        cs_ce_list = creneaux_ouverts
        for c, (cs, ce) in enumerate(cs_ce_list):
            if not agent_disponible(a, jour, cs, ce, horaires_agents,
                                    evenements, date_str, pause_flex):
                for s in SECTIONS:
                    model.add(x[a, c, s] == 0)

    # B3 : vacataires uniquement les jours autorisés
    if jour not in mode_vac:
        for a in agents:
            if is_vacataire(a):
                for c in range(n_cren):
                    for s in SECTIONS:
                        model.add(x[a, c, s] == 0)

    # D1 : roulement samedi ROUGE/BLEU
    if jour == 'Samedi' and samedi_type:
        for a in agents:
            if is_vacataire(a):
                continue
            roul_agent = roulement_agents.get(a)
            if roul_agent and roul_agent != samedi_type:
                for c in range(n_cren):
                    for s in SECTIONS:
                        model.add(x[a, c, s] == 0)

    # C1/C2 : durées consécutives max
    max_consec = 4 * 60 if jour in ('Mercredi', 'Samedi') else 2 * 60 + 30
    for a in agents:
        for c_start in range(n_cren):
            # Construire la fenêtre de créneaux consécutifs à partir de c_start
            total_dur = 0
            c_end = c_start
            while c_end < n_cren:
                cs_e, ce_e = creneaux_ouverts[c_end]
                # Consécutif = le créneau suivant commence là où le précédent finit
                if c_end > c_start:
                    cs_prev, ce_prev = creneaux_ouverts[c_end - 1]
                    if cs_e != ce_prev:
                        break  # pas consécutif
                total_dur += ce_e - cs_e
                if total_dur > max_consec:
                    # Si tous les créneaux c_start..c_end sont actifs → violation
                    consec_vars = [x[a, c, s] for c in range(c_start, c_end + 1)
                                   for s in SECTIONS]
                    model.add(sum(consec_vars) <= (c_end - c_start))
                c_end += 1

    # C3 : pause déjeuner ≥ 1h (12h-14h) pour réguliers sauf Delphine
    agents_pause_oblig = [a for a in agents
                          if not is_vacataire(a) and a != 'Delphine']
    pause_creneaux = [c for c, (cs, ce) in enumerate(creneaux_ouverts)
                      if cs >= 720 and ce <= 840]  # 12h-14h = 720-840 min
    for a in agents_pause_oblig:
        if pause_creneaux:
            # Au moins 1 créneau de pause (= non travaillé) dans 12h-14h
            travail_pause = [x[a, c, s] for c in pause_creneaux for s in SECTIONS]
            # Durée couverte en pause doit être ≥ 60 min
            # Simplification : si ≥ 1h de créneaux dispo en pause, au moins 1 doit être libre
            pause_dur = sum(creneaux_ouverts[c][1] - creneaux_ouverts[c][0]
                            for c in pause_creneaux)
            if pause_dur >= 60:
                # L'agent ne peut pas travailler TOUS les créneaux de pause
                model.add(sum(travail_pause) < len(travail_pause))

    # F1/F2 : besoins Jeunesse exacts
    jour_key = jour
    if jour == 'Samedi':
        jour_key = f'Samedi_{samedi_type.lower()}' if samedi_type else 'Samedi_rouge'
    elif jour == 'Vendredi':
        jour_key = 'Vendredi'

    # Normaliser la période
    periode_key = None
    for k in besoins_jeunesse:
        if 'Vacances' in k and 'Hors' not in k and 'Vacances' in periode:
            if 'Hors' not in periode:
                periode_key = k
                break
        elif 'Hors' in k and 'Hors' in periode:
            periode_key = k
            break
    if not periode_key:
        periode_key = list(besoins_jeunesse.keys())[0] if besoins_jeunesse else None

    if periode_key:
        besoins_jour = besoins_jeunesse.get(periode_key, {}).get(jour_key, {})
        for c, (cs, ce) in enumerate(creneaux_ouverts):
            # Trouver le besoin pour ce créneau
            cren_str = f'{cs//60:02d}:{cs%60:02d}-{ce//60:02d}:{ce%60:02d}'
            besoin = besoins_jour.get(cren_str, 0)
            jeunesse_vars = [x[a, c, 'Jeunesse'] for a in agents]
            model.add(sum(jeunesse_vars) == besoin)

    # K3 (dure) : vacataire seul en Jeunesse uniquement 12h-14h
    for c, (cs, ce) in enumerate(creneaux_ouverts):
        is_in_12_14 = (cs >= 720 and ce <= 840)
        if not is_in_12_14:
            for a_vac in [a for a in agents if is_vacataire(a)]:
                # Si vacataire en Jeunesse → au moins 1 régulier aussi en Jeunesse
                reguliers_j = [x[a, c, 'Jeunesse'] for a in agents if not is_vacataire(a)]
                model.add(x[a_vac, c, 'Jeunesse'] <= sum(reguliers_j))

    # ══ CONTRAINTES MOLLES (pénalités) ════════════════════════

    penalites = []

    # G1 : respecter le planning type
    # Convertir planning_type_jour en créneaux indexés
    pt_indexed = {}
    for cren_str, sections_agents in planning_type_jour.items():
        parsed = parse_creneau(cren_str)
        if not parsed:
            continue

        # Trouver le(s) créneau(x) correspondant(s) dans creneaux_ouverts
        for c, (cs, ce) in enumerate(creneaux_ouverts):
            if cs >= parsed[0] and ce <= parsed[1]:
                pt_indexed.setdefault(c, {s: [] for s in SECTIONS})
                for s in SECTIONS:
                    pt_indexed[c][s] = sections_agents.get(s, [])

    # Pour chaque créneau PT, pénaliser si l'agent PT n'est pas à sa place
    for c, sections_dict in pt_indexed.items():
        for s, pt_agents in sections_dict.items():
            for a_pt in pt_agents:
                if a_pt in agents:
                    # Pénalité si l'agent PT n'est PAS dans sa section PT
                    not_in_pt = model.new_bool_var(f'not_in_pt_{a_pt}_{c}_{s}')
                    model.add(not_in_pt == 1 - x[a_pt, c, s])
                    penalites.append(POIDS['G1_planning_type'] * not_in_pt)
                else:
                    # Agent PT absent → pénalité si remplacement par agent
                    # de section différente (G2)
                    wrong_sect = []
                    for a in agents:
                        sect_prim = (affectations.get(a) or [''])[0]
                        if sect_prim != s:
                            wrong_sect.append(x[a, c, s])
                    if wrong_sect:
                        v = model.new_bool_var(f'wrong_sect_{c}_{s}')
                        model.add(sum(wrong_sect) >= 1).only_enforce_if(v)
                        model.add(sum(wrong_sect) == 0).only_enforce_if(v.negated())
                        penalites.append(POIDS['G2_meme_section_repl'] * v)

    # J1 : section principale prioritaire
    for a in agents:
        sects = affectations.get(a, [])
        if not sects:
            continue
        sect_prim = sects[0]
        cat = categories.get(a)
        sects_equiv = set(sects[:2]) if cat == 'A' else {sect_prim}
        for c in range(n_cren):
            for s in SECTIONS:
                if s not in sects_equiv and s in sects:
                    penalites.append(POIDS['J1_section_principale'] * x[a, c, s])

    # J3 : responsables déprioritisés
    for a in responsables:
        if a in agents:
            for c in range(n_cren):
                for s in SECTIONS:
                    penalites.append(POIDS['J3_responsable'] * x[a, c, s])

    # K1 : vacataires en dernier recours
    for a in agents:
        if is_vacataire(a):
            for c in range(n_cren):
                for s in SECTIONS:
                    penalites.append(POIDS['K1_vac_dernier_recours'] * x[a, c, s])

    # K2 : Vacataire 1 avant Vacataire 2
    vac1 = 'Vacataire 1'
    vac2 = 'Vacataire 2'
    if vac1 in agents and vac2 in agents:
        for c in range(n_cren):
            for s in SECTIONS:
                # Si vac2 travaille et vac1 ne travaille pas → pénalité
                v2_works = x[vac2, c, s]
                v1_not   = model.new_bool_var(f'v1_not_{c}_{s}')
                v1_total = sum(x[vac1, c2, s2] for c2 in range(n_cren) for s2 in SECTIONS)
                model.add(v1_not == 0).only_enforce_if(model.new_constant(1))
                penalites.append(POIDS['K2_vac1_avant_vac2'] * v2_works)

    # I1 : non-fragmentation (pénalité si agent travaille des créneaux non consécutifs)
    for a in agents:
        for c in range(n_cren - 1):
            cs_c,  ce_c  = creneaux_ouverts[c]
            cs_n,  ce_n  = creneaux_ouverts[c + 1]
            if ce_c != cs_n:  # pas consécutifs
                # Pénalité si l'agent travaille c mais pas c+1, et travaille c+2 ou plus
                for c2 in range(c + 2, n_cren):
                    travaille_c  = sum(x[a, c, s]  for s in SECTIONS)
                    travaille_c2 = sum(x[a, c2, s] for s in SECTIONS)
                    gap = model.new_bool_var(f'gap_{a}_{c}_{c2}')
                    model.add(travaille_c  >= 1).only_enforce_if(gap)
                    model.add(travaille_c2 >= 1).only_enforce_if(gap)
                    penalites.append(POIDS['I1_non_fragmentation'] * gap)

    # Objectif : minimiser les pénalités
    model.minimize(sum(penalites))

    # ══ RÉSOLUTION ════════════════════════════════════════════
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    solver.parameters.num_search_workers  = 4
    status = solver.solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # ══ EXTRACTION DE LA SOLUTION ═════════════════════════════
    result = {}
    for c in range(n_cren):
        result[c] = {s: [] for s in SECTIONS}
        for a in agents:
            for s in SECTIONS:
                if solver.value(x[a, c, s]) == 1:
                    result[c][s].append(a)

    return result


# ══════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE PRINCIPAL
# ══════════════════════════════════════════════════════════════

def compute_full_planning(filepath):
    """
    Calcule le planning complet du mois.
    Retourne (weeks_data, metadata) au même format que l'ancien moteur.
    """
    raw = load_excel_data(filepath)

    params         = parse_parametres(raw)
    affectations, categories, responsables, pause_flex = parse_affectations(raw)
    horaires_agents = parse_horaires_agents(raw)
    roulement_type, roulement_exceptions = parse_roulement_samedi(raw)
    besoins_jeunesse = parse_besoins_jeunesse(raw)
    evenements       = parse_evenements(raw)
    planning_type    = parse_planning_type(raw)

    calendrier = build_calendar(params['mois'], params['annee'], params['samedis'])

    agents_tous = list(affectations.keys())

    weeks_data = []
    for semaine in calendrier:
        week_num  = semaine['num']
        periode   = params['semaines'].get(week_num, 'Hors Vacances scolaires')
        week_plan = {'week_num': week_num, 'jours': []}

        for jour_info in semaine['jours']:
            date_str   = jour_info['date']
            jour       = jour_info['jour']
            sam_type   = jour_info.get('samedi_type')

            # Roulement samedi (avec exceptions)
            roulement_agents = dict(roulement_type)
            for agent_exc, roul_exc in roulement_exceptions.get(week_num, {}).items():
                roulement_agents[agent_exc] = roul_exc.upper()

            # Agents éligibles ce jour
            agents_eligibles = []
            for a in agents_tous:
                if is_vacataire(a):
                    if jour in params['mode_vac']:
                        agents_eligibles.append(a)
                else:
                    h = horaires_agents.get(a, {}).get(jour)
                    if h and any(v is not None for v in h):
                        agents_eligibles.append(a)

            # Planning type pour ce jour
            if jour == 'Samedi' and sam_type:
                pt_jour_key = f'Samedi_{sam_type}'
            else:
                pt_jour_key = jour
            pt_jour = planning_type.get(pt_jour_key, {})

            # Créneaux ouverts
            creneaux_ouverts = params['creneaux']

            # Résolution CP-SAT
            solution = solve_day(
                jour=jour,
                date_str=date_str,
                creneaux_ouverts=creneaux_ouverts,
                agents_eligibles=agents_eligibles,
                affectations=affectations,
                categories=categories,
                responsables=responsables,
                pause_flex=pause_flex,
                horaires_agents=horaires_agents,
                evenements=evenements,
                besoins_jeunesse=besoins_jeunesse,
                planning_type_jour=pt_jour,
                roulement_agents=roulement_agents,
                samedi_type=sam_type,
                periode=periode,
                mode_vac=params['mode_vac'],
            )

            week_plan['jours'].append({
                'date':      date_str,
                'jour':      jour,
                'sam_type':  sam_type,
                'creneaux':  creneaux_ouverts,
                'solution':  solution,   # {cren_idx: {section: [agents]}}
                'infaisable': solution is None,
            })

        weeks_data.append(week_plan)

    metadata = {
        'mois':       params['mois'],
        'annee':      params['annee'],
        'evenements': evenements,
    }

    return weeks_data, metadata
