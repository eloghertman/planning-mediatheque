"""
planning_checker.py — Vérification d'un planning déjà rempli / modifié à la main.

Contrairement à planning_engine_cpsat.py (qui CALCULE un planning), ce module
RELIT un planning déjà généré et modifié, et signale les contradictions avec
les contraintes dures. Un seul fichier en entrée : le classeur Excel du
planning (même structure que la sortie de generate_planning_excel_septembre.py).

Fonction principale : verifier_planning(file_bytes) -> list[Anomalie]
"""

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from io import BytesIO

import openpyxl


# ─────────────────────────────────────────────────────────────
#  RÉFÉRENTIEL MÉTIER (voir context_projet_mediatheque_v30.md §3, §7)
# ─────────────────────────────────────────────────────────────

HABILITATIONS = {
    'Marie-France':    ['RDC', 'Adulte', 'M & F'],
    'Anne-Françoise':  ['Adulte', 'Jeunesse', 'M & F', 'RDC'],
    'Christine':       ['Adulte', 'RDC'],
    'Léa':             ['Adulte', 'M & F', 'RDC', 'Jeunesse'],
    'Chloé':           ['Adulte', 'RDC', 'Jeunesse'],
    'Macha':           ['RDC', 'Adulte'],
    'Delphine':        ['M & F', 'RDC', 'Jeunesse', 'Adulte'],
    'Barbara':         ['M & F', 'Jeunesse'],
    'Stéphane':        ['M & F'],
    'Stéphanie':       ['Jeunesse', 'RDC'],
    'Robin':           ['Jeunesse', 'RDC'],
    'Guillaume':       ['Jeunesse', 'RDC'],
    'Agnès':           ['Jeunesse'],
    'Tiphaine':        ['Jeunesse', 'RDC', 'M & F', 'Adulte'],
}
ALL_AGENTS_CONNUS = list(HABILITATIONS.keys()) + ['Vacataire 1', 'Vacataire 2', 'Vacataire 3', 'Vacataire']

PAUSE_EXEMPTS = {'Delphine'}          # jamais de contrôle de pause déjeuner (§7 C3)
AGENTS_IGNORES = {'lydie'}            # a quitté l'équipe, ignorée partout

JOURS_ORDRE = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI', 'DIMANCHE']

PAUSE_FENETRE = (12 * 60, 14 * 60)    # 12h-14h, en minutes
PAUSE_MIN_LIBRE = 60                  # minutes


# ─────────────────────────────────────────────────────────────
#  UTILITAIRES TEXTE / TEMPS
# ─────────────────────────────────────────────────────────────

def normalize(s):
    if not s:
        return ''
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s


def est_vacataire(nom):
    return 'vacataire' in normalize(nom)


def est_eloise(nom):
    return normalize(nom) in ('eloise', 'eloïse'.replace('ï', 'i'), normalize('Eloïse'))


def est_ignore(nom):
    return normalize(nom) in AGENTS_IGNORES


def fmt_min(m):
    if m is None:
        return '?'
    h, mn = divmod(int(m), 60)
    return f"{h}h{mn:02d}" if mn else f"{h}h"


def is_creneau(val):
    return isinstance(val, str) and re.match(r'^\d{1,2}:\d{2}-\d{1,2}:\d{2}$', val.strip()) is not None


def parse_creneau(val):
    a, b = val.strip().split('-')
    h1, m1 = a.split(':')
    h2, m2 = b.split(':')
    return int(h1) * 60 + int(m1), int(h2) * 60 + int(m2)


def parse_heure_texte(txt):
    """'8h30' / '17h' -> minutes. None si non trouvé."""
    m = re.search(r'(\d{1,2})h(\d{2})?', txt)
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2) or 0)


def parenthese_finale(texte):
    m = re.search(r'\(([^)]*)\)\s*$', texte or '')
    return m.group(1) if m else None


def extraire_agents_et_fenetre(texte, defaut_debut, defaut_fin, agents_connus):
    """À partir d'un texte du type 'Accueil de classe (Stéphanie, 10h15-10h45)'
    ou 'congé (Marie-France, Christine)', retourne (liste_agents, debut, fin)."""
    agents_trouves = []
    debut, fin = defaut_debut, defaut_fin
    inner = parenthese_finale(texte)
    if inner:
        for part in inner.split(','):
            part = part.strip()
            m = re.match(r'^(\d{1,2})h(\d{2})?-(\d{1,2})h(\d{2})?$', part)
            if m:
                h1, m1, h2, m2 = m.groups()
                debut = int(h1) * 60 + int(m1 or 0)
                fin = int(h2) * 60 + int(m2 or 0)
            else:
                for agent in agents_connus:
                    if normalize(agent) == normalize(part):
                        agents_trouves.append(agent)
                        break
    # Cas particulier "congé (Nom1, Nom2)" : le nom du champ précède la parenthèse.
    if texte and normalize(texte).startswith('conge') and not agents_trouves and inner:
        for part in inner.split(','):
            part = part.strip()
            for agent in agents_connus:
                if normalize(agent) == normalize(part):
                    agents_trouves.append(agent)
                    break
    return agents_trouves, debut, fin


# ─────────────────────────────────────────────────────────────
#  ANOMALIE
# ─────────────────────────────────────────────────────────────

@dataclass
class Anomalie:
    gravite: str          # 'rouge' (impossibilité) ou 'jaune' (suspect)
    semaine: str
    jour: str
    message: str
    regle: str = ''


# ─────────────────────────────────────────────────────────────
#  LECTURE — cellules / cellules fusionnées
# ─────────────────────────────────────────────────────────────

def build_merge_map(ws):
    merge_map = {}
    for mc in ws.merged_cells.ranges:
        top_val = ws.cell(row=mc.min_row, column=mc.min_col).value
        for row in range(mc.min_row, mc.max_row + 1):
            for col in range(mc.min_col, mc.max_col + 1):
                merge_map[(row, col)] = top_val
    return merge_map


def get_cell(ws, row, col, merge_map):
    if (row, col) in merge_map:
        return merge_map[(row, col)]
    return ws.cell(row=row, column=col).value


# ─────────────────────────────────────────────────────────────
#  LECTURE — feuille "Semaine_N" (grille principale)
# ─────────────────────────────────────────────────────────────

def lire_jours_semaine(ws):
    """Découpe la feuille en blocs 'journée' à partir des titres en colonne A
    ('  MARDI  1 Septembre 2026', éventuellement suivi de '— SAMEDI BLEU')."""
    merge_map = build_merge_map(ws)
    jours = []
    max_row = ws.max_row
    r = 1
    while r <= max_row:
        val = ws.cell(row=r, column=1).value
        if isinstance(val, str):
            v = val.strip().upper()
            jour_trouve = next((j for j in JOURS_ORDRE if v.startswith(j)), None)
            if jour_trouve:
                samedi_type = None
                if 'BLEU' in v:
                    samedi_type = 'BLEU'
                elif 'ROUGE' in v:
                    samedi_type = 'ROUGE'
                # Ligne d'en-tête juste après (contient 'Créneau'), données ensuite.
                r_data = r + 2
                creneaux = []
                while r_data <= max_row and is_creneau(ws.cell(row=r_data, column=1).value):
                    debut, fin = parse_creneau(ws.cell(row=r_data, column=1).value)
                    rdc = get_cell(ws, r_data, 2, merge_map)
                    adulte = get_cell(ws, r_data, 3, merge_map)
                    mf = get_cell(ws, r_data, 4, merge_map)
                    jeun = [get_cell(ws, r_data, c, merge_map) for c in (5, 6, 7)]
                    accueil = get_cell(ws, r_data, 8, merge_map)
                    reunion = get_cell(ws, r_data, 9, merge_map)
                    absence = get_cell(ws, r_data, 10, merge_map)

                    def clean(v):
                        if v in (None, '—', ''):
                            return None
                        return v

                    creneaux.append({
                        'row': r_data, 'debut': debut, 'fin': fin,
                        'rdc': clean(rdc), 'adulte': clean(adulte), 'mf': clean(mf),
                        'jeunesse': [clean(x) for x in jeun],
                        'accueil': clean(accueil), 'reunion': clean(reunion),
                        'absence': clean(absence),
                    })
                    r_data += 1
                jours.append({
                    'jour': jour_trouve, 'titre': val.strip(),
                    'samedi_type': samedi_type,
                    'row_titre': r, 'row_debut_data': r + 2, 'row_fin_data': r_data - 1,
                    'creneaux': creneaux,
                })
                r = r_data
                continue
        r += 1
    return jours


def lire_notes_agents_jour(ws, row_debut, row_fin):
    """Lit les colonnes W/X (Nom/Événement) et Y/Z (Nom/Événement) pour les
    lignes d'un jour donné. Retourne une liste de (agent, texte_note)."""
    notes = []
    for r in range(row_debut, row_fin + 1):
        for col_nom, col_evt in ((23, 24), (25, 26)):  # W/X, Y/Z
            nom = ws.cell(row=r, column=col_nom).value
            evt = ws.cell(row=r, column=col_evt).value
            if nom and evt and str(evt).strip():
                notes.append((str(nom).strip(), str(evt).strip()))
    return notes


# ─────────────────────────────────────────────────────────────
#  LECTURE — feuille "Semaine_N_Agent" (vue par agent = horaires réels)
# ─────────────────────────────────────────────────────────────

def lire_vue_agent(ws):
    """Retourne dict agent -> jour -> {'arrivee':min|None,'depart':min|None,'conge':bool}."""
    result = {}
    max_row = ws.max_row
    max_col = ws.max_column
    header = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]
    jours_cols = {}
    for c, h in enumerate(header, start=1):
        if isinstance(h, str):
            v = h.strip().upper()
            jour_trouve = next((j for j in JOURS_ORDRE if v.startswith(j)), None)
            if jour_trouve:
                jours_cols[jour_trouve] = c

    current_agent = None
    r = 2
    while r <= max_row:
        colA = ws.cell(row=r, column=1).value
        if colA is None:
            current_agent = None
            r += 1
            continue
        if not is_creneau(colA):
            current_agent = str(colA).strip()
            result.setdefault(current_agent, {j: {'arrivee': None, 'depart': None, 'conge': False}
                                                for j in jours_cols})
            r += 1
            continue
        if current_agent:
            for jour, col in jours_cols.items():
                val = ws.cell(row=r, column=col).value
                if isinstance(val, str):
                    v = val.strip()
                    vlow = normalize(v)
                    if vlow.startswith('arriv'):
                        h = parse_heure_texte(v)
                        if h is not None:
                            # Certain·es agent·es (pause flexible) ont 2 lignes
                            # "Arrivée" le même jour (avant/après une pause
                            # variable) : on ne garde ici que la BORNE LA PLUS
                            # TÔT, pour ne pas déclencher de fausse alerte sur
                            # le retour de pause. La pause elle-même reste
                            # couverte par la règle "pause déjeuner" (§ R4).
                            prev = result[current_agent][jour]['arrivee']
                            result[current_agent][jour]['arrivee'] = h if prev is None else min(prev, h)
                    elif vlow.startswith('conge'):
                        result[current_agent][jour]['conge'] = True
                    elif vlow.startswith('depar'):
                        h = parse_heure_texte(v)
                        if h is not None:
                            prev = result[current_agent][jour]['depart']
                            result[current_agent][jour]['depart'] = h if prev is None else max(prev, h)
        r += 1
    return result


# ─────────────────────────────────────────────────────────────
#  CONSTRUCTION DES OCCURRENCES PAR AGENT / JOUR
# ─────────────────────────────────────────────────────────────

def construire_occurrences_jour(jour_data, agents_connus):
    occ = defaultdict(list)
    for cren in jour_data['creneaux']:
        cs, ce = cren['debut'], cren['fin']
        for label, val in (('RDC', cren['rdc']), ('Adulte', cren['adulte']), ('M & F', cren['mf'])):
            if val:
                occ[val].append({'debut': cs, 'fin': ce, 'type': label, 'detail': label})
        for val in cren['jeunesse']:
            if val:
                occ[val].append({'debut': cs, 'fin': ce, 'type': 'Jeunesse', 'detail': 'Jeunesse'})
        if cren['accueil']:
            agents, d, f = extraire_agents_et_fenetre(cren['accueil'], cs, ce, agents_connus)
            for a in agents:
                occ[a].append({'debut': d, 'fin': f, 'type': 'Accueil/Animation', 'detail': cren['accueil']})
        if cren['reunion']:
            agents, d, f = extraire_agents_et_fenetre(cren['reunion'], cs, ce, agents_connus)
            for a in agents:
                occ[a].append({'debut': d, 'fin': f, 'type': 'Réunion', 'detail': cren['reunion']})
        if cren['absence']:
            agents, d, f = extraire_agents_et_fenetre(cren['absence'], cs, ce, agents_connus)
            for a in agents:
                occ[a].append({'debut': d, 'fin': f, 'type': 'Absence', 'detail': cren['absence']})
    return occ


def fusionner_occurrences(liste):
    groups = defaultdict(list)
    for o in liste:
        groups[(o['type'], o['detail'])].append(o)
    fusion = []
    for items in groups.values():
        items.sort(key=lambda x: x['debut'])
        cur = None
        for it in items:
            if cur is None:
                cur = dict(it)
            elif it['debut'] <= cur['fin']:
                cur['fin'] = max(cur['fin'], it['fin'])
            else:
                fusion.append(cur)
                cur = dict(it)
        if cur:
            fusion.append(cur)
    return fusion


# ─────────────────────────────────────────────────────────────
#  RÈGLES DE VÉRIFICATION
# ─────────────────────────────────────────────────────────────

def verifier_jour(jour_data, semaine_label, vue_agent, agents_connus, anomalies):
    jour = jour_data['jour']
    occ_brutes = construire_occurrences_jour(jour_data, agents_connus)

    # bornes d'ouverture approximatives ce jour = 1er début / dernière fin des créneaux
    if jour_data['creneaux']:
        ouverture_debut = jour_data['creneaux'][0]['debut']
        ouverture_fin = jour_data['creneaux'][-1]['fin']
    else:
        ouverture_debut = ouverture_fin = None

    for agent, liste in occ_brutes.items():
        if est_ignore(agent):
            continue
        occs = fusionner_occurrences(liste)
        occs_travail = [o for o in occs if o['type'] != 'Absence']

        # R8 — Eloïse ne doit jamais apparaître
        if est_eloise(agent):
            for o in occs:
                anomalies.append(Anomalie(
                    'rouge', semaine_label, jour,
                    f"Eloïse apparaît dans le planning ({o['type']}, {fmt_min(o['debut'])}-{fmt_min(o['fin'])}) "
                    f"— elle ne doit jamais être affectée.",
                    'Eloïse jamais planifiée'))
            continue

        # R1 — horaires contractuels (vue par agent)
        info_h = vue_agent.get(agent, {}).get(jour, {})
        arrivee, depart = info_h.get('arrivee'), info_h.get('depart')
        for o in occs_travail:
            if arrivee is not None and o['debut'] < arrivee:
                anomalies.append(Anomalie(
                    'rouge', semaine_label, jour,
                    f"{agent} est indiqué·e en {o['type']} dès {fmt_min(o['debut'])}, "
                    f"mais son horaire indique une arrivée à {fmt_min(arrivee)}.",
                    'Horaires contractuels'))
            if depart is not None and o['fin'] > depart:
                anomalies.append(Anomalie(
                    'rouge', semaine_label, jour,
                    f"{agent} est indiqué·e en {o['type']} jusqu'à {fmt_min(o['fin'])}, "
                    f"mais son horaire indique un départ à {fmt_min(depart)}.",
                    'Horaires contractuels'))

        # R2/R3 — chevauchements (y compris congé/absence vs travail)
        for i in range(len(occs)):
            for j in range(i + 1, len(occs)):
                a, b = occs[i], occs[j]
                if a['debut'] < b['fin'] and b['debut'] < a['fin']:
                    if a['type'] == 'Absence' or b['type'] == 'Absence':
                        autre = b if a['type'] == 'Absence' else a
                        anomalies.append(Anomalie(
                            'rouge', semaine_label, jour,
                            f"{agent} est en congé/absence mais apparaît aussi en {autre['type']} "
                            f"({autre['detail']}) de {fmt_min(autre['debut'])} à {fmt_min(autre['fin'])}.",
                            'Congé = jamais planifié'))
                    else:
                        anomalies.append(Anomalie(
                            'rouge', semaine_label, jour,
                            f"{agent} est indiqué·e en {a['type']} ({a['detail']}) de "
                            f"{fmt_min(a['debut'])} à {fmt_min(a['fin'])} ET en {b['type']} ({b['detail']}) "
                            f"de {fmt_min(b['debut'])} à {fmt_min(b['fin'])} — ces deux horaires se chevauchent.",
                            'Un agent à un seul endroit à la fois'))

        # R5 — habilitations
        if not est_vacataire(agent) and agent in HABILITATIONS:
            for o in occs_travail:
                if o['type'] in ('RDC', 'Adulte', 'M & F', 'Jeunesse') and o['type'] not in HABILITATIONS[agent]:
                    anomalies.append(Anomalie(
                        'rouge', semaine_label, jour,
                        f"{agent} est affecté·e en {o['type']} de {fmt_min(o['debut'])} à {fmt_min(o['fin'])}, "
                        f"section non habilitée (habilitations : {', '.join(HABILITATIONS[agent])}).",
                        'Habilitations par section'))
        elif not est_vacataire(agent) and agent not in HABILITATIONS:
            anomalies.append(Anomalie(
                'jaune', semaine_label, jour,
                f"'{agent}' n'est pas reconnu·e dans la liste habituelle des agents — vérifier l'orthographe "
                f"ou une éventuelle nouvelle recrue non encore répertoriée.",
                'Agent inconnu'))

        # R6 — vacataires jamais au RDC
        if est_vacataire(agent):
            for o in occs_travail:
                if o['type'] == 'RDC':
                    anomalies.append(Anomalie(
                        'rouge', semaine_label, jour,
                        f"{agent} (vacataire) est affecté·e au RDC de {fmt_min(o['debut'])} à {fmt_min(o['fin'])} "
                        f"— un vacataire ne doit jamais être au RDC.",
                        'Vacataires jamais au RDC'))

        # R4 — pause déjeuner (suspect, pas certain sans le fichier de préparation)
        if agent not in PAUSE_EXEMPTS and not est_vacataire(agent):
            pres_debut = arrivee if arrivee is not None else ouverture_debut
            pres_fin = depart if depart is not None else ouverture_fin
            if pres_debut is not None and pres_fin is not None:
                fen_debut = max(pres_debut, PAUSE_FENETRE[0])
                fen_fin = min(pres_fin, PAUSE_FENETRE[1])
                if fen_fin - fen_debut >= PAUSE_MIN_LIBRE:
                    segs = sorted(
                        [(max(o['debut'], fen_debut), min(o['fin'], fen_fin))
                         for o in occs_travail if o['debut'] < fen_fin and o['fin'] > fen_debut],
                        key=lambda x: x[0])
                    libre_max = 0
                    curseur = fen_debut
                    for d, f in segs:
                        if d > curseur:
                            libre_max = max(libre_max, d - curseur)
                        curseur = max(curseur, f)
                    libre_max = max(libre_max, fen_fin - curseur)
                    if libre_max < PAUSE_MIN_LIBRE:
                        anomalies.append(Anomalie(
                            'jaune', semaine_label, jour,
                            f"{agent} ne semble pas avoir au moins 1h vraiment libre entre 12h et 14h "
                            f"(sur la plage {fmt_min(fen_debut)}-{fmt_min(fen_fin)} où il/elle est présent·e). "
                            f"À vérifier — peut être normal si son contrat prévoit une présence continue ce jour-là.",
                            'Pause déjeuner'))

    # R7 — vacataire seul en Jeunesse hors 12h-14h
    for cren in jour_data['creneaux']:
        jeunesse_agents = [a for a in cren['jeunesse'] if a and not est_ignore(a)]
        if jeunesse_agents and all(est_vacataire(a) for a in jeunesse_agents):
            if not (cren['debut'] >= PAUSE_FENETRE[0] and cren['fin'] <= PAUSE_FENETRE[1]):
                anomalies.append(Anomalie(
                    'rouge', semaine_label, jour,
                    f"Jeunesse {fmt_min(cren['debut'])}-{fmt_min(cren['fin'])} : uniquement des vacataires "
                    f"({', '.join(jeunesse_agents)}) — autorisé seulement sur 12h-14h.",
                    'Vacataire seul en Jeunesse'))

    # R10 — garde-fou : rien trouvé nulle part ce jour
    notes = lire_notes_agents_jour  # placeholder, complété dans verifier_planning


# ─────────────────────────────────────────────────────────────
#  FONCTION PRINCIPALE
# ─────────────────────────────────────────────────────────────

def verifier_planning(file_bytes):
    """file_bytes : bytes du classeur Excel du planning déjà rempli.
    Retourne une liste d'Anomalie."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    anomalies = []

    semaine_sheets = sorted(
        [n for n in wb.sheetnames if re.match(r'^Semaine_\d+$', n)],
        key=lambda n: int(re.search(r'\d+', n).group())
    )

    for sn in semaine_sheets:
        ws = wb[sn]
        agent_sheet_name = f"{sn}_Agent"
        vue_agent = {}
        if agent_sheet_name in wb.sheetnames:
            vue_agent = lire_vue_agent(wb[agent_sheet_name])
        else:
            anomalies.append(Anomalie(
                'jaune', sn, '',
                f"L'onglet '{agent_sheet_name}' est introuvable : les horaires contractuels "
                f"(règle 'arrivée/départ') n'ont pas pu être vérifiés pour cette semaine.",
                'Fichier incomplet'))

        jours = lire_jours_semaine(ws)
        for jour_data in jours:
            verifier_jour(jour_data, sn, vue_agent, ALL_AGENTS_CONNUS, anomalies)

            # Garde-fou : rien trouvé (ni H/I/J, ni notes W-Z) ce jour-là
            rien_dans_grille = all(
                not c['accueil'] and not c['reunion'] and not c['absence']
                for c in jour_data['creneaux']
            )
            notes = lire_notes_agents_jour(ws, jour_data['row_debut_data'], jour_data['row_fin_data'])
            if rien_dans_grille and not notes:
                anomalies.append(Anomalie(
                    'jaune', sn, jour_data['jour'],
                    f"Aucun événement noté ce jour (ni dans les colonnes Accueil/Animation/Réunion/Absence, "
                    f"ni dans les notes agents W-Z) — à vérifier si c'est normal.",
                    'Garde-fou : rien trouvé'))

            # Cohérence notes agents (W-Z) <-> ce qui apparaît dans H/I/J
            for agent, texte in notes:
                if est_ignore(agent):
                    continue
                fragment = texte.split(' ', 1)[-1] if re.match(r'^\d{1,2}h', texte) else texte
                fragment_norm = normalize(fragment)[:20]
                trouve = False
                for c in jour_data['creneaux']:
                    for champ in (c['accueil'], c['reunion'], c['absence']):
                        if champ and fragment_norm and fragment_norm in normalize(champ):
                            trouve = True
                if not trouve and fragment_norm:
                    anomalies.append(Anomalie(
                        'jaune', sn, jour_data['jour'],
                        f"Note ajoutée par {agent} (« {texte} ») ne semble pas se retrouver dans le planning "
                        f"(colonnes Accueil/Animation, Réunion ou Absence) — à vérifier manuellement.",
                        'Note non répercutée'))

    return anomalies


# ─────────────────────────────────────────────────────────────
#  AFFICHAGE (utilisable directement en dehors de Streamlit)
# ─────────────────────────────────────────────────────────────

def resumer(anomalies):
    n_rouge = sum(1 for a in anomalies if a.gravite == 'rouge')
    n_jaune = sum(1 for a in anomalies if a.gravite == 'jaune')
    return n_rouge, n_jaune


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Planning_Semaine1_avec_notes_agents.xlsx'
    with open(path, 'rb') as f:
        data = f.read()
    anomalies = verifier_planning(data)
    n_rouge, n_jaune = resumer(anomalies)
    print(f"{len(anomalies)} anomalies détectées ({n_rouge} rouge, {n_jaune} jaune)\n")
    for a in anomalies:
        marqueur = '🔴' if a.gravite == 'rouge' else '🟡'
        print(f"{marqueur} [{a.semaine} — {a.jour}] {a.message}")
