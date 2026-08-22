"""
regeneration_calcul.py — Brique 2 de la "régénération partielle".

Prend la sortie de la brique 1 (regeneration_lecture.lire_planning_pour_regeneration)
et relance le VRAI moteur CP-SAT (solve_day, celui de planning_engine_cpsat.py —
aucune logique de calcul dupliquée) uniquement sur les jours à régénérer,
dans l'ordre chronologique de la semaine, en enchaînant le compteur d'équité
hebdomadaire d'un jour à l'autre exactement comme le fait compute_full_planning
pour une génération complète.

⚠️ Deux petites fonctions (construire_grille_vacances_jour /
resoudre_besoins_jour) sont dupliquées ici depuis planning_engine_cpsat.py,
car elles y sont définies EN INTERNE (fonctions imbriquées dans
compute_full_planning) et ne peuvent donc pas être importées telles quelles.
Pas grave à ce stade, mais à garder en tête : si leur logique change un jour
dans le moteur principal, il faudra penser à répercuter le changement ici
aussi (idéalement, on sortira ces deux fonctions du moteur pour les rendre
réellement partagées — amélioration possible plus tard, pas urgente).

Ne touche à AUCUN fichier Excel — produit juste un résultat en mémoire, au
même format que celui que `compute_full_planning` donne pour un jour
('creneau_idx' -> {section: [agents]}), pour que la brique 3 (écriture +
alertes visuelles) puisse s'en servir exactement comme pour une génération
normale.
"""

from planning_engine_cpsat import (
    solve_day, parse_creneau, is_vacataire,
)
from planning_checker import JOUR_CAPITALISE, JOURS_ORDRE


class ErreurCalcul(Exception):
    """Erreur bloquante avant même de lancer le calcul (ex. conflit non
    résolu dans la zone à régénérer)."""
    pass


# ─────────────────────────────────────────────────────────────
#  DUPLIQUÉ DEPUIS planning_engine_cpsat.compute_full_planning
#  (voir avertissement en tête de fichier)
# ─────────────────────────────────────────────────────────────

def _resoudre_besoins_jour(besoins_jeunesse, jour_x, samedi_type_x):
    periode_key = next((p for p in besoins_jeunesse if 'Hors' not in p), None)
    if not periode_key:
        return {}
    jours_dict = besoins_jeunesse.get(periode_key, {})
    if jour_x == 'Samedi' and samedi_type_x:
        def _norm(s):
            return s.lower().replace('_', ' ').replace('-', ' ').strip()
        cible = _norm(f'samedi {samedi_type_x}')
        jour_key = next((k for k in jours_dict if _norm(k) == cible),
                         f'Samedi_{samedi_type_x.lower()}')
    else:
        jour_key = jour_x
    return jours_dict.get(jour_key, {})


def _construire_grille_vacances_jour(besoins_jeunesse, params, jour_x, samedi_type_x):
    besoins_jour = _resoudre_besoins_jour(besoins_jeunesse, jour_x, samedi_type_x)
    ranges = []
    for cren_str, besoin in besoins_jour.items():
        parsed = parse_creneau(cren_str)
        if parsed:
            ranges.append((parsed[0], parsed[1], besoin))
    ranges.sort()
    if not ranges:
        return []
    blocs_standards = (params['creneaux_ms'] if jour_x in ('Mercredi', 'Samedi')
                        else params['creneaux_mjv'])
    merged = []
    for bs, be in blocs_standards:
        sous = [(cs, ce, b) for (cs, ce, b) in ranges if cs >= bs and ce <= be]
        if not sous:
            merged.append((bs, be))
            continue
        cur = list(sous[0])
        for cs, ce, b in sous[1:]:
            if cs == cur[1] and b == cur[2]:
                cur[1] = ce
            else:
                merged.append((cur[0], cur[1]))
                cur = [cs, ce, b]
        merged.append((cur[0], cur[1]))
    return merged


def _construire_swap_map(jour_cap, sam_type, semaine_num, roulement_type, roulement_exceptions):
    """Reprend telle quelle la logique de compute_full_planning pour bâtir
    le swap_map (remplacements liés aux exceptions de roulement samedi)."""
    swap_map = {}
    if jour_cap == 'Samedi' and sam_type:
        exc = roulement_exceptions.get(semaine_num, {})
        vers_autre = {a: r for a, r in exc.items() if r != sam_type}
        vers_ce_sam = {a: r for a, r in exc.items() if r == sam_type}
        for a_absent, r_absent in vers_autre.items():
            for a_repl, r_repl in vers_ce_sam.items():
                if a_absent not in swap_map:
                    swap_map[a_absent] = a_repl
    return swap_map


# ─────────────────────────────────────────────────────────────
#  BRIQUE 2 — POINT D'ENTRÉE
# ─────────────────────────────────────────────────────────────

def regenerer_jours(lecture_resultat, cumul_hebdo_initial=None):
    """
    - lecture_resultat : sortie de lire_planning_pour_regeneration() (brique 1).
    - cumul_hebdo_initial : {agent: minutes} — compteur d'équité hebdomadaire
      à utiliser comme point de départ. Laisser à None pour une régénération
      de semaine complète (équivaut à démarrer à zéro, comme une génération
      normale). ⚠️ Pour une régénération partielle (certains jours de la
      semaine restent fixes), il n'y a pas aujourd'hui de moyen fiable de
      reconstituer ce compteur à partir du fichier Excel déjà rempli (cette
      information n'y est pas conservée) — limite déjà signalée.

    Ne bloque PAS sur les conflits détectés dans la zone à régénérer (choix
    explicite du 20/08) : le calcul tourne quand même, les données déjà
    notées (Accueil/Réunion/Absence) restent des contraintes figées telles
    quelles, même si deux d'entre elles se contredisent entre elles — dans
    ce cas l'agent est simplement considéré indisponible sur l'ensemble des
    deux plages, ce qui ne bloque jamais techniquement le calcul. Les
    conflits (ConflitFixe, cf. brique 1) restent transmis dans le résultat
    pour que la brique 3 les affiche en alerte visuelle (encadré rouge) sur
    le planning généré.

    Retourne :
    {
        'semaine_num': int,
        'conflits_a_signaler': [ConflitFixe],  # transmis tels quels depuis la brique 1
        'jours': [
            {'date', 'jour' (MAJUSCULE, ex 'MERCREDI'), 'jour_cap' (ex 'Mercredi'),
             'sam_type', 'creneaux', 'solution', 'infaisable', 'alertes',
             'cumul_hebdo_apres'},
            ...  # un par jour régénéré, DANS L'ORDRE CHRONOLOGIQUE
        ],
    }
    """
    prep = lecture_resultat['prep']
    params = prep.get('params', {})
    semaine_num = lecture_resultat['semaine_num']
    jours_dispo = lecture_resultat['jours_data_bruts']
    evenements = lecture_resultat['evenements_regeneres']

    affectations = prep.get('affectations', {})
    categories = prep.get('categories', {})
    responsables = prep.get('responsables', {})
    pause_flex = prep.get('pause_flex', set())
    priorite_rdc = prep.get('priorite_rdc', {})
    horaires_agents = prep.get('horaires_agents', {})
    besoins_jeunesse = prep.get('besoins_jeunesse', {})
    planning_type = prep.get('planning_type', {})
    roulement_type = prep.get('roulement_type', {})
    roulement_exceptions = prep.get('roulement_exceptions', {})
    jours_speciaux = prep.get('jours_speciaux', {})
    presences_vac = params.get('presences_vac', {})
    mode_vac = params.get('mode_vac', set())

    periode_semaine = params.get('semaines', {}).get(semaine_num, 'Hors Vacances scolaires')

    agents_tous = list(affectations.keys())

    # Jours à régénérer, dans l'ORDRE CHRONOLOGIQUE de la semaine (pas
    # l'ordre dans lequel l'utilisatrice les a tapés) — indispensable pour
    # enchaîner le compteur d'équité correctement.
    jours_ordonnes = sorted(
        lecture_resultat['jours_regeneres'],
        key=lambda j: JOURS_ORDRE.index(j) if j in JOURS_ORDRE else 99
    )

    cumul_hebdo = dict(cumul_hebdo_initial or {})
    resultats_jours = []

    for jour_maj in jours_ordonnes:
        jour_data = jours_dispo[jour_maj]
        jour_cap = JOUR_CAPITALISE.get(jour_maj, jour_maj.capitalize())
        date_str = jour_data.get('date_str')
        sam_type = jour_data.get('samedi_type')

        periode_effective = periode_semaine
        js_info = jours_speciaux.get(date_str)
        if js_info and js_info.get('vacances'):
            periode_effective = 'Vacances Scolaires'

        # Agents éligibles ce jour (même logique que compute_full_planning)
        agents_eligibles = []
        use_presences = bool(presences_vac)
        for a in agents_tous:
            if is_vacataire(a):
                if use_presences:
                    if date_str in presences_vac and a in presences_vac[date_str]:
                        agents_eligibles.append(a)
                elif jour_cap in mode_vac:
                    agents_eligibles.append(a)
            else:
                h = horaires_agents.get(a, {}).get(jour_cap)
                if h and any(v is not None for v in h):
                    agents_eligibles.append(a)

        # Planning type de ce jour
        pt_jour_key = f'Samedi_{sam_type}' if (jour_cap == 'Samedi' and sam_type) else jour_cap
        pt_jour = planning_type.get(pt_jour_key, {})

        # Créneaux ouverts
        creneaux_vac = (_construire_grille_vacances_jour(besoins_jeunesse, params, jour_cap, sam_type)
                         if 'Hors' not in periode_effective else [])
        if creneaux_vac:
            creneaux_ouverts = creneaux_vac
        elif jour_cap in ('Mercredi', 'Samedi'):
            creneaux_ouverts = params.get('creneaux_ms', [])
        else:
            creneaux_ouverts = params.get('creneaux_mjv', [])

        swap_map = _construire_swap_map(jour_cap, sam_type, semaine_num,
                                          roulement_type, roulement_exceptions)

        solution, alertes, depas_jour = solve_day(
            jour=jour_cap,
            date_str=date_str,
            creneaux_ouverts=creneaux_ouverts,
            agents_eligibles=agents_eligibles,
            affectations=affectations,
            categories=categories,
            responsables=responsables,
            pause_flex=pause_flex,
            priorite_rdc=priorite_rdc,
            horaires_agents=horaires_agents,
            evenements=evenements,
            besoins_jeunesse=besoins_jeunesse,
            planning_type_jour=pt_jour,
            roulement_agents=roulement_type,
            samedi_type=sam_type,
            periode=periode_effective,
            mode_vac=mode_vac,
            swap_map=swap_map,
            presences_vac=presences_vac,
            cumul_hebdo_avant=cumul_hebdo,
        )

        for a, d in depas_jour.items():
            cumul_hebdo[a] = cumul_hebdo.get(a, 0) + d

        resultats_jours.append({
            'date': date_str,
            'jour': jour_maj,
            'jour_cap': jour_cap,
            'sam_type': sam_type,
            'creneaux': creneaux_ouverts,
            'solution': solution,
            'infaisable': solution is None,
            'alertes': alertes,
            'cumul_hebdo_apres': dict(cumul_hebdo),
        })

    return {'semaine_num': semaine_num,
            'conflits_a_signaler': lecture_resultat['conflits'],
            'jours': resultats_jours}


def resumer_calcul(resultat):
    """Résumé texte simple, pour affichage / débogage."""
    lignes = [f"Semaine {resultat['semaine_num']} — {len(resultat['jours'])} jour(s) recalculé(s) :"]
    for j in resultat['jours']:
        statut = "❌ INFAISABLE" if j['infaisable'] else "✅ solution trouvée"
        lignes.append(f"  {j['jour']} ({j['date']}) — {statut}"
                       + (f", {len(j['alertes'])} alerte(s) de remplissage" if j['alertes'] else ""))
        if j['alertes']:
            for cren_idx, section, msg in j['alertes']:
                lignes.append(f"      [{section}] créneau {cren_idx} : {msg}")
    return '\n'.join(lignes)
