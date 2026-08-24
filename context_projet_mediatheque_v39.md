# Contexte Projet — Planning Médiathèque
*Version : v39 — 24 août 2026 (insensibilité à la casse généralisée à toutes les recherches d'onglets, cf. §25.10)*

---

## 1. PRÉSENTATION

Application Streamlit (déployée sur Streamlit Cloud, dépôt GitHub
`planning-mediatheque`) : une seule page, en défilement, organisée en 3
blocs indépendants (voir §17) :
1. Créer l'onglet Événements à partir des fichiers sources bruts
2. Générer le planning mensuel (fusion Événements + Préparation → moteur
   CP-SAT → fichier Excel final)
3. Vérifier un planning déjà rempli/modifié à la main (pas encore construit)

**Fichiers projet actifs (à jour en permanence — une seule version de
chacun, jamais de doublon numéroté) :**
- `app.py` — Streamlit, **fonctionnel** depuis le 18/08 (blocs 1 et 2 ;
  bloc 3 encore un aperçu non branché). Réécrit de zéro cette session-là,
  n'importe plus du tout `planning_engine.py` / `excel_writer.py`.
- `sources_to_evenements.py` — génère l'onglet Événements (bloc 1). Très
  largement corrigé le 18/08 après tests sur les vrais fichiers sources
  (voir §17) — les descriptions de structure du §15 ci-dessous sont
  **obsolètes sur plusieurs points**, se fier au §17 en cas de divergence.
- `planning_engine_cpsat.py` — moteur CP-SAT (version active). Bug Jeunesse
  corrigé le 18/08 (voir §17).
- `generate_planning_excel_septembre.py` — génère le planning Excel final
  formaté (voir §8). Rendu paramétrable le 18/08 (accepte n'importe quel
  fichier de préparation, pas seulement celui de septembre — le nom de
  fichier est trompeur mais gardé tel quel pour ne pas casser les imports
  existants). Bug d'affichage "vue par agent" corrigé le 18/08 (voir §17).
  **24/08 : `Paramètres` et `Affectations` ne sont plus recopiés en
  très masqué — embarqués visibles et modifiables (comme Planning_type/
  horaires d'équipes, mais éditables) ; `embarquer_horaires_agents_visible`
  change de comportement : la grille est désormais modifiable elle aussi
  (plus verrouillée en lecture seule). Voir §25.**
- `planning_checker.py` — vérifie un planning déjà rempli (Bloc 3) et sert
  aussi de brique de lecture à la régénération partielle (Bloc 4, via
  `charger_donnees_preparation`, réutilisée telle quelle par
  `regeneration_lecture.py`). **24/08 : sait enfin lire la grille "horaires
  d'équipes" (avant, seule l'ancienne liste à plat `Horaires_Des_Agents`
  était lue, y compris pour le Bloc 4) ; détection par mise en page
  (case A6/H6), tolérante à un onglet renommé. Voir §19 pour sa création,
  §21 pour ses 2 premiers correctifs, §25 pour celui-ci.**
- `regeneration_lecture.py` / `regeneration_calcul.py` / `regeneration_ecriture.py`
  — les 3 briques de la "régénération partielle" (Bloc 4, voir §20). Le
  bloc 4 de `app.py` (upload → semaine → jour(s) → régénération) est
  **branché et fonctionnel** (confirmé le 24/08 — l'item correspondant de
  la liste "à faire" du §23 est donc fait, la fusion exacte n'a pas été
  documentée dans une session dédiée).
- `compare_planning.py` — compare un planning calculé au fichier de
  référence Excel (voir §10). Pas touché depuis le 18/08.
- `requirements.txt` — **nouveau fichier obligatoire** pour Streamlit Cloud
  (absent avant le 18/08, cause de plusieurs pannes de déploiement) :
  `streamlit`, `pandas`, `openpyxl`, `ortools`.

**Fichiers legacy — à supprimer du dépôt** (plus importés par rien depuis
la réécriture de `app.py` le 18/08) : `excel_writer.py`, `planning_engine.py`
(ancien moteur glouton).

⚠️ **Piège GitHub identifié le 18/08** : en re-uploadant un fichier déjà
existant sans utiliser la fonction "remplacer" (crayon → Edit), GitHub (ou
l'OS local) peut créer un doublon nommé `nom_fichier (1).py` au lieu
d'écraser l'original — l'app plante alors avec une erreur d'import
générique ("Oh no" / ModuleNotFoundError) sans dire pourquoi. Toujours
vérifier le nom exact du fichier listé sur GitHub après un upload.

⚠️ À faire au début de chaque session : uploader la dernière version des
fichiers `.py` actifs depuis les outputs si plus récentes que celles du
projet (voir liste ci-dessus).

---

## 2. RÉSULTATS (mai 2026)

| Étape | Différences vs référence |
|-------|--------------------------|
| Ancien moteur glouton | ~82 |
| CP-SAT sans présences vacataires | ~70 |
| CP-SAT avec présences vacataires (session précédente) | 23 |
| CP-SAT après corrections bugs 08/2026 (avant réglage vacataire) | ~94 |
| **CP-SAT après réglage vacataire V1/V2 08/2026** (hors Eloïse, hors 30 mai) | **70** |
| ...dont dû au 6 mai (arbitrage manuel non codé, voir §10) | 18 |
| **Total hors 6 mai (état "brut" avant tout arbitrage manuel)** | **52** |
| Objectif original | < 10 (à revoir : voir §10 sur la nature des différences restantes) |

**Historique du chiffre 94 → 70** : la comparaison automatique complète (script
`compare_planning.py`) a d'abord révélé ~94 différences après correction des bugs bruts
(§9), un chiffre plus élevé que les 23 précédents car c'était la première comparaison
exhaustive jamais faite sur l'intégralité du mois avec le vrai fichier de référence. Une
fois le réglage vacataire V1/V2 mis en place (§7) et les faux-positifs écartés (agent
"Eloïse" non modélisé, journée du 30 mai hors-sujet), le chiffre réel est retombé à 70,
dont 18 concentrées sur le seul 6 mai (qui représente un niveau de finesse — l'arbitrage
manuel — volontairement non codé pour l'instant).

---

## 3. AGENTS

| Agent | Resp | Pause flex | Sections (priorité) | Cat |
|-------|------|-----------|---------------------|-----|
| Marie-France | | | RDC, Adulte, MF | |
| Anne-Françoise | OUI | OUI | Adulte, Jeunesse, MF, RDC | A |
| Christine | | OUI | Adulte, RDC | A |
| Léa | | | Adulte, MF, RDC, Jeunesse | A |
| Chloé | | | Adulte, RDC, Jeunesse | |
| Macha | | OUI | RDC, Adulte | |
| Delphine | OUI | OUI | MF, RDC, Jeunesse, Adulte | A |
| Barbara | | | MF, Jeunesse | |
| Stéphane | | | MF uniquement | |
| Stéphanie | OUI | | Jeunesse, RDC | |
| Robin | | | Jeunesse, RDC | |
| Guillaume | | | Jeunesse, RDC | |
| Agnès | | | Jeunesse | |
| Tiphaine | | | Jeunesse, RDC, MF, Adulte | |
| Vacataire 1 | | | Jeunesse, MF, Adulte | VAC |
| Vacataire 2 | | | Jeunesse, MF, Adulte | VAC |

- Lydie : partie (supprimée)
- Cat **A** : sections 1 et 2 équivalentes
- Stéphane : MF uniquement (dure absolue)
- Vacataires : jamais en RDC (dure absolue)
- Barbara exception : 5h consécutives MF samedi après-midi (validé)

---

## 4. STRUCTURE FICHIER DE PRÉPARATION

### Onglet Paramètres

Clés importantes :
```
Liste_des_créneaux_mardi_jeudi_vendredi   10:00-12:30;12:30-13:00;...
Liste_des_créneaux_mercredi_samedi        10:00-11:00;11:00-12:00;...
Samedi_1..4                               ROUGE / BLEU
Semaine_1..4                              Hors Vacances Scolaires / Vacances Scolaires
```

**Tableau Présence Vacataire** (début de la section Paramètres, après les clés) :
| Date | Vacataire | Heure début | Heure fin |
|------|-----------|-------------|-----------|
| 06-mai-26 | Vacataire 1 | 10h | 19h |
| 06-mai-26 | Vacataire 2 | 10h | 19h |
| 09-mai-26 | Vacataire 1 | 10h | 19h |
| 13-mai-26 | Vacataire 1 | 10h | 19h |
| 13-mai-26 | Vacataire 2 | 13h | 19h |
| 23-mai-26 | Vacataire 1 | 10h | 19h |
| 30-mai-26 | Vacataire 1 | 10h | 19h |
| 30-mai-26 | Vacataire 2 | 10h | 19h |

⚠️ Dates en format Excel date natif. Heures en texte "10h", "13h30".  
⚠️ Si tableau présent → vacataires éligibles UNIQUEMENT sur les dates listées.

### Autres onglets

- `Planning_type` ou `planning_type` (P maj géré) : Col A=créneau, B=RDC, C=Adulte, D=M&F, E=Jeunesse. Séparateur ` / `.
- `Événements` : Date texte FR | Début "10h" | Fin | Nom | Agents (;-séparés)
  - ⚠️ Événement SANS agents = ne bloque PERSONNE
- `Horaire_ouverture_mediatheque` : Jour | Début S1 | Fin S1 | Début S2 | Fin S2
- **"horaires d'équipes"** (grille collaborative, cf. §24) : remplace l'ancienne
  liste à plat `Horaires_Des_Agents`. **Détecté par sa mise en page, pas par
  son nom d'onglet** (repérage des cases A6="ADULTES" / H6="JEUNESSE") — Elo
  peut l'appeler comme elle veut dans son fichier de préparation
  (`horaires_des_agents`, `Horaires agents`, peu importe), le moteur le
  retrouve tout seul. L'ancien nom `Horaires_Des_Agents` (liste à plat) reste
  utilisable en repli si l'onglet grille n'est pas présent.

---

## 5. CALENDRIER

- Commence au **premier MARDI** du mois (jours avant exclus)
- `Samedi_N` = samedi de la N-ième semaine complète
- Génère N semaines = nombre de Samedis définis

---

## 6. FONCTIONS PRINCIPALES

```python
parse_parametres(raw)  → {mois, annee, creneaux, creneaux_mjv, creneaux_ms,
                           samedis, semaines, mode_vac, presences_vac}
build_calendar(...)    → [{num, jours: [{date, jour, samedi_type}]}]
solve_day(...)         → {c_idx: {section: [agents]}} ou None
```

**Créneaux par jour** (⚠️ règle corrigée le 08/2026 — `params['creneaux']` seul ne
doit JAMAIS être utilisé, c'est une liste fusionnée invalide) :
```python
creneaux_vacances_jour = (construire_grille_vacances_jour(jour, sam_type)
                           if 'Hors' not in periode_effective else [])
if creneaux_vacances_jour:
    creneaux_ouverts = creneaux_vacances_jour  # Jour "vacances" : grille construite
                                                # PAR JOUR depuis Besoins_Jeunesse
elif jour in ('Mercredi', 'Samedi'):
    creneaux_ouverts = params['creneaux_ms']   # Mercredi, Samedi (hors vacances)
else:
    creneaux_ouverts = params['creneaux_mjv']  # Mardi, Jeudi, Vendredi (hors vacances)
```

**Grille "vacances" — corrigée une 2e fois le 08/2026** : construite à partir des
tranches fines de `Besoins_Jeunesse`, mais les sous-tranches CONSÉCUTIVES ayant le
MÊME besoin en Jeunesse sont fusionnées, et cette fusion ne dépasse JAMAIS les limites
des blocs standards (`creneaux_mjv`/`creneaux_ms`) — sinon ça fusionnait aussi les
blocs RDC/Adulte/MF du planning-type, qui ont leurs propres frontières (bug détecté :
un bloc de 14h-19h s'est retrouvé fusionné en un seul créneau, cassant tout le
remplissage). Résultat : un bloc homogène comme "15h30-17h" reste un seul créneau
(au lieu d'être fragmenté à tort en 15h30-16h + 16h-17h avec 2 agents différents),
mais un vrai changement de besoin (ex: 17h-18h=2 personnes puis 18h-19h=1) reste
bien scindé en 2 créneaux séparés.

**Période effective** (mode vacances) : le réglage `Semaine_N` (par semaine) sert de défaut,
mais un jour marqué "vacances" dans l'onglet **Jours_speciaux** prime dessus au cas par cas
(ex : un jour ponctuel en vacances au sein d'une semaine "Hors Vacances").

**Éligibilité vacataires :**
```python
if presences_vac:  # tableau présent → exclusif
    eligible = date_str in presences_vac and agent in presences_vac[date_str]
else:
    eligible = jour in mode_vac  # fallback
```

---

## 7. CONTRAINTES (mises à jour 08/2026 — liste complète et vérifiée dans le code)

### Dures (bloquent toujours, jamais d'exception)

| Code | Règle |
|------|-------|
| A1 | Habilitations respectées (`affectations`) |
| A2 | Vacataires jamais en RDC |
| A3 | Stéphane MF uniquement |
| A4 | Max 1 agent par section/créneau (RDC/Adulte/MF) |
| D13 | 1 agent = 1 section maximum par créneau (pas de double-affectation simultanée) |
| B1/B2 | Disponibilité contractuelle (`Horaires_Des_Agents`) + événements bloquants |
| B3 | Vacataires uniquement les jours où ils sont dans `Présence Vacataire` (ou `mode_vac` en repli si le tableau est vide) |
| C3 | Pause déjeuner ≥ 60 min réellement libres dans la fenêtre 12h-14h (réguliers sauf Delphine ; vacataires les jours où ils sont présents). **Corrigée 08/2026**, voir §9 |
| D1 | Roulement samedi ROUGE/BLEU + exceptions individuelles |
| K3 | Vacataire seul en Jeunesse uniquement autorisé 12h-14h (sinon il faut au moins 1 régulier avec lui) |
| E3 | Événement sans agents listés = ne bloque personne |

### Souples (pénalités, ordre de priorité du plus fort au plus faible)

```
D_FILL = 5000  (ex-contrainte dure, rendue souple 08/2026)
           → Le PT prévoit quelqu'un en RDC/Adulte/MF → le solveur DOIT très
             fortement y mettre exactement 1 agent (remplaçant si l'agent prévu
             est absent)
           → Si structurellement impossible (conflit réel, ex: Vendredi 15 mai) :
             le créneau reste vide + ALERTE, au lieu de rendre tout le jour
             infaisable
           → Priorité la PLUS HAUTE de toutes les pénalités

JEUNESSE = 200  (ex-contrainte dure F1/F2, rendue souple 08/2026)
           → Nb agents Jeunesse visé = PT (hors vacances) ou Besoins_Jeunesse
             (vacances, via tranches fines fusionnées — voir §6)
           → Volontairement plus FAIBLE que D_FILL : en cas de conflit (un seul
             agent dispo pour 2 besoins simultanés), le solveur privilégie
             D_FILL et laisse un manque en Jeunesse (+ ALERTE), conforme à ce
             qu'on observe dans le planning de référence (ex: Vendredi 15 mai
             17h-18h : Barbara reste au M&F, Jeunesse reste à 1 au lieu de 2)
           → Plafonné aux agents réellement disponibles (`nb_requis = min(pt, possible)`)

CONSÉCUTIF = 150 par violation (souple depuis la session précédente)
           → Max 2h30 d'affilée Mar/Jeu/Ven ; max 4h Mer/Sam (exception Barbara
             samedi après-midi : 5h, validée par la directrice)
           → Permet un dépassement si les absences l'imposent, plutôt que de
             bloquer tout le jour

G1 = 100  Suivre l'agent prévu au planning-type (PT) dans sa section
           → Réduit à 30 sur Adulte/MF quand un vacataire est présent ce
             jour-là (peu importe Mercredi/Samedi ou non — Présence Vacataire
             est la seule source de vérité, plus rien câblé en dur sur un jour
             de semaine particulier)
           → RDC garde G1=100 même les jours avec vacataire (vacataires jamais
             en RDC de toute façon)

G2 = 50   Remplaçant idéalement de la même section principale que l'absent

J1 = 30   Agent dans sa section principale (ou les 2 premières si catégorie A)
           → Section 3e position : poids 30 ; section 4e : poids 70 (fortement
             déconseillé)
           → Vacataires exemptés de J1 quand ils sont présents (peu importe le
             jour) : ils remplissent librement les sections laissées libres

J3 = 25   Responsables de section déprioritisés (moins sollicités si possible)

VACATAIRE 1 = bonus (pénalité négative), priorité de section Jeunesse(90) >
           MF(70) > Adulte(50) — chaque créneau travaillé par Vacataire 1 est
           RÉCOMPENSÉ, pour maximiser son usage tant qu'il est présent
           (règle utilisatrice précisée 08/2026)

VACATAIRE 2+ = 10 par créneau travaillé (pénalité, "dernier recours")
           → N'est utilisé que si un régulier ne peut vraiment pas couvrir
             (D_FILL=5000 et Jeunesse=200 le forcent quand même si nécessaire)
           → Le choix fin de QUI il remplace ensuite (créneaux les plus longs,
             puis journées les plus chargées à soulager) reste volontairement
             NON codé — trop fin/subjectif, fait à la main par l'utilisatrice
             après génération (cf. exemple du 6 mai, §10)

I1 = 20   Non-fragmentation : préférer des blocs de créneaux consécutifs pour
           un même agent plutôt que des créneaux isolés dans la journée

ÉQUITÉ DES REMPLACEMENTS (nouvelle règle 08/2026, résolution en 2 PASSES —
           pas un simple poids, voir explication ci-dessous) :
           → Objectif : si plusieurs agents interchangeables doivent dépasser
             leurs heures PT pour remplacer des absents, répartir ce
             dépassement plutôt que le concentrer sur un seul (ex: si Léa fait
             +2h de remplacement, Chloé doit aussi en faire +2h plutôt que 0)
           → Responsables de section EXCLUS (traités à part via J3)
           → Franchise de 60 minutes avant que l'écart ne soit pénalisé
```

**Pourquoi l'équité est résolue en 2 passes, pas juste un poids de plus** : un
simple poids, même faible, pouvait faire préférer une petite alerte Jeunesse à
un léger déséquilibre d'heures sur certains jours très chargés en absences,
ce qu'on ne voulait surtout pas. Solution : le solveur calcule d'abord la
MEILLEURE solution possible sur toutes les priorités structurelles (D_FILL,
Jeunesse, G1, consécutif...) SANS tenir compte de l'équité ; puis, cette
qualité structurelle figée comme contrainte, il cherche en plus la répartition
la plus équitable. Ça garantit mathématiquement que l'équité ne peut jamais
dégrader le reste — pas besoin de deviner le bon poids.

**Système d'alertes** : `solve_day` retourne `(solution, alertes)` au lieu de
`solution` seule. `alertes` est une liste `[(cren_idx, section, message)]` qui
documente chaque créneau où le besoin n'a pas pu être entièrement couvert
malgré la pénalité forte — affiché dans le planning Excel généré par une
bordure rouge + commentaire sur la cellule concernée (voir §8).

**Nettoyage 08/2026** : le dictionnaire `POIDS` contenait plusieurs entrées
mortes (H1, H2, H3, K1, K2, I2) — jamais utilisées dans le code, remplacées
par les mécanismes ci-dessus codés directement dans `solve_day`. Supprimées
pour éviter toute confusion future. `POIDS` ne contient plus que :
`G1_planning_type`(100), `G2_meme_section_repl`(50), `J1_section_principale`(30),
`J3_responsable`(25), `I1_non_fragmentation`(20).

---

## 7bis. ÉQUITÉ HEBDOMADAIRE (nouveau, 08/2026)

**Problème identifié** : l'équité H2 du §7 est calculée **uniquement à
l'intérieur d'une même journée** — `solve_day` traite chaque jour de façon
indépendante, sans aucune mémoire des jours précédents. Résultat : si le même
agent est le seul remplaçant "à égalité" plusieurs jours d'affilée dans la
même semaine, rien ne l'empêche d'être choisi systématiquement (le compteur
d'équité repart de zéro chaque jour).

**Règle validée par l'utilisatrice** : ajouter un 2e niveau d'équité, calculé
**sur la semaine**, en plus (et non à la place) du niveau journalier existant :
- Franchise **journalière** inchangée : 60 min (§7, H2).
- Franchise **hebdomadaire** nouvelle : 3h (180 min) tolérées sur l'ensemble
  de la semaine avant pénalisation.
- Les deux se cumulent dans le même objectif secondaire (`penalites_equite`,
  2e passe de résolution) — un remplacement qui dégrade à la fois l'équilibre
  du jour ET creuse l'écart de la semaine est donc davantage évité qu'un
  remplacement qui n'en dégrade qu'un seul.
- L'écart mesuré reste toujours le **dépassement par rapport au PT de chacun**
  (minutes travaillées − minutes prévues au PT pour cet agent ce jour-là),
  jamais un total d'heures brut — un agent à 17h/semaine au PT et un autre à
  12h/semaine ne sont donc jamais comparés sur leur volume total, seulement
  sur leur écart individuel par rapport à leur propre référence.
- Comme pour le niveau journalier, cette équité ne peut **jamais empêcher un
  remplacement réellement nécessaire** : elle ne fait que départager entre
  choix par ailleurs strictement équivalents pour les priorités structurelles
  (D_FILL, Jeunesse, G1, consécutif...). Si une seule personne peut couvrir,
  elle est mise, même si son compteur hebdo est déjà chargé.

**Implémentation technique** :
- `solve_day` reçoit un nouveau paramètre `cumul_hebdo_avant` (dict
  `{agent: minutes de dépassement net déjà cumulées cette semaine, avant ce
  jour}`), fourni par `compute_full_planning`.
- `solve_day` retourne désormais **3 valeurs** (et non 2) :
  `(result, alertes, depas_jour)` — `depas_jour` est le dépassement NET de ce
  jour par agent (peut être négatif), à cumuler par l'appelant.
- `compute_full_planning` maintient un dict `cumul_hebdo`, **remis à zéro à
  chaque nouvelle semaine** (chaque itération de la boucle `for semaine in
  calendrier`), mis à jour après chaque jour traité :
  `cumul_hebdo[a] += depas_jour[a]`.
- Chaque entrée de `week_plan['jours']` contient désormais aussi
  `cumul_hebdo_apres` (dict), utile pour du debug/traçabilité si besoin.
- Testé (08/2026) sur un scénario synthétique à 2 remplaçants strictement
  équivalents sur 3 jours consécutifs de la même semaine : le moteur alterne
  bien entre les deux (B, puis A, puis B) au lieu de toujours choisir le
  même — comportement confirmé avant livraison.

⚠️ **Testé sur données réelles mai 2026 (08/2026)** : voir §7ter ci-dessous — résultat
net : -5 différences (89 → 84), mais avec un bug supplémentaire trouvé et corrigé au
passage (verrou manquant RDC/Adulte/MF, indépendant de l'équité hebdo elle-même).

---

## 7ter. VALIDATION SUR DONNÉES RÉELLES (08/2026) — équité hebdo + bug du verrou manquant

**Contexte** : premier test de l'équité hebdo (§7bis) sur le vrai fichier mai 2026 via
`compare_planning.py`. Résultat inattendu au premier essai : **92 différences, pire que
le baseline (89)** avant tout changement.

**Cause identifiée — bug distinct, pas un défaut de l'équité hebdo elle-même** : sur les
sections RDC/Adulte/MF, rien n'empêchait le solveur d'affecter un agent à un créneau où
le planning-type ne prévoit **personne** (contrairement à Jeunesse, qui avait déjà ce
verrou : `sum_j == 0` si besoin nul). Ce trou ne coûtait rien dans le calcul de score —
le solveur pouvait laisser vide OU remplir sans différence, au hasard de son ordre de
recherche interne. Ajouter les nouvelles variables d'équité hebdo a suffi à faire
basculer ce choix arbitraire sur plusieurs créneaux (ex : vendredi 29 mai 14h-15h, un
creux hors planning-type, s'est retrouvé rempli à tort en RDC + Adulte + MF).

**Correction appliquée** : verrou symétrique à celui de Jeunesse — pour chaque créneau et
chaque section RDC/Adulte/MF où le PT ne demande personne, le nombre d'agents affectés
est désormais forcé à 0 (`model.add(sum(...) == 0)`), au lieu d'être laissé libre.

**Résultat final (correctif + équité hebdo combinés)** :

| Version | Différences totales (brut, sans exclusions) |
|---------|---------------------------------------------|
| Baseline (avant toute modif de cette session) | 89 |
| Équité hebdo seule (bug du verrou pas encore trouvé) | 92 (pire) |
| Correctif du verrou seul (sans équité hebdo) | 90 |
| **Correctif + équité hebdo (livré)** | **84** |

- **12 différences corrigées**, dont plusieurs cas typiques d'arbitrage entre agents
  équivalents visés par l'équité hebdo (ex : 7 mai RDC/Adulte Anne-Françoise↔Léa, 12/13/16
  mai arbitrages Jeunesse Robin/Agnès/Stéphanie).
- **7 nouvelles différences apparues**, principalement des bascules dans l'autre sens —
  notamment plusieurs créneaux MF (9 et 23 mai) où la référence attend Vacataire 1 mais
  où un régulier est désormais choisi. À surveiller : possible interaction entre l'équité
  hebdo (qui ne devrait pourtant pas s'appliquer aux vacataires, exclus de
  `agents_equite`) et le choix régulier/vacataire sur ces créneaux précis.
- **Bilan net : -5 différences** (89 → 84). Amélioration réelle mais modeste — l'équité
  hebdo, comme tout mécanisme de départage entre choix à égalité de coût, peut faire
  pencher la balance dans un sens ou dans l'autre selon les jours ; elle ne garantit pas
  de n'améliorer QUE les cas ciblés.

⚠️ **Non encore fait** : recalculer le chiffre "hors Eloïse / 30 mai / 13 mai vacataire"
(§10-11) sur cette nouvelle version pour le comparer proprement au 70 précédent — le 84
ci-dessus est un chiffre BRUT (sans exclusions), pas directement comparable au 70 du §10.
Le script `compare_planning.py` du projet ne fait actuellement AUCUNE exclusion
automatique malgré ce qu'indiquait une note précédente de ce contexte — à vérifier/corriger.

---

## 7quater. CORRECTIONS 09/2026 (avant génération planning septembre)

**Contexte** : premier lancement sur le fichier de préparation de septembre 2026 (5
semaines, et apparition d'un 3e vacataire — "Vacataire 3" — jamais vu jusque-là). Deux
bugs supplémentaires trouvés à cette occasion, indépendants de l'équité hebdo :

**1. Reproductibilité du solveur (corrigé)** : `num_search_workers = 4` sans graine fixée
faisait varier le résultat d'un lancement à l'autre, même sur des données identiques (le
"84" annoncé en fin de session précédente était en réalité un tirage chanceux, pas une
vraie mesure). Une graine seule (`random_seed`) ne suffit PAS à garantir un résultat
identique en recherche parallèle (l'ordre d'arrivée des 4 chercheurs dépend du temps réel
d'exécution, pas seulement de la graine) — il faut aussi repasser à
`num_search_workers = 1` pour un déterminisme total. **Chiffre de référence fiable
obtenu après ce correctif (avec verrou + équité hebdo) : 86 → 84 brut (baseline → final),
63 → 60 hors exclusions.**

**2. Vacataires au-delà de 2 — écrasement silencieux (corrigé, bug potentiellement
présent depuis l'origine du projet)** : deux endroits du code normalisaient tout nom
contenant "Vacataire" en 'Vacataire 2' si le texte contenait un "2", et en 'Vacataire 1'
SINON — ce qui écrasait silencieusement tout "Vacataire 3" (ou plus) en "Vacataire 1".
Corrigé par extraction du numéro réel via regex (`re.search(r'\d+', ...)`). Sans ce
correctif, "Vacataire 3", présente le 19 septembre dans le fichier de préparation,
aurait été totalement invisible pour le moteur (fusionnée à tort avec Vacataire 1).

**3. Tableau "Présence Vacataire" — première ligne systématiquement ignorée (corrigé,
bug présent depuis l'origine du projet, découvert en re-testant mai)** : la ligne
d'en-tête du tableau ("Présence Vacataire" en colonne A) contient AUSSI les libellés de
colonnes (Date/Vacataire/Heure début/Heure fin) sur la MÊME ligne — il n'y a donc qu'UNE
seule ligne d'en-tête, pas deux. Le code sautait `+2` lignes au lieu de `+1`, ignorant
systématiquement la toute première ligne de données du tableau, sur TOUS les fichiers de
préparation depuis le début du projet. Concrètement sur mai : "6 mai, Vacataire 1"
n'avait jamais été vue par le moteur. Une bonne partie des écarts attribués jusqu'ici à
"l'arbitrage manuel fin du 6 mai" (§10) provenait probablement en réalité de ce bug de
lecture, pas d'un raffinement manuel non modélisable. Après correction, les écarts du 6
mai passent de 19 à 15 sur cette seule journée (léger effet de bord ailleurs dans le
mois, total global quasi stable : 85 brut / 61 hors exclusions).

⚠️ **Ce dernier point mérite d'être creusé une prochaine fois** : si la lecture des
données vacataires était fausse depuis le début, les conclusions du §10 sur "18
différences dues à l'arbitrage manuel du 6 mai, hors scope" doivent être reconsidérées —
ce n'était peut-être pas (ou pas seulement) un choix éditorial hors-scope, mais un vrai
bug de lecture.

---

## 7quinquies. RÈGLE CALENDRIER — semaine à cheval sur deux mois (09/2026)

**Règle métier précisée par l'utilisatrice** : une semaine SP (Mardi→Samedi) **entamée**
dans le mois en cours doit être **terminée jusqu'au Samedi**, même si Jeudi/Vendredi/
Samedi tombent sur le mois suivant. Le mois suivant démarre alors à son propre premier
Mardi, APRÈS la fin de cette semaine à cheval (pas de chevauchement, pas de trou).

Exemple concret : septembre 2026 se termine un mercredi (30/09). La semaine 5
(Mardi 29 → Samedi 3 octobre) doit donc inclure Jeudi 1er octobre, Vendredi 2 octobre et
Samedi 3 octobre dans le planning DE SEPTEMBRE. Le planning d'octobre commencera
directement au Mardi 6 octobre (première semaine complète d'octobre).

**Bug corrigé** : `build_calendar` s'arrêtait auparavant strictement à la fin du mois
calendaire (`while d.month == mois_num`), coupant net une semaine entamée sans jamais la
terminer — la semaine 5 de septembre s'arrêtait au mercredi 30, jeudi/vendredi/samedi
disparaissaient purement et simplement du planning généré. Corrigé : la boucle continue
tant qu'on est dans le mois OU qu'une semaine est en cours de construction (elle se
termine alors forcément au Samedi suivant, donc pas de risque de boucle infinie, et
aucune NOUVELLE semaine n'est démarrée une fois sorti du mois).

Effet de bord corrigé en même temps dans `generate_planning_excel_septembre.py` (script
de génération Excel) : le libellé de chaque jour et le titre de chaque semaine
affichaient auparavant un nom de mois figé (celui du mois demandé) — desormais calculés
depuis la date réelle de chaque jour, pour afficher correctement "1 Octobre" et non
"1 Septembre" sur les jours débordants.

---

## 7sexies. BUG CORRIGÉ — événements/congés invisibles si la date n'a pas d'année (09/2026)

**Symptôme signalé par l'utilisatrice** : le planning Excel de septembre généré avait ses
colonnes Accueil/Animation/Réunion/Absence entièrement vides, alors que l'onglet
"Événements" du fichier de préparation contenait bien 35 lignes (congés, réunions,
événements publics...).

**Cause** : `_parse_fr_date` exigeait TOUJOURS un jour + un mois + une ANNÉE explicites
dans le texte de la cellule (ex: "mardi 5 mai 2026"). Le fichier de mai écrivait bien
l'année, mais celui de septembre écrit les dates SANS année ("mardi 1 septembre") — la
fonction retournait alors `None` pour chaque ligne, et `parse_evenements` ignorait
silencieusement les 35 lignes en bloc.

**Corrigé** : `_parse_fr_date` accepte maintenant un paramètre `annee_defaut`, utilisé
quand aucune année n'est trouvée dans le texte. `parse_evenements(raw, annee_defaut=...)`
le transmet, alimenté par l'année lue dans l'onglet Paramètres (`params['annee']`). Mis à
jour dans `compute_full_planning` (moteur) ET dans le script `generate_planning_excel_*`
(qui appelle `parse_evenements` séparément pour construire les colonnes du fichier Excel).

⚠️ **Point de vigilance pour la suite** : si un mois à cheval sur deux années civiles est
un jour traité (ex: semaine à cheval décembre→janvier), une seule année par défaut ne
suffira plus — à surveiller le cas échéant, mais non pertinent pour septembre 2026.

---

## 7septies. BUG CORRIGÉ — exceptions de roulement Samedi (ex: "Stéphane Bleu semaine 1") ignorées (09/2026)

**Symptôme signalé par l'utilisatrice** : sur la 2e version du fichier de préparation
septembre, l'onglet Roulement_Samedi précise que Stéphane fait exceptionnellement le
roulement BLEU en semaine 1 (au lieu de son ROUGE habituel) — le planning généré ne
reflétait pas cette exception.

**Cause** : `parse_roulement_samedi` exigeait que la cellule "numéro de semaine" de la
table d'exceptions soit un nombre PUR (`"1"`, testé avec `.isdigit()`). Le fichier
contient en réalité `"semaine_1"` (texte) — `.isdigit()` échoue sur ce texte, la variable
`current_sem` reste à `None`, et l'exception est silencieusement ignorée (jamais ajoutée
au dict `exceptions`).

**Corrigé** : extraction du numéro via regex (`re.search(r'\d+', ...)`) plutôt
qu'un test d'égalité strict — accepte "semaine_1", "Semaine 1", "S1", "1", etc.

**Effet observé une fois corrigé** : Stéphane apparaît bien disponible sur le Samedi
BLEU de la semaine 1 (5 septembre), et couvre plusieurs créneaux MF. Effet d'entraînement
normal (pas un bug) : sa disponibilité recalcule tout l'équilibre de la journée, donc
Delphine/Christine (qui le remplaçaient jusque-là ce jour précis) ne sont plus
nécessaires, et Vacataire 1 est davantage sollicitée par ricochet.

---

## 8. GÉNÉRATION EXCEL (script inline, pas excel_writer.py)

Produit onglets `Semaine_N` avec 9 colonnes A-I (Créneau, RDC, Adulte, M&F, Jeunesse, Accueil, Animation, Réunion, Absence).

### ⚠️ is_open_fixed — OBLIGATOIRE dans le script de génération

```python
def is_open_fixed(jour, cs, ce):
    """Fusionne les plages adjacentes pour éviter faux "—" (ex: Mercredi 12h-13h)."""
    ranges = sorted(hor_ouv.get(jour, []))
    if not ranges: return False
    merged = [list(ranges[0])]
    for s, e in ranges[1:]:
        if s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    return any(cs >= s and ce <= e for s, e in merged)
```

---

## 9. BUGS CORRIGÉS

| Bug | Correction |
|-----|-----------|
| Événement sans agents bloquait tout le monde | `if not ev['agents'] or agent not in ev['agents']` |
| Calendrier incluait 1er-2 mai en S1 séparée | `build_calendar` part du premier Mardi |
| samedis_params mal numérotés | Comptage depuis le premier Mardi |
| `Planning_type` P majuscule non reconnu | `raw.get('Planning_type') or raw.get('planning_type')` |
| Séparateur virgule au lieu de ` / ` | `' / '.join(lst)` |
| Vendredi infaisable avec trop d'absences | Contrainte consécutive rendue souple (pénalité 150) |
| Mercredi 12h-13h affiché "—" à tort | `is_open_fixed` avec fusion des plages adjacentes |
| Éligibilité vacataire ignorait le tableau Présence Vacataire | `agent_disponible` prend désormais `presences_vac` en paramètre : vérifie d'abord la date précise dans le tableau, puis se rabat sur `mode_vac` seulement si rien n'est précisé |
| Jeunesse infaisable si agents PT absents | Nombre exact exigé plafonné au nombre d'agents Jeunesse réellement disponibles (`nb_requis = min(nb_pt_jeunesse, nb_possible)`) au lieu d'exiger toujours le chiffre du PT |
| **`horaires_agents._presences_vac` — "étiquette fantôme"** | Cette référence n'était jamais créée nulle part : `hasattr(...)` retournait toujours False. `agent_disponible` prend maintenant `presences_vac` directement en paramètre, transmis explicitement à chaque appel (dans `solve_day` et dans `compute_full_planning`) |
| **Faute de frappe `roul_type`** (au lieu de `roulement_type`) | Faisait planter `compute_full_planning` sur TOUT samedi avec exception de roulement. Corrigé |
| **`params['creneaux']` utilisé pour tous les jours** | Cette clé est une fusion invalide des listes Mardi/Jeudi/Vendredi ET Mercredi/Samedi. Remplacé par une sélection explicite selon le jour (voir §6) |
| **Jours_speciaux jamais utilisé** | La fonction `parse_jours_speciaux` existait mais son résultat n'était jamais branché dans `compute_full_planning`. Un jour marqué "vacances" ponctuellement (ex: 15 mai, au sein d'une semaine "Hors Vacances") était donc traité à tort en mode normal |
| **Grille horaire "vacances" manquante** | Aucune liste de créneaux "vacances" n'existe dans l'onglet Paramètres. Construite désormais depuis les tranches du tableau Besoins_Jeunesse (identiques pour tous les jours en mode vacances : 11 tranches de 10h à 19h, avec des demi-heures) |
| **Clé "Samedi Bleu" mal reconnue** | Le code cherchait `Samedi_bleu` (tiret bas, minuscule) mais l'Excel écrit `samedi bleu` (espace). Résultat : 0 personne exigée en Jeunesse tout le Samedi Bleu. Recherche désormais insensible à la casse/espaces/tirets |
| **D_FILL et F1/F2 rendues souples** | Auparavant contraintes dures → toute la journée devenait infaisable au moindre conflit réel (cf. Vendredi 15 mai : Barbara ne peut être à la fois en MF et en Jeunesse). Converties en pénalités très fortes (D_FILL=5000, Jeunesse=200) avec système d'alertes (§7) : le planning se génère quand même, avec les seules cases vraiment impossibles marquées en alerte |
| **Grille "vacances" fusionnait trop / pas assez** | 1er correctif : grille unique partagée par tous les jours (trop rigide, ne reflétait pas les variations par jour). 2e correctif (même session) : fusion des sous-tranches Besoins_Jeunesse à même besoin, mais fusion cantonnée à l'intérieur des blocs standards (jamais à cheval sur 2 blocs PT) — sinon un bloc de 5h pouvait se former par erreur et casser tout le remplissage RDC/Adulte/MF. Voir §6 |
| **Pause déjeuner jamais réellement imposée** | La contrainte comparait une somme sur (créneau × section) à son propre total de variables — toujours vraie car un agent ne peut être que dans 1 section à la fois (D13), donc la pause n'empêchait jamais rien. Un agent pouvait travailler 4h d'affilée en plein 12h-14h sans aucune pause (détecté : Macha, mercredi 13 mai). Corrigée : calcul de la durée réellement travaillée (en minutes) dans la fenêtre 12h-14h, doit laisser au moins 60 min libres |
| **Vacataires "en dernier recours" — contredisait la nouvelle règle utilisatrice** | `K1_vac_dernier_recours` pénalisait tout usage de vacataire les jours de semaine (sauf samedi), câblé en dur. Remplacé par la règle V1/V2 précisée par l'utilisatrice (Vacataire 1 maximisé, Vacataire 2 dernier recours seulement) — voir §7 |
| **Règle H2/H3 équité déclarée mais jamais codée** | Le poids existait dans `POIDS` depuis le début mais n'était référencé nulle part dans le code. Implémentée le 08/2026 sous forme de résolution en 2 passes (§7) suite à une demande précise de l'utilisatrice (Léa surchargée en remplacement vs Chloé sous-utilisée, mercredi 13 mai) |
| **`POIDS` contenait 6 entrées mortes** (H1, H2, H3, K1, K2, I2) | Jamais référencées dans le code, remplacées par les mécanismes ci-dessus. Supprimées pour éviter toute confusion future |

---

## 10. DIFFÉRENCES RESTANTES (mai 2026, après réglage vacataire V1/V2 08/2026)

**Total : 70** (hors "Eloïse", hors 30 mai — voir §11), dont **18 sur le seul 6 mai**.

| Famille | Nb approx. | Statut |
|---------|-----|--------|
| **6 mai — arbitrage manuel fin** | 18 | ✅ Compris, pas un bug. Le 6 mai dans la référence représente le résultat APRÈS ajustement manuel (Vacataire 2 placé intelligemment). Notre logique "brute" (V1 maximisé/V2 dernier recours) ne réplique pas ce raffinement — volontairement non codé, voir §7 |
| **13 mai — logique "brute", résiduel** | 8 | Bien compris après analyse détaillée : ~3 sont de l'arbitrage régulier pur (catégorie A, sans lien vacataire), ~4 sont des permutations mineures sur le placement exact de la pause déjeuner de Vacataire 1 (choix équivalents), 1 est un doublon "Eloïse" |
| **9, 23 mai — résiduel vacataire** | 15 | Même nature que le 13 mai : petites permutations sur les créneaux exacts, pas d'erreur de règle identifiée |
| **Arbitrage entre réguliers "catégorie A"** (Léa/Christine/Delphine/Marie-France sur RDC/Adulte) | ~15 | ✅ Compris, pas un bug : ces agents ont leurs 2 premières sections déclarées "équivalentes" (voir §3). Le solveur n'a aucune préférence entre eux → choix arbitraire qui ne colle pas toujours à la référence. Nécessiterait un critère de départage supplémentaire (ex. équité d'heures) pour être réduit |
| **Arbitrage Jeunesse entre agents habilités** (Robin/Tiphaine/Agnès/Stéphanie/Guillaume) | ~15 | Même nature que ci-dessus, appliqué à la section Jeunesse (beaucoup d'agents habilités, aucune préférence codée entre eux) |
| **Jeunesse : régulier en trop (Barbara+Tiphaine vs Tiphaine seule)** | ~4 | Cas où le PT vise 2 personnes mais la référence n'en montre qu'une — à surveiller si ça se reproduit souvent, pourrait indiquer que la règle J1 (section secondaire) est encore trop permissive certains jours |
| Jours "vacances" restants (15, 16 mai, hors résiduel vacataire ci-dessus) | ~12 | Grille horaire désormais correcte ; différences = arbitrages de personnel, pas des blocages |

**Conclusion de la session** : le moteur est maintenant stable (aucun blocage sur les 20
jours de mai), et la majorité des différences restantes sont **des choix arbitraires entre
options légitimement équivalentes**, pas des erreurs de règle. Réduire davantage ce chiffre
demanderait soit (a) coder l'arbitrage manuel fin (créneaux longs / heures totales — jugé
trop subjectif pour l'instant, cf. §7), soit (b) ajouter un critère de départage générique
pour les agents "équivalents" (ex. équité d'heures sur le mois) — à évaluer si utile.

**Script de comparaison** : `compare_planning.py` (nouveau, 08/2026) compare automatiquement
la sortie de `compute_full_planning` au fichier de référence Excel (`Mai_2026_planning.xlsx`,
onglets `Semaine_1..4`), créneau par créneau et section par section. Réutilisable pour toute
future comparaison. Gère la normalisation de coquilles connues (ex: "Vacataire1" sans espace).

---

## 11. ÉTAT DU MOTEUR EN PROJET

✅ **planning_engine_cpsat.py est à jour (08/2026)** avec toutes les corrections du §9 et
le réglage vacataire V1/V2 du §7. Fichier de référence à utiliser pour la suite : celui
livré dans les outputs de cette session.

⚠️ **Point ouvert — agent "Eloïse"** : c'est la **responsable de la médiathèque**. Elle ne
fait pas de service public de façon régulière — elle dépanne uniquement quand le planning
n'a pas de solution avec les agents de la liste "Horaires des agents" (clarifié par
l'utilisatrice, 08/2026). **Décision : ne pas la modéliser dans le moteur.** À la place, le
système d'alertes (§7, D_FILL/F1F2 souples) joue déjà ce rôle : quand aucun agent normal ne
peut couvrir un créneau, une alerte est levée plutôt qu'un blocage — c'est exactement
l'équivalent numérique du réflexe "on appelle Eloïse". Les différences liées à Eloïse dans
la comparaison (6 cas sur mai) sont donc à ignorer définitivement, pas à corriger.

✅ **Résolu — journée du 30 mai** : un événement spécial hors médiathèque avait lieu ce
jour-là, avec un planning à part où les vacataires étaient mobilisés différemment. Ce jour
est à exclure de toute comparaison automatique (déjà fait dans `compare_planning.py`).

⚠️ **Point mineur repéré, non corrigé** : les clés `Durée_SP_max_idéale` (2h30) et
`Durée_SP_max_tolérée` (4h) existent dans l'onglet Paramètres mais ne sont pas lues par le
moteur — ces valeurs restent codées en dur dans `solve_day` (`max_consec_defaut`). Sans impact
actuel car les valeurs codées en dur correspondent au fichier, mais fragile si le fichier
change un jour.

⚠️ **Vacataires — action utilisatrice en cours** : l'utilisatrice a indiqué vouloir supprimer
les lignes vacataires de l'onglet `Horaires_Des_Agents`, le tableau `Présence Vacataire` de
l'onglet Paramètres devenant la SEULE source de vérité pour leurs horaires. Le code le permet
déjà : `agent_disponible` utilise `presences_vac` en priorité, et ne se rabat sur
`horaires_agents` que si aucune entrée n'existe pour cette date précise (dans ce cas, un
vacataire absent du tableau de présence sera simplement indisponible — comportement correct).

---

## 12. PROCHAINES PRIORITÉS

0. **NOUVEAU 12/08 — à valider en priorité** : équité "événements comptent comme SP"
   (§13.16). Testée uniquement sur septembre 2026 (0 alerte/infaisable, aucune
   régression de couverture) — PAS ENCORE comparée au planning de référence mai 2026.
   Relancer `compare_planning.py` en tout premier à la prochaine session pour vérifier
   que ça réduit bien les buckets "arbitrage catégorie A" / "arbitrage Jeunesse" (~30
   différences, §10) sans en créer de nouvelles.
1. **NOUVEAU 12/08** : démarrer les plannings d'**octobre et novembre 2026** — prévoir
   le fichier de préparation Excel de chaque mois (même structure que septembre) au
   début de la session.
2. Décider si l'arbitrage manuel fin de Vacataire 2 (créneaux longs / heures totales,
   exemple du 6 mai) mérite d'être codé, ou reste manuel définitivement
3. Confirmer que l'utilisatrice a bien retiré les lignes vacataires de `Horaires_Des_Agents`
   dans son fichier Excel, puis re-régénérer pour vérifier la stabilité
4. (Mineur) Faire lire `Durée_SP_max_idéale` / `Durée_SP_max_tolérée` depuis l'Excel au lieu du code en dur
5. Brancher générateur Excel proprement (module au lieu de script inline) — inclure
   l'affichage des ALERTES (§7) dans la sortie Excel finale ; supprimer `app.py`/
   `excel_writer.py`/`planning_engine.py` (legacy) une fois `app.py` rebranché dessus
6. Réparer app Streamlit
7. Objectif révisé : 70 différences hors Eloïse/30 mai, dont 52 hors arbitrage manuel du
   6 mai (chiffre d'AVANT §13.16, à recalculer — voir point 0) — considérer ce chiffre
   comme une bonne base de travail plutôt que viser < 10 à tout prix, certaines
   différences étant des choix légitimement équivalents

---

## 13. RÉSOLUTION EN 3 PASSES + RATTRAPAGE DES MANQUES D'HEURES (nouveau, 08/2026)

**Déclenché par 2 remarques de l'utilisatrice sur le planning de septembre 2026** :
1. Le moteur déplaçait parfois un agent de sa section principale sans nécessité réelle
   (ex. mercredi 2 septembre : Macha déplacée du RDC vers l'Adulte alors que d'autres
   options existaient) — "chaise musicale" inutile.
2. L'équité des heures (§7bis) ne surveillait que le DÉPASSEMENT du planning-type,
   jamais le MANQUE. Un agent présent (ex. Marie-France) pouvait ainsi se retrouver
   largement sous son propre volume habituel sans que rien ne le corrige, alors que
   les réguliers doivent absorber en priorité les absences avant les vacataires.

### 13.1 Passage de 2 à 3 passes de résolution

`penalites` (structurel pur) a été scindé en 3 objectifs résolus successivement,
chacun figé avant de passer au suivant (même principe qu'avant, étendu) :

| Passe | Contenu | Rôle |
|-------|---------|------|
| 1 — `penalites` | D_FILL (5000), JEUNESSE (200), CONSÉCUTIF (150) | Couverture des besoins — jamais négociable |
| 2 — `penalites_stabilite` (NOUVEAU) | G1 : agent PT présent non déplacé (100, réduit à 30 sur Adulte/MF si vacataire présent) | Ne déplacer un agent bien placé QUE si structurellement nécessaire — ne peut plus être "battu" par une accumulation de petites préférences (J1/G2/I1) comme avant |
| 3 — `penalites_equite` | G2 (50), J1 (30/70), J3 (25), I1 (20), Vacataire 1 bonus / Vacataire 2 malus (10), équité DÉPASSEMENT hebdo (franchise 3h), équité MANQUE hebdo (NOUVEAU, franchise 0) | Qualité fine du remplacement + répartition équitable, sans jamais dégrader les passes 1 et 2 |

**Pourquoi avoir sorti G1 de la passe 1** : avant, G1 (stabilité) était mélangé dans
la même somme que J1/G2/I1/J3 — un agent pouvait donc être déplacé si la somme des
petites préférences ailleurs (30+50+20...) dépassait par accident le bénéfice de le
laisser en place. En l'isolant dans sa propre passe, la stabilité est désormais
garantie tant qu'elle ne coûte rien sur la couverture des besoins.

**Vérifié sur le cas concret mercredi 2/09** : après correction, Macha est toujours
déplacée du RDC vers l'Adulte à 17h-19h — mais on a vérifié que c'est cette fois une
vraie nécessité : elle est la SEULE personne encore présente et habilitée Adulte à
cette heure (Anne-Françoise finit à 17h15, Chloé à 17h45, Marie-France à 14h, Tiphaine
à 17h30 ce jour-là). Le comportement est donc correct : plus de déplacement "gratuit",
seulement les déplacements réellement obligés.

### 13.2 Rattrapage des manques d'heures (symétrique au dépassement)

Ajout d'un terme `manque_hebdo` strictement symétrique à `depas_hebdo` (§7bis) :
mesure combien un agent reste SOUS son cumul hebdomadaire personnalisé
(`cumul_hebdo_avant + depas_par_agent[a]` négatif), avec **franchise = 0** (contrairement
au dépassement, toléré jusqu'à 3h) — décision utilisatrice : "elle ne peut que faire
plus, pas moins", donc le moindre manque doit être rattrapé dès qu'une marge existe.

- Le "cumul personnalisé" de chaque agent n'inclut déjà QUE les jours où il est
  réellement éligible cette semaine (`agents_eligibles` exclut ses propres jours
  d'absence, mécanisme déjà en place) → un agent absent 2 jours sur 5 n'est jamais
  pénalisé pour ces 2 jours, seulement comparé à son propre planning-type des jours où
  il est effectivement présent. Répond directement à la mise en garde de l'utilisatrice
  du 11/08 (ne pas surcharger les jours restants pour "compenser" une absence).
- Responsables exclus (même liste `agents_equite` que pour le dépassement) — cohérent
  avec "on préserve leur temps".
- Effet de bord recherché et confirmé utile : entre plusieurs agents de section
  équivalente candidats à un même remplacement (ex. Léa/Christine/Delphine/Marie-France
  sur RDC/Adulte), le solveur préfère désormais naturellement celui qui est le plus en
  retard sur ses heures — résout aussi une partie de l'arbitrage "catégorie A" du §10,
  sans mécanisme de départage séparé à coder.
- **Validé sur données réelles (septembre 2026)** : Marie-France apparaissait à 9h vs
  "11,5h PT" dans le diagnostic initial — écart qui semblait confirmer le problème. En
  creusant après coup, ces 11,5h ne tenaient pas compte du fait qu'elle est elle-même en
  congé le mardi matin cette semaine-là (2h30) : son cumul personnalisé réel est donc
  9h, exactement ce qu'elle obtient. Bon signe : le mécanisme donne le bon résultat une
  fois qu'on regarde correctement l'objectif individuel (jours de présence réels), pas
  le planning-type brut.

### 13.3 Plafond quotidien (nouveau, sécurité anti-surcharge)

Contrainte DURE (pas une pénalité) : un régulier (hors vacataires) ne peut jamais
dépasser **420 minutes (7h) dans la même journée**, quelle que soit la pression du
rattrapage hebdomadaire ci-dessus. Valeur par défaut à ajuster si besoin — aucune
donnée réelle de mai/septembre n'a de journée PT dépassant ce seuil, donc aucun risque
de casser une journée déjà valide.

### 13.4 Validation

- Testé sur le fichier réel septembre 2026 (`SEPTEMBRE2026_Preparation_Planning_Mediatheque.xlsx`)
- **0 alerte, 0 jour infaisable sur l'ensemble du mois** après les 3 changements — aucune
  régression de couverture détectée.
- ⚠️ Non testé : effet sur le chiffre de comparaison mai 2026 (`compare_planning.py`).
  Cette session n'avait pas de fichier de référence pour septembre. **À faire dès que
  possible** : relancer `compare_planning.py` sur mai 2026 avec ce moteur pour vérifier
  que ces changements réduisent bien les différences (notamment les buckets "arbitrage
  catégorie A" et "arbitrage Jeunesse" du §10) sans en créer de nouvelles.
- Fichier livré : `planning_engine_cpsat.py` (remplace la version précédente dans le
  projet).


### 13.5 CORRECTIF — régression détectée par l'utilisatrice le 11/08 (Stéphanie/Macha)

**Symptôme signalé** : mardi 1er septembre 10h-12h30, le RDC était donné à Stéphanie
(responsable, section secondaire) alors que Macha (section primaire RDC) était
totalement libre ce créneau — contraire aux deux corrections tout juste faites.

**Cause identifiée** : `penalites_equite` (passe 3) mélangeait deux choses de nature
différente : des poids fixes par créneau (G2=50, J1=30/70, J3=25, I1=20) ET l'écart
d'heures compté directement en MINUTES (1 point par minute au-delà de la franchise).
Résultat : quelques dizaines de minutes de dépassement suffisaient à faire préférer
un remplaçant moins bien assorti (ou une responsable, dont les heures ne sont même pas
comptées) plutôt que d'accepter un léger dépassement chez le bon candidat — l'inverse
de l'intention.

**Correctif** : passage à 4 passes au lieu de 3. `penalites_equite` (passe 3) a été
scindée en :
- `penalites_qualite` (passe 3/4) : G2, J1, J3, I1, préférence vacataires — QUI est le
  meilleur remplaçant, sans regarder les heures.
- `penalites_equite` (passe 4/4, dernière) : dépassement + manque d'heures — ne sert
  plus qu'à départager des choix par ailleurs strictement équivalents en qualité.

**Revérifié** : mardi 1/09 10h-12h30 → RDC = Macha (corrigé). 0 alerte, 0 jour
infaisable sur tout le mois après ce correctif.

⚠️ Cette découverte illustre un risque general à surveiller pour toute future
modification : dès qu'une pénalité est comptée dans une UNITÉ différente (minutes vs
poids fixes par créneau) et mélangée dans la même somme qu'une autre catégorie de
règles, un effet de bord peut apparaître à distance qui n'est pas visible à la lecture
du code, seulement lors de la vérification concrète sur un vrai planning. D'où
l'importance de vérifier plusieurs jours/exemples concrets après chaque changement,
pas seulement le cas qui a motivé le changement.

### 13.6 CORRECTIFS — 3 remarques utilisatrice sur la semaine 1 de septembre (11/08)

**1. Barbara, mercredi 13h-14h (MF)** : la pause déjeuner obligatoire (C3, ≥1h entre
12h-14h) s'appliquait à TOUS les réguliers sans exception, alors que certains agents
(Barbara le mercredi : 8h30-15h sans coupure enregistrée) n'ont réellement aucune
pause à cette période. Corrigé : un agent dont `Horaires_Des_Agents` indique une
présence continue ce jour-là (fm == da, même logique que la pause flexible déjà
existante) est désormais automatiquement exempté de cette pause artificielle.

**2. Vendredi 4/09 — séparation du seuil "idéal" (2h30) et "toléré" (4h)** : la
préférence "éviter plus de 2h30 d'affilée" était rangée au même niveau que la
couverture des besoins (passe 1, avant même la stabilité), ce qui l'a faite
systématiquement l'emporter sur des considérations bien plus importantes (rester en
section primaire, éviter les responsables) — cf. Robin/Macha déplacés au profit de
Stéphanie/Anne-Françoise ce jour-là. Corrigé : deux seuils désormais distincts —
- **Idéal (2h30 semaine / 4h mercredi-samedi)** : dépassement toléré, pénalité (40)
  déplacée en passe 3/4 (`penalites_qualite`) — comparée équitablement à G2/J1/J3 au
  lieu de les écraser automatiquement.
- **Toléré (4h partout désormais, 5h Barbara le samedi)** : reste en passe 1 (poids
  150, comme l'ancien seuil unique) — désormais un vrai second seuil, pas juste un
  renommage. Testé : le rendre "presque dur" (poids 1000) cassait la couverture
  Jeunesse du samedi 5/09 (2/3 agents) → gardé à 150 pour préserver le principe
  "mieux vaut un léger dépassement qu'une alerte".

Revérifié sur la proposition concrète de l'utilisatrice (Marie-France au RDC tout
l'après-midi vendredi, 15h30-19h = 3h30, sous le seuil toléré) : c'est maintenant le
choix retenu spontanément par le moteur, plus besoin de forcer quoi que ce soit.

**3. Jeudi 3/09, Léa vs Anne-Françoise (RDC 17h-19h)** : diagnostiqué comme un
prolongement du même mécanisme que le correctif §13.5 (Stéphanie/Macha) mais plus
subtil — ici la passe qualité (G2/J1/J3) était STRICTEMENT indifférente entre deux
solutions à 145 points chacune (Anne-Françoise seule sur le créneau, OU Léa sur ce
créneau + Anne-Françoise sur le créneau qu'elle libère). Comme Léa avait déjà +2h de
dépassement cette semaine-là (encore sous la franchise de 3h, mais proche), l'équité
(passe 4) évitait de l'alourdir et laissait faire Anne-Françoise — qui, elle, n'est
jamais comptée. Effet indirectement corrigé par le changement du point 2 ci-dessus
(qui redistribue différemment les coûts de la journée) : le résultat après correctif
utilise Guillaume (section secondaire, non-responsable, coût qualité le plus bas —
30 points) plutôt que Léa ou Anne-Françoise. Pas de changement de règle supplémentaire
nécessaire pour l'instant ; à surveiller si le cas se represente ailleurs sous une
forme où aucune 3e option n'existe (Léa vs responsable, sans échappatoire).

**Validation** : 0 alerte, 0 jour infaisable sur tout le mois après ces 3 correctifs.

### 13.7 Colonne "Priorité_remplacement_RDC" dans Affectations (08/2026)

Remplace la règle codée en dur "Adulte > Jeunesse à égalité" par une colonne
directement éditable par l'utilisatrice dans l'onglet **Affectations** :

`Priorité_remplacement_RDC` (nombre entier, plus petit = préféré). Colonne lue par
**en-tête** (pas par position), donc robuste si d'autres colonnes sont ajoutées entre
temps. Utilisée uniquement comme départage FIN (poids 2×valeur, en passe 3/4 qualité)
entre agents déjà à égalité de rang J1 — ne peut jamais l'emporter sur une vraie
différence de section (30 vs 70).

Valeurs actuelles (données par l'utilisatrice, 08/2026) :
- 1 : Léa, Chloé
- 2 : Stéphanie, Robin, Guillaume
- 3 : Tiphaine

Si la colonne est absente ou une cellule vide pour un agent → aucun départage
appliqué pour cet agent (comportement neutre, pas d'erreur).

**Signature modifiée** : `parse_affectations()` retourne désormais 5 valeurs
(ajout de `priorite_rdc`), et `solve_day()` prend un paramètre `priorite_rdc`
supplémentaire. Tout appelant externe (ex: `generate_planning_excel_septembre.py`)
a été mis à jour en conséquence — à vérifier si d'autres scripts appellent
`parse_affectations()` directement.

### 13.8 Règle de travail — toujours revérifier les données sources

Suite à une erreur de raisonnement (mercredi 11/08 : affirmation incorrecte sur la
section principale de Macha, contredite par les données déjà affichées plus tôt dans
la même conversation) : **revérifier systématiquement les données sources
(Affectations, PT, Horaires_Des_Agents...) au moment de chaque affirmation factuelle,
plutôt que de se fier à ce qui a été retenu plus tôt dans la conversation** —
particulièrement important sur ce projet où les arbitrages se jouent à quelques
points de pénalité près.

### 13.9 Incident de process — cache Python obsolète (11/08)

Le fichier livré juste après l'ajout de `Priorité_remplacement_RDC` ne reflétait PAS
la règle (Guillaume au lieu de Chloé, mercredi 10h-11h) alors que la vérification
Python faite juste avant montrait le bon résultat. Cause : un cache de compilation
Python (`__pycache__`) obsolète utilisé par le script générateur au moment de la
génération de l'Excel final, malgré l'édition du fichier source juste avant.
**Consigne ajoutée** : toujours `rm -rf __pycache__` avant toute génération finale
d'un livrable, immédiatement après une modification de `planning_engine_cpsat.py`,
même si une vérification Python juste avant semblait déjà correcte — la vérification
et la génération finale peuvent ne pas partager le même état de cache.

### 13.10 Améliorations cosmétiques du fichier Excel (11/08)

1. **Heure exacte des événements décalés** : si un événement ne correspond pas
   exactement aux bornes du créneau affiché (ex: "Bébé se livre" 10h15-10h45 dans
   le créneau 10h-11h), l'heure exacte est maintenant préfixée dans la cellule
   ("10h15-10h45 Bébé se livre (Tiphaine)"). Fonction `label_evenement()`.
2. **Fusion verticale des cellules identiques consécutives**, sur TOUTES les
   colonnes (RDC/Adulte/M&F/Jeunesse/Accueil/Animation/Réunion/Absence), quand le
   même contenu se répète sur des créneaux d'affilée d'une même journée. Fonction
   `fusionner_cellules_identiques()`, appelée en fin de bloc de chaque journée.
   Les valeurs vides/'—' ne sont jamais fusionnées.
3. **Calcul d'heures protégé de la fusion** : ajout de 4 colonnes techniques
   cachées L/M/N/O (copies systématiques, jamais fusionnées, de B/C/D/E). Le récap
   d'heures dynamique lit désormais L-O au lieu de B-E — la fusion visuelle de
   B-E n'a donc aucun effet sur le calcul (vérifié : Christine, mercredi 30/09,
   B15:B16 fusionnées visuellement, mais L15/L16 gardent chacune "Christine" →
   ses heures sont toujours comptées correctement, recalcul indépendant = 13,0h
   semaine 5).
4. **Couleur unique** pour Accueil/Animation/Réunion/Absence (FFE0E0EC / FFE8E8F2),
   distincte des 4 couleurs de section RDC/Adulte/M&F/Jeunesse.
5. **Créneaux sans service public discrets** (mardi/jeudi/vendredi uniquement) :
   police taille 8, italique, gris (FF999999) au lieu du style normal.

Script concerné : `generate_planning_excel_septembre.py` — à répliquer sur le
générateur des mois suivants une fois nommé/dupliqué.

### 13.11 CORRECTIF — récap d'heures redevenu statique (11/08)

La 1re version des colonnes techniques L-O (§13.10) écrivait des VALEURS figées au
moment de la génération, pas des formules — donc une modification manuelle d'un
agent en B-E dans Excel n'était plus répercutée dans le récap d'heures. Corrigé :
L/M/N/O sont maintenant des formules `=IF(B{r}<>"",B{r},L{r-1})` (recopie la valeur
du dessus uniquement quand la cellule visible est vide, càd uniquement en cas de
fusion — jamais sur la 1re ligne d'un jour). Revérifié avec un recalcul réel
(LibreOffice headless, pas juste une lecture de formule) : modifier l'agent dans une
cellule fusionnée (B15:B16, Christine→Macha) répercute correctement -2h/+2h dans le
récap. Le calcul reste donc juste ET dynamique, y compris avec les cellules fusionnées.

### 13.12 Vue par agent (nouvel onglet Semaine_X_Agent, 08/2026)

Ajout d'un onglet `Semaine_X_Agent` juste après chaque `Semaine_X`, sur demande
utilisatrice : un planning par agent (blocs empilés verticalement, un par agent),
colonnes = jours de la semaine, lignes = grille horaire fine (union des bornes de
créneaux de tous les jours, pour aligner mardi/jeudi/vendredi — créneaux larges —
avec mercredi/samedi — créneaux fins). Chaque cellule est une **formule** (pas une
valeur figée) qui cherche le nom de l'agent dans les colonnes techniques cachées
L/M/N/O de l'onglet `Semaine_X` correspondant et affiche la section trouvée
(RDC/Adulte/M&F/Jeunesse) ou rien. Entièrement dynamique et vérifié avec un vrai
recalcul (LibreOffice) : modifier un agent dans le planning global répercute
immédiatement la vue par agent, dans les deux sens (l'agent retiré disparaît de son
ancien créneau, le nouvel agent apparaît sur le même créneau).

Fonctions ajoutées : `grille_fine_commune()` (construit la grille horaire unifiée
de la semaine) et `generer_vue_agent()` (construit l'onglet, une formule
`IF(ISNUMBER(SEARCH(...)))` en cascade par cellule, une par section possible).

### 13.13 Cosmétique — gras/italique (08/2026)

- Noms d'agents dans RDC/Adulte/M&F/Jeunesse : désormais en **gras** (en plus de
  la couleur par agent).
- Contenu des colonnes Accueil/Animation/Réunion/Absence : désormais en *italique*.

### 13.14 Ajustements vue par agent + retrait fusion B-E (11/08)

- **Fusion B-E retirée** (RDC/Adulte/M&F/Jeunesse) sur le planning global : risque
  identifié par l'utilisatrice — défusionner une cellule dans Excel vide les cellules
  non-ancres SANS les marquer comme "à remplir", et la formule des colonnes cachées
  L-O (`=IF(B{r}<>"",B{r},L{r-1})`) recopierait alors silencieusement l'ancien agent
  → calcul d'heures ET vue par agent faux, sans avertissement. Fusion conservée
  uniquement sur Accueil/Animation/Réunion/Absence (F-I), où une erreur de ce type
  n'a aucune conséquence sur un calcul.
- **Vue par agent** : en-tête (jours) désormais figé en haut (`freeze_panes='B2'`),
  ne se répète plus par agent — plus compact. Fond de cellule coloré selon la
  section trouvée (mêmes couleurs que le planning global : bleu clair=RDC, vert=
  Adulte, jaune=M&F, rose=Jeunesse), via mise en forme conditionnelle Excel
  (`openpyxl.formatting.rule.FormulaRule`, une règle par section, s'applique sur
  toute la grille) — texte noir gras (plus de couleur par agent ici, la couleur de
  fond suffit à se repérer). Bandes alternées gris très clair / blanc entre chaque
  bloc agent pour mieux les séparer visuellement.

### 13.15 CORRECTIFS — couleurs invisibles + fuite des événements (11/08)

**1. Couleurs de section invisibles dans Excel (bug dxf)** : la mise en forme
conditionnelle utilisait `fgColor` pour le remplissage, qui fonctionne pour un
remplissage NORMAL mais est ignoré par Excel pour un remplissage de mise en forme
CONDITIONNELLE (dxf) — Excel utilise `bgColor` dans ce cas précis. LibreOffice est
plus tolérant et affichait quand même la couleur, ce qui a caché le problème lors
des vérifications précédentes. Corrigé : `bgColor` renseigné en plus de `fgColor`
sur tous les `PatternFill` utilisés dans des `FormulaRule`.

**2. Événements qui "fuyaient" sur les heures suivantes** : les colonnes cachées
P/Q/R (miroirs de Accueil/Animation/Réunion, ajoutées pour la vue par agent)
utilisaient la même logique "si vide, recopier la ligne du dessus" que L-O
(RDC/Adulte/M&F/Jeunesse). Cette logique suppose qu'une cellule vide = continuation
d'une fusion — vrai pour L-O (toujours remplies sur un créneau ouvert) mais FAUX
pour les événements (vides la plupart du temps) : une réunion de 10h-11h30 se
retrouvait recopiée jusqu'à 19h. Corrigé en profondeur : `fusionner_cellules_identiques()`
prend maintenant un paramètre `hidden_map` et écrit les formules des colonnes
cachées en fonction des fusions RÉELLEMENT décidées (référence directe à la ligne
du haut de chaque fusion, plus jamais de supposition). B-E n'étant plus fusionnées
(cf. §13.14), leurs colonnes cachées L-O sont redevenues de simples références
directes (`=B{r}`, etc.) — plus besoin de logique de repli du tout.

**Vérifié** : Réunion pôle (mercredi 2/09, 10h-11h30) n'apparaît plus que sur ses
2 créneaux réels dans la vue par agent (avant : fuyait jusqu'à 19h). Calcul d'heures
toujours juste (Christine semaine 5 = 13h, inchangé).

### 13.16 ÉQUITÉ — les événements comptent comme du service public (12/08)

**Règle utilisatrice** : un agent occupé par un événement (accueil de classe, animation,
réunion) n'est pas "disponible" pour autant — il donne une heure de travail face au public
tout aussi légitime qu'une heure de comptoir. Jusqu'ici, l'équité (§7/§7bis) ne comptait
QUE les heures de comptoir par rapport au planning-type ; un agent très sollicité en
événements avait donc l'air "sous son quota" et pouvait être choisi en priorité comme
remplaçant, alors qu'il était déjà chargé. Équivalence retenue, volontairement simple :
**1h événement = 1h service public, quel que soit le type d'événement** (Accueil/
Animation/Réunion) — pas de pondération différenciée entre types.

**Implémentation** : dans `solve_day`, le dépassement net par agent (`depas_par_agent`,
utilisé pour les 2 niveaux d'équité jour + semaine, §7bis) inclut désormais les minutes
d'événements du jour de l'agent :
`dépassement = (service public réel + minutes événements du jour) − service public prévu au PT`.
Comme ce chiffre est celui repris tel quel pour l'équité hebdomadaire (`cumul_hebdo`),
les événements comptent aussi bien pour la franchise journalière (60 min) que pour la
franchise hebdomadaire (180 min) — aucun autre changement d'architecture, toujours en
tout dernier (passe 4/4), toujours un simple critère de départage qui ne peut jamais
empêcher un remplacement réellement nécessaire.

⚠️ **Piège évité, important pour la suite** : l'onglet Événements contient AUSSI les
congés/RTT/formations/absences (ex. "congé" 9h-19h), pas seulement de vrais événements
travaillés. Si on les avait comptés, un agent en congé toute la journée aurait été crédité
de ~10h de "travail équivalent" — l'inverse de l'intention. Corrigé via une nouvelle
fonction `_est_evenement_absence(nom)` qui exclut les entrées dont le nom contient
congé/RTT/vacation/absence/formation (mêmes mots-clés que la colonne "Absence" du
générateur Excel) avant de calculer les minutes d'événements de chaque agent. Seuls les
vrais événements (accueil, animation, réunion...) sont comptés.

**Vérifié (12/08)** : testé sur le fichier réel de septembre 2026 — 0 alerte, 0 jour
infaisable, aucune régression. Vérifié manuellement que les vraies réunions/animations
sont bien comptées et que les congés en sont bien exclus (`Reunion pôle`, `Portage`,
`forum des associations`... comptés ; `congé` exclu).

⚠️ **Non testé** : impact sur le chiffre de comparaison mai 2026 (`compare_planning.py`).
À faire en priorité à la prochaine session (voir §12) — cette règle est justement
susceptible de réduire les buckets "arbitrage catégorie A" et "arbitrage Jeunesse" (~30
différences au total, §10), puisque c'est précisément le type de départage qui manquait
entre agents par ailleurs équivalents.

**Fichier livré** : `planning_engine_cpsat.py` (remplace la version précédente dans le
projet).

### 13.17 COSMÉTIQUE — bandeau nom d'agent illisible (12/08)

**Symptôme** : dans les onglets `Semaine_X_Agent` (vue par agent), la bannière affichant
le nom de l'agent était en fond bleu très foncé (`2C3E50`) avec le texte dans la couleur
propre à l'agent (`AGENT_COLORS`) — illisible pour les agents dont la couleur de police
est elle-même sombre (ex. Robin `002060`, bleu foncé sur bleu foncé).

**Correctif** : fond de la bannière passé en gris clair (`D9D9D9`), texte inchangé
(toujours coloré par agent, gras) — contraste largement suffisant sur fond clair pour
toutes les couleurs de la palette `AGENT_COLORS`.

**Fichier livré** : `generate_planning_excel_septembre.py` (remplace la version
précédente dans le projet). Changement ciblé uniquement (pas de régénération complète du
fichier Excel demandée par l'utilisatrice) — à vérifier visuellement à la prochaine
génération complète.

### 13.18 REFONTE PRÉSENTATION EXCEL — essais itératifs sur Semaine_1 (08/2026)

**Contexte** : longue session d'essais de présentation, volontairement limités à
`Semaine_1` + `Semaine_1_Agent` pour itérer vite avant validation sur le planning
complet (5 semaines). Résultat : `generate_planning_excel_septembre.py` mis à jour
avec tous les changements validés ci-dessous, testé avec succès sur les 5 semaines
(4561 formules, 0 erreur) avant livraison — mais le fichier livré à l'utilisatrice
reste volontairement l'essai Semaine_1 tant que la présentation n'est pas validée
en totalité.

**1. Fond coloré par agent (planning global ET vue par agent)**

Remplace le fond par section (bleu=RDC, vert=Adulte, etc.). Palette
`AGENT_FILL_COLORS` reprise **telle quelle** d'une capture d'écran fournie par Elo
(couleurs qu'elle utilise déjà à la main) : Delphine, Stéphanie, Christine,
Guillaume, Macha, Stéphane, Tiphaine, Chloé, Robin, Anne-Françoise, Marie-France,
Agnès. Léa et Barbara n'apparaissaient pas dans sa capture → couleurs provisoires
(or clair / mauve clair) à valider. Vacataires nommément identifiés dans le
planning (Clara-Jade, Vacataire Marie, en plus des Vacataire 1/2/3 génériques) :
même gris neutre pour tous, pas de couleur individuelle (demande explicite —
ce sont des vacataires, pas la peine de les distinguer).

Couleur de texte choisie automatiquement par contraste (`_texte_lisible()` — noir
sur fond clair, blanc sur fond foncé selon la luminance), plutôt que noir fixe :
certaines couleurs de la capture (bleu Marie-France, violet Tiphaine...) sont trop
sombres pour du texte noir.

**En-têtes de colonnes** : fond gris uniforme (`FFCCCCCC`, même gris que
"Créneau") sur toutes les colonnes — fini les couleurs différentes par section
sur la ligne d'en-tête (le fond par agent, dans le contenu, suffit à se repérer).

**2. Jeunesse éclatée en 3 colonnes, Accueil+Animation fusionnées**

Jeunesse peut avoir jusqu'à 3 agents sur un même créneau (vérifié : jamais le cas
pour RDC/Adulte/M&F, qui restent à 1 seule colonne chacune) — avec un fond par
agent, plusieurs agents dans une même case posait un problème de lisibilité
(rayures illisibles). Solution : 3 colonnes `Jeunesse 1` / `Jeunesse 2` /
`Jeunesse 3`, un agent par colonne, fond uni. En contrepartie, `Accueil` et
`Animation` fusionnées en une seule colonne `Accueil / Animation` (même logique
de détection de mots-clés qu'avant, juste regroupée — `classer_evenement()` ne
retourne plus que 2 catégories : `Réunion` ou `Accueil / Animation`).

**Fusion des 3 colonnes Jeunesse — ATTENTION, changement en 2 temps** :
- 1er essai : fusion verticale du CONTENU des 3 colonnes quand le même agent
  enchaîne plusieurs créneaux (comme Accueil/Animation/Réunion/Absence).
- **Version retenue (dernière demande utilisatrice)** : ce n'est PAS ce qui était
  voulu. Le contenu des 3 colonnes Jeunesse reste **non fusionné** (une ligne par
  créneau, toujours). Seule la **ligne d'en-tête** fusionne les 3 cases en une
  seule case "Jeunesse" (`ws.merge_cells` sur la ligne d'en-tête uniquement,
  colonnes E:G), pour indiquer visuellement que ce sont 3 sous-colonnes d'une même
  section — sans toucher aux données. `fusionner_cellules_identiques()` ne
  s'applique donc plus qu'aux colonnes Accueil/Animation, Réunion, Absence
  (colonnes 8-10), comme avant l'introduction de Jeunesse 1/2/3.

Colonnes cachées mises à jour en conséquence : L=RDC, M=Adulte, N=M&F,
O/P/Q=Jeunesse 1/2/3, R=Accueil/Animation, S=Réunion (Absence n'a jamais eu de
colonne cachée, pas besoin pour la vue par agent). `RECAP_SOURCE_COLS` couvre
maintenant B→L, C→M, D→N, E→O, F→P, G→Q (6 colonnes sources pour le récap heures,
au lieu de 4).

**3. Colonne A (Créneau) figée**

`ws.freeze_panes = 'B1'` sur l'onglet `Semaine_X` (existait déjà sur la vue par
agent mais avait été oublié sur le planning global).

**4. Format d'affichage des événements à horaires/agents non standards —
RÈGLE DÉSORMAIS DOCUMENTÉE**

Quand un événement ne correspond pas pile au créneau affiché et/ou concerne des
agents précis, les précisions (agents concernés, horaires exacts) s'affichent
**entre parenthèses**, à la suite du nom : `"Bébés se livrent (10h15-10h45)"`,
`"Réunion pôle (Anne-Françoise, Stéphanie)"`, ou les deux combinés
`"forum des association (Robin, 14h-15h)"`. Cette règle n'était pas documentée
avant cette session — c'est fait (`label_evenement()`, `generate_planning_excel_septembre.py`).
Anciennement, un préfixe était utilisé (`"10h15-10h45 Bébés se livrent"`) —
abandonné au profit des parenthèses.

**5. BUG CORRIGÉ — "forum des association" (Robin) affiché sur 14h-19h au lieu de
14h-15h**

Dans le fichier de préparation de septembre, l'événement est enregistré avec une
fin à 18h (`ce=1080`) alors que Robin n'y est en réalité que de 14h à 15h.
**⚠️ À corriger à la source** : onglet Événements du fichier de préparation,
colonne fin de l'événement "forum des association" du 5 septembre → 15:00 au lieu
de 18:00. Un correctif ponctuel (bornage de l'événement à 15h00 en mémoire) a été
appliqué dans le script d'essai Semaine_1 pour valider le rendu correct, mais
**n'a pas été repris dans le script de production** `generate_planning_excel_septembre.py`
(daté, spécifique à ce mois) — la vraie correction doit se faire dans le fichier
source.

**6. Vue par agent (`Semaine_X_Agent`) — plusieurs ajustements**

- **Grille étendue à partir de 8h** (2 créneaux fixes 8h-9h et 9h-10h ajoutés
  avant les créneaux réels du planning, pour voir les horaires d'arrivée même
  avant l'ouverture au public).
- **Créneaux de plus d'1h découpés en blocs d'1h** (ex : 17h-19h → 17h-18h +
  18h-19h), pour une lecture plus fine des horaires.
- **Arrivées/départs décalés** : quand l'horaire contractuel réel d'un agent
  (`Horaires_Des_Agents`) tombe en plein milieu d'un créneau plutôt que sur son
  bord, la cellule affiche `"Arrivée 8h30"` ou `"Départ 17h15"` à la place du
  contenu habituel — sur le fond habituel de l'agent (pas de couleur dédiée,
  seul le texte en gras signale la particularité).
- **Congé** : toutes les cases d'un agent en congé sont grisées avec la mention
  "Congé", sur toute la durée réelle du congé (journée complète ou partielle) —
  calculé depuis l'onglet Événements (congés), indépendamment du planning.
- **Un seul code visuel pour "pas au travail"** : hachures grises partout
  (`HATCH_FILL`), qu'il s'agisse d'un horaire personnel hors contrat ou d'un
  créneau où la médiathèque n'est pas ouverte. L'ancien gris à tiret (`—`,
  `FERME_FILL`) a été retiré — un seul code à maintenir, plus simple.
  Cas "agent censé être là mais créneau non suivi par le planning" (ex:
  préparation avant l'ouverture) : cellule vide, fond de la couleur de l'agent,
  sans hachure (distinct de "pas au travail").
- **Texte réel de l'événement, pas la catégorie** : la formule affiche
  maintenant le contenu réel de la cellule source (ex: "Prépa portage") plutôt
  qu'un libellé générique "Accueil / Animation". ⚠️ Limite connue : ça ne
  fonctionne que si le nom de l'agent est explicitement écrit dans le texte de
  l'événement (ex: "forum des association (Robin, Eloïse)") — un événement qui
  ne cite aucun nom (comme "Prépa portage") n'apparaîtra dans la vue d'aucun
  agent en particulier, même si tout le monde y participe potentiellement.
- **Prénom de l'agent retiré de son propre texte d'événement** : dans le tableau
  de Marie-France, "Reunion pôle (Anne-Françoise, Stéphanie, Delphine,
  **Marie-France**, Eloïse)" devient "Reunion pôle (Anne-Françoise, Stéphanie,
  Delphine, Eloïse)" — inutile de se citer soi-même dans son propre tableau.
  Implémenté via une chaîne de `SUBSTITUTE()` Excel (retire "Nom, ", ", Nom", ou
  "Nom" seul, puis nettoie les parenthèses vides résiduelles).

**7. Fichier de référence Semaine_1 (usage ponctuel, PAS repris en production)**

Une fonction `charger_reference_solution()` a été écrite pour relire un fichier
Excel déjà validé par Elo (`Planning_Septembre_2026.xlsx`, ancien format à 1 seule
colonne Jeunesse) et remplacer les affectations RDC/Adulte/M&F/Jeunesse calculées
par le solveur CP-SAT par celles de la référence, le temps de valider que le rendu
Semaine_1 correspond exactement à ce qu'Elo attend (vérifié : 0 écart). Cette
fonction et son usage sont **conservés uniquement dans le script d'essai**
(disponible si besoin de refaire un exercice de recalage similaire un autre mois)
et **retirés du script de production** `generate_planning_excel_septembre.py`, qui
continue de générer directement depuis la solution du solveur CP-SAT (comportement
normal, pas de fichier de référence à fournir chaque mois).

**Fichiers livrés cette session** :
- `generate_planning_excel_septembre.py` (production, 5 semaines, testé 0 erreur)
- `Essai_Presentation_S1.xlsx` (Semaine_1 seule, dernière version validée)
- `context_projet_mediatheque_v27.md` (ce fichier)
- `planning_engine_cpsat.py` **inchangé** cette session (aucune modification du
  moteur de résolution nécessaire — uniquement des changements de présentation
  Excel).

### 13.19 CORRECTIF — vue par agent : aucun prénom ne doit apparaître (08/2026)

**Symptôme** : `label_evenement()` (point 4/§13.18) inclut les prénoms des agents
concernés entre parenthèses (ex: `"Reunion pôle (Anne-Françoise, Stéphanie,
Delphine, Marie-France, Eloïse)"`). Repris tel quel dans la vue par agent via les
colonnes cachées R/S, le correctif §13.18-6 (retrait du SEUL prénom de l'agent
courant via `SUBSTITUTE`) ne suffisait pas : les prénoms des AUTRES agents
restaient visibles dans le tableau de chacun.

**Règle utilisatrice, plus stricte** : dans la vue par agent, **aucun prénom,
jamais** — ni celui de l'agent concerné, ni ceux des autres. Seul le nom de
l'événement, complété de l'horaire exact entre parenthèses si l'événement ne
correspond pas pile au créneau affiché (ex: `"Réunion pôle (10h-11h30)"`).

**Solution retenue** : plutôt que d'essayer de retirer les prénoms d'un texte
libre par formule Excel (fragile), calcul de DEUX libellés séparés, dès la
génération Python, à partir des mêmes données structurées (`ev['agents']`,
horaires) :
- `label_evenement()` (inchangé) → texte complet avec prénoms, pour le planning
  global (`Semaine_X`, colonnes Accueil/Animation et Réunion).
- `label_evenement_sans_noms()` (nouveau) → jamais de prénom, uniquement le nom
  de l'événement + horaire si besoin, pour la vue par agent.

Deux nouvelles colonnes cachées **T** (Accueil/Animation sans prénoms) et **U**
(Réunion sans prénoms) ajoutées aux onglets `Semaine_X`, écrites en valeur directe
(pas une formule miroir, contrairement à L-S) puisque le texte diffère
structurellement de H/I. La vue par agent détecte toujours si l'agent est
concerné via R/S (texte complet, recherche du prénom) mais **affiche** désormais
T/U (texte sans prénom) : `IF(ISNUMBER(SEARCH(prénom, R)), T, ...)`. Le mécanisme
`SUBSTITUTE` de retrait du seul prénom courant (§13.18-6) est retiré, devenu
inutile.

**Limite connue** : pour la Semaine_1 (mécanisme de référence, §13.18-7, non
repris en production), la version "sans noms" d'un événement récupéré depuis le
fichier de référence est obtenue par un repli générique (tout ce qui précède la
1ère parenthèse) plutôt que par reconstruction structurée — un éventuel horaire
mentionné dans cette même parenthèse n'est alors pas repris. N'affecte que ce
mécanisme d'essai, pas la production.

**Vérifié** : "Reunion pôle" (Semaine_1, mercredi 2/09) s'affiche identiquement
sans aucun prénom dans les tableaux d'Anne-Françoise, Stéphanie et Marie-France ;
"forum des association" (Robin, samedi, créneau pile 14h-15h) sans horaire
redondant puisqu'il correspond exactement au créneau affiché.

**Fichiers livrés** : `generate_planning_excel_septembre.py` (testé sur les 5
semaines, 4561 formules, 0 erreur), `Essai_Presentation_S1.xlsx` (dernière
version), `context_projet_mediatheque_v27.md` (ce fichier, mis à jour).

### 13.20 CORRECTIF — pause déjeuner invisible pour les agents "pause flexible" (08/2026)

**Symptôme signalé par Elo** : dans la vue par agent, Macha et Anne-Françoise
n'ont jamais de pause déjeuner hachurée — parfois une case affiche "Arrivée
13h30" alors qu'aucune pause n'a été indiquée juste avant.

**Cause identifiée** : Macha, Anne-Françoise, Christine et Delphine sont marquées
"pause flexible" (colonne dédiée de l'onglet Affectations, `pause_flex`). Cette
notion existe pour le **solveur** (`planning_engine_cpsat.py`) : elle
l'**autorise**, si besoin, à placer exceptionnellement un de ces agents sur son
créneau de pause nominal — mais ça ne veut pas dire que l'agent n'a pas de pause.
`_dans_horaires_agent()` (fonction d'AFFICHAGE uniquement, dans le générateur
Excel — aucun rapport avec le moteur de résolution) reprenait cette même règle et
en tirait la mauvaise conclusion : pour un agent "pause flexible", elle
considérait TOUTE la journée (`dm` à `fa`) comme "en horaires", sans jamais
hachurer l'écart entre `fm` et `da` (la pause nominale lue dans
`Horaires_Des_Agents`). Le libellé "Arrivée/Départ" (§13.18-6, indépendant de
cette fonction) continuait lui à se déclencher normalement sur le bord `da`
— d'où l'incohérence visuelle (pas de hachure, mais une "Arrivée" au milieu de
rien).

**Vérifié avant correctif** : sur la Semaine_1, ni Macha ni Anne-Françoise ne sont
jamais réellement affectées à un rayon pendant leur créneau de pause nominal
(recherché dans la solution du solveur) — corriger l'affichage ne risque donc pas
de masquer une vraie affectation cette semaine-là. À revérifier si un jour le
solveur utilise effectivement cette flexibilité pour un agent donné (l'agent
apparaîtrait alors normalement dans son rayon malgré la case marquée pause,
puisque l'affichage réagit toujours en premier lieu à ce que contient le
planning réel — seul le cas "rien n'est prévu à cet endroit" bascule en pause).

**Correctif** : `_dans_horaires_agent()` ne fait plus de cas particulier pour les
agents "pause flexible" — la pause nominale (`fm` à `da`) est désormais toujours
hachurée, pour tout le monde, cohérent avec la règle utilisatrice ("tous les
agents ont une heure de pause sauf s'ils terminent tôt, à 14h ou 15h"). Le
paramètre `pause_flex` de la fonction est conservé dans la signature (pour ne pas
casser les appels) mais n'est plus utilisé. **Aucun changement côté moteur**
(`planning_engine_cpsat.py`) — la "pause flexible" garde tout son sens pour le
solveur, seul l'affichage Excel change.

**Vérifié après correctif** : Macha (mardi, pause 12h30-13h30) → case
12h30-13h00 hachurée, case 13h00-14h00 affiche "Arrivée 13h30" (cohérent).
Testé sur les 5 semaines, 0 erreur de formule.

**Fichiers livrés** : `generate_planning_excel_septembre.py`,
`Essai_Presentation_S1.xlsx`, `context_projet_mediatheque_v27.md`.

## 14. ZONE DE NOTES AGENTS — saisie au fil de l'eau (09/2026, demande utilisatrice)

### Besoin

Le fonctionnement réel de l'équipe est itératif : les agents notent au fil de
l'eau ce qui change (réunion imprévue, absence, départ anticipé...), et la
personne en charge du planning ajuste. L'ancien modèle artisanal (fichier
`03_Mars.xlsx`, fourni en référence) avait, à droite de chaque journée, un
petit tableau "Nom / Événement" pour ça. Le nouveau planning ne l'avait pas
— ajouté ici, avec cascade automatique.

### Où c'est dans le fichier

Pour **chaque journée** (pas un seul tableau par semaine — un tableau par
jour, comme dans l'ancien modèle), deux paires de colonnes juste après les
colonnes cachées existantes :
- **W (Nom) / X (Événement)** — 7 premiers agents réguliers (hors vacataires)
- **Y (Nom) / Z (Événement)** — 7 agents suivants

Les noms (colonnes W et Y) sont **pré-remplis** par le générateur, avec le
même fond coloré que dans le planning principal (`AGENT_FILL_COLORS`). Seule
la colonne Événement est à remplir par l'agent. En-tête de cette colonne :
*"Événements à ajouter (ex. format : 14h-15h Accueil classe)"*.

### Format attendu dans la case Événement

- **L'heure, si donnée, se met TOUJOURS en premier**, suivie d'un espace,
  puis le texte libre. Ex : `14h-15h accueil de classe`, `17h30 part`.
  Une heure qui n'est pas en tout premier n'est pas reconnue (l'événement
  est alors traité comme "toute la journée" — comportement volontaire,
  protège contre les oublis de format plutôt que de deviner).
- **Mots-clés de classement** (accents/majuscules ignorés, un seul mot
  suffit, n'importe où dans le texte) :
  - Réunion : réunion, reunion, rdv
  - Absence : congé, conge, absent, part
  - Accueil / Animation : tout le reste (par défaut)
- Une note ne remplace jamais ce qui existe déjà dans la case — elle
  s'ajoute à la suite (séparateur `; `).

### Mécanique technique (formules Excel, pas de macro)

1. Colonnes cachées d'analyse par case Événement (catégorie / texte nettoyé
   de l'horaire / heure de début-fin en décimal) — mêmes principes que
   `COL_DUREE`, juste appliqués à la zone de notes.
2. Une formule matricielle (`TEXTJOIN`+`IF`, via `ArrayFormula`) à l'ancre de
   chaque bloc fusionné de H (Accueil/Animation), I (Réunion) et J (Absence)
   combine le texte déjà généré avec les nouvelles notes qui correspondent
   (bonne catégorie + bon créneau horaire).
3. Les colonnes cachées T et U (versions sans prénom, utilisées par la vue
   par agent) sont reconstruites selon le même principe, à partir de H et I.
4. Les colonnes cachées R et S (miroirs de H et I) sont réécrites ligne par
   ligne pour pointer vers l'ANCRE RÉELLE de leur bloc — corrige un défaut de
   `fusionner_cellules_identiques()` : une suite de lignes vides
   consécutives, bien que non fusionnées visuellement, étaient jusqu'ici
   traitées comme un seul bloc logique (toutes pointant vers la première
   ligne de la suite). Sans gravité tant que rien n'est écrit dedans, mais
   ça cassait la remontée d'une nouvelle note ajoutée sur une ligne du
   milieu. Correction appliquée de façon générale (pas seulement là où une
   note existe), aucune régression attendue sur l'affichage existant.

### Limite connue : l'Absence ne remonte PAS dans la vue par agent

La vue par agent (`Semaine_N_Agent`) affiche déjà "Congé" / "Arrivée Xh" /
"Départ Xh" — mais ce texte est calculé UNE FOIS par le générateur Python à
partir des événements du fichier de préparation (`conge_par_agent_jour`,
§ dans `generer_vue_agent`), pas par une formule qui lit la colonne J. Une
note d'absence tapée dans la nouvelle zone remonte donc bien dans le planning
principal (colonne J) mais **pas** dans la vue par agent. Les notes Réunion
et Accueil/Animation, elles, remontent bien aux deux endroits (via T/U, qui
sont des formules).
→ Pas corrigé pour l'instant (portée limitée à ce qui a été demandé) — à
faire si besoin, sur demande explicite.

### Verrouillage des cellules à formule (09/2026, demande utilisatrice)

À la toute fin de `generer()`, `verrouiller_cellules_formules(wb)` parcourt
tous les onglets et verrouille (`Protection(locked=True)`) toute cellule dont
la valeur est une formule (`str` commençant par `=`, ou `ArrayFormula`) —
détection cellule par cellule, pas une liste de colonnes à part. Toutes les
autres cellules (dont la nouvelle colonne Événement, et les cases
d'affectation RDC/Adulte/M&F/Jeunesse, qui sont du texte simple) restent
éditables. Protection de feuille activée (`ws.protection.sheet = True`), sans
mot de passe — objectif : éviter les écrasements accidentels en tapant
directement dans Excel, pas empêcher un déverrouillage volontaire (Excel →
Révision → Ôter la protection de la feuille, aucun mot de passe à saisir).

### Fichiers livrés

`generate_planning_excel_septembre.py` (toutes les modifications ci-dessus
intégrées, s'appliquent automatiquement aux 5 semaines à chaque génération),
`Planning_Semaine1_avec_notes_agents.xlsx` (exemple, 2 onglets seulement —
Semaine_1 + Semaine_1_Agent, pas le classeur complet), ce fichier de contexte.

---

## 15. GÉNÉRATION AUTOMATIQUE DE L'ONGLET ÉVÉNEMENTS DEPUIS LES FICHIERS SOURCES BRUTS (session 14/08)

⚠️ **Plusieurs descriptions ci-dessous se sont révélées fausses une fois
testées sur les VRAIS fichiers d'Elo (session 18/08)** — écrites le 14/08
sans exemple réel sous la main, elles décrivaient une structure plausible
mais pas celle réellement utilisée. **Voir §17 pour la structure réelle
corrigée et validée.** Gardé ci-dessous tel quel pour l'historique de la
décision (règles métier, principe "sources brutes font foi"), mais pour
toute question de colonnes/format, se fier uniquement au §17.

### Contexte / demande utilisatrice

Jusqu'ici, Elo compilait à la main, chaque mois, un onglet "Événements" à
partir de plusieurs fichiers sources qu'elle tient déjà (congés, accueils
crèche, accueils de classe, lectures du jeudi matin...). Demande : que
l'assistant lise directement ces fichiers sources et construise lui-même
l'onglet Événements, au même format que celui attendu par
`planning_engine.parse_evenements` (Date texte FR | Début | Fin | Nom |
Agents `;`-séparés).

**Nouveau fichier livré : `sources_to_evenements.py`** — module autonome,
pas encore branché sur `planning_engine.py` / l'app Streamlit (prochaine
étape, cf. §15 "Sur l'horizon"). Contient un parseur dédié par type de
fichier source, plus une fonction d'écriture de l'onglet Excel final.

### Sources gérées et règles de lecture (validées avec Elo)

**1. Congés équipe** (`parse_conges`) — un onglet par mois (nom = mois en
français capitalisé : "Mai", "Septembre"...). Ligne d'en-tête repérée par la
cellule "Nom de l'employé" en colonne B ; les numéros de jour sont sur cette
même ligne, à partir de la colonne C. **Règle validée : n'importe quelle
lettre dans une case (C, CS, M, récup...) = congé journée complète (9h-19h),
peu importe le motif.** Une valeur numérique < 1 (ex: 0,5) = demi-journée ;
comme le fichier ne dit pas si c'est le matin ou l'après-midi, l'heure est
laissée vide et la case est **surlignée en jaune avec un commentaire**, pour
qu'Elo complète à la main. Lydie et Eloïse, si elles apparaissent dans ce
fichier, sont ignorées silencieusement (cf. règle générale : Lydie a quitté,
Eloïse n'est jamais planifiée automatiquement).

**2. Accueil crèches** (`parse_accueil_creche`) — cellule A1 = heure de
l'accueil (ex: "10h-11h"), colonne A = mois, colonne B = jour ("jeudi 7"),
colonne C = nom de la crèche (un événement est créé seulement si cette
colonne est remplie). Intervenante toujours Tiphaine. **Si A1 est vide (cas
rencontré sur les deux fichiers crèche testés), l'heure est laissée vide et
la case est surlignée en jaune** — à vérifier/compléter par Elo, ou à
corriger à la source si l'heure devrait toujours y être.

**3. Accueil de classe** (`parse_accueil_classe`) — colonne A = date,
colonne B = initiales de l'intervenant (SD=Stéphanie, TV=Tiphaine,
GC=Guillaume, RL=Robin, BP=Barbara, DR=Delphine, EG=Eloïse), colonne D =
créneau horaire, colonne E = nom de l'école (événement créé seulement si
rempli). **Si aucune initiale n'est indiquée en colonne B, la case Agents est
laissée vide et surlignée en jaune** (règle validée par Elo le 14/08).

**4. Lecture du jeudi matin** (`parse_lecture_jeudi_matin`) — cellule A1 =
heure fixe de la séance (ex: "10h-10h30"). Une séance occupe plusieurs
lignes (une ligne par enfant) ; seule la 1re ligne de la séance porte la date
+ code groupe en colonne B (ex: "06/11/2025 IM Gp 1"), et la ligne "Total" de
fin de séance porte l'intervenant(e) en colonne I — en texte libre ("Agnès et
Tiphaine", "Agnès Stéphanie"...), reconnu par recherche des prénoms d'agents
connus dans le texte. **Nom de l'événement toujours fixe : "lectures AM/AP"**
(validé par Elo — ce n'est pas une donnée du fichier). Si le texte de la
colonne I ne correspond à aucun agent connu (ex: séance annulée), l'événement
est quand même créé (avec la date/heure) mais **sans agent, surligné en
jaune**, avec le texte d'origine repris dans le commentaire pour qu'Elo
puisse vérifier (annulation ou vraie faute de saisie).

**⚠️ Non testé ce mois-ci** : aucune donnée de mai 2026 dans le fichier
lectures fourni (le fichier saute d'avril à juin) — le code est écrit et
documenté selon la structure observée sur les autres mois, mais n'a pas pu
être validé par comparaison sur un vrai mois. À refaire dès qu'un exemple
avec des données de mai (ou tout autre mois testable) sera disponible.

**5. Calendrier des événements déjà saisi** (`parse_calendrier_evenements`)
— passthrough pur : le fichier est déjà dans le format cible, on le relit
tel quel sans interprétation, pour pouvoir le fusionner avec les nouvelles
sources sans perdre ce qui a déjà été saisi à la main les mois précédents.

### Principe de fond : les fichiers sources bruts font foi

**Décision actée avec Elo (14/08)** : en cas de désaccord entre le résultat
produit à partir des fichiers sources et un ancien onglet Événements rempli à
la main, ce sont **les fichiers sources qui font référence**, pas l'ancien
onglet — celui-ci peut contenir des ajouts faits "de tête" par Elo, impossibles
à retracer plusieurs mois après coup.

### Test réalisé — mai 2026 (comparaison avec l'ancien onglet Événements, déjà
rempli à la main pour ce mois-là, utilisé ici uniquement comme repère de
qualité)

- **Congés** : 13 événements générés. Concordance quasi totale avec l'ancien
  onglet. Deux écarts identifiés et acceptés comme des différences de source
  (pas des bugs) :
  - Samedi 9 mai : l'ancien onglet indique Guillaume absent, le fichier
    congés source ne le montre pas absent ce jour-là → non reproduit (le
    fichier source fait foi).
  - Les demi-journées (0,5) génèrent désormais une alerte jaune au lieu
    d'être fondues dans un congé 9h-19h comme le faisait l'ancien onglet.
- **Accueil crèches** : 2 événements générés (7 et 28 mai, Tiphaine), tous
  deux avec l'heure en alerte jaune (A1 vide dans le fichier fourni). Absents
  de l'ancien onglet Événements — pas d'explication trouvée, à surveiller
  dans le temps mais pas bloquant (le fichier source fait foi).
- **Accueil de classe** : 11 événements générés, 9 avec un intervenant
  identifié — tous concordent parfaitement avec l'ancien onglet (même agent,
  même horaire). 2 événements sans initiale dans le fichier source →
  correctement mis en alerte jaune plutôt que de deviner.
- **Lecture jeudi matin** : pas de données de mai dans le fichier fourni,
  non testable ce mois-ci (cf. ci-dessus).

**Verdict** : la lecture automatique des fichiers sources est fiable ; les
seuls écarts viennent d'informations incomplètes ou tranchées "de tête" par
Elo à la source, pas d'erreurs d'interprétation.

### Fichiers livrés cette session

- `sources_to_evenements.py` — module de parsing + génération de l'onglet
  Événements (autonome, pas encore branché sur `planning_engine.py`).
- `Evenements_Mai2026_TEST.xlsx` — résultat du test décrit ci-dessus (congés
  + accueil crèches + accueil de classe ; lecture jeudi matin non incluse,
  pas de données pour mai).
- Ce fichier de contexte (v29).

---

## 17. SESSION DU 18/08 — App Streamlit fonctionnelle (blocs 1 et 2) + corrections critiques du moteur

### Vue d'ensemble de la demande

Reprise de l'app Streamlit (l'ancienne ne fonctionnait plus du tout).
Décision : 3 blocs indépendants sur **une seule page en défilement** (pas
d'onglets séparés) :
1. Créer l'onglet Événements (upload de plusieurs fichiers sources bruts)
2. Générer le planning mensuel (upload Événements + Préparation)
3. Vérifier un planning déjà rempli à la main (pas encore construit)

Chaque bloc est totalement autonome : on peut utiliser le bloc 3 sans être
passé par les blocs 1 et 2 avant, ni le même jour, ni dans la même session.

### 17.1 Déploiement Streamlit Cloud (nouveau)

L'app est déployée sur Streamlit Cloud, dépôt GitHub `planning-mediatheque`.
Plusieurs pannes de déploiement rencontrées et résolues, à retenir pour la
suite :
- **`requirements.txt` est obligatoire et doit porter ce nom exact**
  (pas `requirement.txt` au singulier) — sans lui, aucune bibliothèque tierce
  n'est installée (ni `openpyxl`, ni `ortools`). Contenu final : `streamlit`,
  `pandas`, `openpyxl`, `ortools`.
- **Attention aux doublons de fichiers renommés automatiquement** :
  re-uploader un fichier existant sans passer par "Edit" (crayon) peut créer
  un fichier `nom (1).py` au lieu d'écraser l'original → erreur d'import
  incompréhensible ("Oh no", message redacted). Toujours vérifier le nom
  exact sur GitHub après un upload.
- **Pour voir le détail complet d'une erreur** : page de gestion de l'app
  (accessible via "Manage app", ou en cliquant sur l'app depuis
  share.streamlit.io) → un panneau de logs (type terminal) est présent,
  parfois discret selon la mise en page de l'interface — c'est là qu'apparaît
  le vrai message d'erreur (celui affiché à l'utilisateur est volontairement
  tronqué/"redacted").
- `ortools` est une bibliothèque volumineuse : premier démarrage après son
  ajout plus long que d'habitude (~1-2 min), normal.

### 17.2 Bloc 1 — Créer l'onglet Événements : corrections de structure réelle

Les vrais fichiers sources d'Elo (testés cette session) ont révélé une
structure différente de celle documentée au §15 (écrite sans exemple réel).
**Ce qui suit est la référence à jour.**

**Accueil crèches → renommé "Accueil libre crèche" partout** (app, onglet
Événements, planning final). Structure réelle : cellule A1 = heure fixe
("10h -10h30"), colonne A = mois (rempli une seule fois par bloc), colonne
B = jour ("jeudi 7") — **cellule fusionnée sur plusieurs lignes** quand
plusieurs crèches viennent le même jour (donc "vide" sur les lignes
suivantes, il faut reporter la valeur du dessus), colonne C = nom de la
crèche (une ligne par crèche). **Ce sont des visites en créneaux libres :
aucun agent n'est jamais affecté** (avant cette session, le code assignait
Tiphaine automatiquement à tort). Aucune alerte pour absence d'agent — c'est
l'état normal. Seule une heure introuvable en A1 déclenche une alerte.

**Accueil de classe** — structure réelle : colonne A = initiales de
l'intervenant (pas colonne B comme documenté au §15), colonne B = date en
**texte sans année** ("mardi 3 novembre" — remplie seulement sur la 1re
ligne du jour, à reporter sur les lignes suivantes du même jour), colonne D
= créneau horaire, colonne E = nom de l'école (événement seulement si
rempli). Le fichier empile plusieurs mois à la suite (blocs "NOVEMBRE",
"DECEMBRE"...), d'où un filtrage par nom de mois trouvé dans le texte de la
date plutôt que par position dans le fichier. **Nouveau : colonne J** porte
parfois la mention *"Uniquement en visite libre sur ce créneau"* — dans ce
cas, événement nommé **"Accueil libre école"**, aucun agent, aucune alerte
(comportement normal, comme pour l'accueil crèche). En dehors de ce cas,
comportement inchangé : agent déduit des initiales
(`INITIALES_ACCUEIL_CLASSE` — table corrigée cette session : "GP" était une
erreur de saisie d'Elo, remplacé par "GC"=Guillaume), alerte si initiale
absente ou non reconnue.

**Lecture du jeudi matin → renommé "Lectures AssMat/AssPar" partout.**
Structure réelle entièrement différente de celle documentée au §15 (qui
décrivait un système de ligne "Total" avec du texte libre — inexistant en
pratique) : cellule A1 = heure fixe, colonne B = **vraie date Excel avec
année** (une seule ligne par séance, en haut d'un bloc de lignes fusionnées
— pas de ligne "Total"), colonne J = intervenant(e)(s), plusieurs noms
séparés par `;` si besoin. Alerte si colonne J vide (séance sans intervenant
assigné).

**Calendrier déjà saisi → renommé "Fichier Excel Evenement" dans l'app**,
comportement de réinjection (passthrough) inchangé.

**Nouveau champ "Autres documents"** dans l'app (facultatif, plusieurs
fichiers) — capturés mais **pas encore interprétés automatiquement**, faute
de format connu. À développer si/quand Elo précise leur contenu.

**Nouveau message fixe, bien visible (encadré orange) juste avant le
bouton "Générer"** : rappel de vérifier manuellement les dates de portage,
hors les murs, séances d'éloquence, CAJ, formations, réunions/rdv/événements
exceptionnels — sources non automatisées à ce jour.

**Nouvelle règle générale de surlignage (colonne Agents), dans
`build_onglet_evenements`** : si la colonne Agents est vide OU contient un
texte manifestement provisoire ("à déterminer", "?", "à définir" — variantes
avec/sans accent gérées), **toute la ligne** est surlignée en jaune avec un
commentaire — **sauf si l'intitulé de l'événement contient "libre"**
(Accueil libre crèche / Accueil libre école), où l'absence d'agent est
normale et n'est jamais signalée. Cette règle remplace/généralise
l'ancienne logique de surlignage au cas par cas par parseur.

**Bug corrigé — tri chronologique** : le tri de l'onglet Événements
comparait les heures de début comme du TEXTE ("10h" < "9h" alphabétiquement)
au lieu de les convertir en minutes — plusieurs événements le même jour
pouvaient donc apparaître dans le désordre. Corrigé (conversion en minutes
avant comparaison).

### 17.3 Bloc 2 — Générer le planning mensuel : mise en service

- Upload de 2 fichiers séparés (Événements + Préparation mensuelle, comme
  demandé — pas de fusion manuelle côté utilisatrice). En coulisses, l'app
  copie l'onglet Événements du 1er fichier dans le 2e (le moteur
  `compute_full_planning` n'accepte qu'un seul fichier en entrée), puis
  lance le calcul sur ce fichier fusionné temporaire.
- `generer()` dans `generate_planning_excel_septembre.py` **rendu
  paramétrable** (`generer(input_path, output_path)`, avant : chemins fixes
  câblés sur le fichier de septembre) — nécessaire pour fonctionner avec le
  fichier que l'utilisatrice dépose, quel que soit le mois.
- **Bug corrigé — année codée en dur** : les titres de semaine et
  d'en-têtes de jour affichaient toujours "2026" littéralement, peu importe
  la vraie date. Corrigé (année lue dynamiquement depuis chaque date réelle).
- Alertes de couverture (créneaux non pourvus) affichées clairement à
  l'écran après génération (date, section, message), en plus d'être dans le
  fichier.
- **Nom du fichier téléchargé, dynamique** : `Planning_MoisAnnée.xlsx`
  (ex: `Planning_Novembre2026.xlsx`), déduit de la première date réelle du
  planning calculé — remplace l'ancien nom générique fixe.

### 17.4 Bug critique corrigé — Jeunesse sous-dotée par rapport au Planning type

**Symptôme signalé par Elo** : le planning généré n'affichait jamais plus
d'1 agent en Jeunesse, alors que le Planning type en prévoit souvent 2 ou 3
sur certains créneaux (colonnes E/F/G du planning type = 3 agents Jeunesse
séparés, jamais un seul texte avec "/" — même convention que partout
ailleurs dans le projet).

**Cause** : `parse_planning_type` (dans `planning_engine_cpsat.py`) ne
lisait QUE la colonne E (`row[4]`) pour la Jeunesse, en supposant (à tort)
que plusieurs agents seraient écrits dans une seule cellule séparés par
"/". Les colonnes F et G (2e et 3e agent Jeunesse) étaient silencieusement
ignorées à la lecture — jamais transmises au calcul, qui ne pouvait donc
jamais viser plus d'1 agent hors vacances scolaires.

**Correctif** : lecture des 3 colonnes (E, F, G), agrégées en une seule
liste d'agents Jeunesse pour le créneau. Vérifié sur le fichier réel de
novembre d'Elo : un créneau Mardi 17h-19h qui prévoit "Robin + Agnès" au
planning type affiche désormais bien les deux dans le planning généré (au
lieu de Robin seul avant correction).

⚠️ **Impact à surveiller** : cette correction change potentiellement le
nombre d'agents Jeunesse sur de nombreux créneaux à travers tout le mois
(partout où le PT en prévoyait 2 ou 3). Il serait utile de relancer
`compare_planning.py` contre la référence de mai pour vérifier l'ampleur de
l'effet et confirmer qu'aucune régression n'apparaît ailleurs (cf. §16
"Sur l'horizon", déjà identifié avant cette session pour une autre raison
— la mise à jour de l'équité — désormais d'autant plus nécessaire).

### 17.5 Bug corrigé — vue par agent : durée d'événement calquée sur le gros bloc fusionné

**Symptôme signalé par Elo** : Stéphanie en accueil de classe de 10h à 11h
apparaissait, dans sa vue par agent, comme "en accueil de classe" de 10h à
12h30 (durée du gros créneau fusionné du planning principal, pas celle de
l'événement réel). Même souci pour Macha/Stéphane en portage 9h-12h,
affichés seulement à partir de 10h (heure d'ouverture, alors que la grille
fine de la vue par agent démarre bien à 9h).

**Cause** : la vue par agent (`generer_vue_agent`) détectait la présence
d'un événement (Accueil/Animation, Réunion) via une formule Excel qui allait
chercher le texte déjà écrit dans le gros bloc fusionné du planning
principal (colonnes R/S) — donc le même texte se retrouvait recopié sur
CHAQUE heure fine que ce bloc chevauchait, sans jamais vérifier si
l'événement réel couvrait bien toute cette durée. Un événement démarrant
avant l'heure d'ouverture (portage 9h) n'avait en plus aucun bloc auquel se
raccrocher (le planning principal ne commence qu'à l'ouverture) et
disparaissait purement et simplement.

**Correctif** : nouvelle fonction `_evenement_pour_agent_creneau` qui
compare directement, en Python, l'horaire RÉEL de chaque événement
(`ev['cs']`/`ev['ce']`) au créneau fin affiché — plus aucune dépendance au
découpage du planning principal. Les colonnes RDC/Adulte/M&F/Jeunesse
restent, elles, basées sur le gros bloc (légitime : ces affectations
s'appliquent bien à tout le bloc). Vérifié sur le fichier réel de novembre :
accueil de classe 10h-11h n'apparaît plus que sur ce créneau précis ; portage
9h-12h apparaît désormais correctement sur les 3 heures, y compris avant
l'ouverture.

### 17.6 Fichiers livrés cette session

- `app.py` — réécrit de zéro (page unique en défilement, blocs 1 et 2
  fonctionnels, bloc 3 en aperçu).
- `sources_to_evenements.py` — corrections de structure (§17.2).
- `planning_engine_cpsat.py` — correctif Jeunesse (§17.4).
- `generate_planning_excel_septembre.py` — paramétrage `generer()`,
  correctif année codée en dur, correctif vue par agent (§17.3, §17.5).
- `requirements.txt` — nouveau (§17.1).
- Ce fichier de contexte (v30).

---

## 19. SESSION DU 19/08 — Bloc 3 (Vérification de planning) construit et branché

### Contexte

Une fois le planning généré (bloc 2), deux personnes le modifient à la main
en parallèle : Elo (corrections manuelles) et chaque agent (zone de notes
W-Z, §14). Ces ajustements manuels peuvent introduire des erreurs que le
calcul automatique ne peut plus rattraper. Le bloc 3 relit un planning déjà
rempli et signale les contradictions avec les contraintes dures — sans rien
recalculer, juste comme un correcteur qui relit une copie.

### Nouveau fichier : `planning_checker.py`

Fonction principale : `verifier_planning(file_bytes) -> list[Anomalie]`.
Un seul fichier en entrée (le planning déjà rempli), pas besoin de
redéposer le fichier de Préparation séparément (cf. onglets cachés
ci-dessous). Anomalies classées 🔴 (impossibilité certaine) ou 🟡 (suspect,
à vérifier).

**Principe général** : pour chaque jour de chaque onglet `Semaine_N`,
construit la liste des "occurrences" de chaque agent (où il apparaît :
sections B-G, Accueil/Animation, Réunion, Absence — colonnes H/I/J,
dépliées si fusionnées), puis compare :

- **R1+R4 — Horaires contractuels + pause déjeuner** : réutilise
  directement `agent_disponible()` de `planning_engine_cpsat.py` (la
  fonction que le moteur de calcul utilise lui-même pour décider si un
  agent peut être placé sur un créneau) — même règle, même vérité, pas de
  logique dupliquée. Résultat certain (🔴) si les onglets de préparation
  sont présents dans le fichier (cf. ci-dessous).
- **R2 — Congé = jamais planifié** : chevauchement entre une occurrence
  "Absence" (texte de la colonne J, qui contient les noms) et toute autre
  occurrence du même agent.
- **R3 — Un agent à un seul endroit à la fois** : chevauchement temporel
  entre deux occurrences quelconques du même agent (toutes colonnes
  confondues, y compris Accueil/Réunion vs sections).
- **R5 — Habilitations par section** : table `Affectations` si disponible,
  sinon liste codée en dur en secours (mode dégradé).
- **R6 — Vacataires jamais au RDC.**
- **R7 — Vacataire seul en Jeunesse** : autorisé seulement 12h-14h.
- **R8 — Eloïse jamais planifiée.**
- **R9 — Roulement samedi Bleu/Rouge** (nouveau, nécessite les onglets de
  préparation) : compare la couleur du jour (lue dans le titre du bloc
  journée, ex. "SAMEDI BLEU") à la couleur individuelle de l'agent
  (`Roulement_Samedi`, exceptions par semaine incluses).
- **Présence vacataire** (nouveau, nécessite les onglets de préparation) :
  vérifie que le vacataire n'est planifié que sur les dates/horaires du
  tableau "Présence Vacataire" du `Paramètres`.
- **Garde-fou** : si aucune info n'est trouvée ni dans H/I/J ni dans les
  notes W-Z pour tout un jour, message neutre "à vérifier si c'est normal"
  plutôt qu'un silence (protège contre un bug de lecture qui ferait croire
  à tort que tout va bien).
- **Cohérence notes agents (W-Z) ↔ H/I/J** : si une note ajoutée par un
  agent ne se retrouve dans aucune des colonnes H/I/J du jour, signalée en
  🟡 (la formule de cascade a peut-être raté le texte).

### Onglets de préparation recopiés "très masqués" (demande utilisatrice)

Idée d'Elo pour éviter d'avoir à redéposer 2 fichiers (planning + fichier de
préparation) à chaque vérification : `generate_planning_excel_septembre.py`
recopie désormais, à la fin de `generer()`, les valeurs des onglets
`Paramètres`, `Horaires_Des_Agents`, `Affectations`, `Roulement_Samedi` du
fichier de préparation source dans le classeur généré, sous les noms
`_prep_Paramètres`, `_prep_Horaires_Des_Agents`, `_prep_Affectations`,
`_prep_Roulement_Samedi`, avec `sheet_state = 'veryHidden'` — invisibles
dans Excel, y compris via le menu clic droit > Afficher (contrairement à un
simple masquage). Seul un programme (`planning_checker.py`, ou VBA) peut les
relire. Ce n'est pas un fichier séparé : les onglets restent physiquement
dans le même classeur .xlsx.

**L'onglet `Événements` n'est volontairement PAS recopié** (demande
explicite d'Elo) : une fois le planning modifié à la main, il n'est plus à
jour. La référence reste les colonnes H/I/J + les notes W-Z du planning
lui-même.

`planning_checker.py` cherche ces onglets `_prep_*` au chargement
(`charger_donnees_preparation()`) et, s'ils sont présents, appelle
directement `parse_parametres`, `parse_affectations`,
`parse_horaires_agents`, `parse_roulement_samedi` de
`planning_engine_cpsat.py` — une seule source de vérité pour la lecture de
ces onglets, partagée entre le moteur de calcul et le vérificateur.

**Mode dégradé** : si un fichier a été généré avec une version antérieure
de `generate_planning_excel_septembre.py` (donc sans ces onglets cachés),
`planning_checker.py` se rabat automatiquement sur une vérification
approximative (horaires devinés via la "vue par agent", pause déjeuner en
🟡 systématique, roulement samedi et présence vacataire non vérifiables) et
le signale explicitement en tête de rapport.

> **⚠️ Mise à jour 24/08 (voir §25) : `Paramètres` et `Affectations` ne sont
> plus dans ce lot "très masqué"** — ils sont recopiés visibles et
> modifiables (même mécanisme que Planning_type ci-dessous, mais éditable).
> `charger_donnees_preparation()` les cherche désormais d'abord sous leur
> nom visible, avec repli automatique sur `_prep_Paramètres`/
> `_prep_Affectations` pour les fichiers générés avant ce changement. Le
> paragraphe ci-dessus reste exact pour `Roulement_Samedi`, `Besoins_Jeunesse`,
> `Jours_speciaux` et l'ancien `Horaires_Des_Agents` (liste à plat, repli),
> qui restent très masqués.

### Autre changement (demande utilisatrice) : surlignage jaune des événements incomplets

Dans `generate_planning_excel_septembre.py` : quand un événement
Accueil/Animation ou Réunion arrive sur le planning sans agent renseigné,
la cellule H ou I est surlignée en jaune fluo (`FFFFFF00`, même couleur
qu'au bloc 1) avec un commentaire explicatif. Le cas "sans horaire" est
également prévu dans le code par sécurité, mais ne peut normalement plus se
produire à ce stade : `parse_evenements()` exclut déjà en amont tout
événement sans heure de début/fin. Exception : les événements "libre"
(Accueil libre crèche/école) ne sont jamais surlignés, l'absence d'agent y
étant normale (même règle qu'au bloc 1).

### `app.py` — bloc 3 branché

Le bloc 3, jusqu'ici affiché en aperçu grisé, est maintenant fonctionnel :
upload d'un fichier planning, bouton "Vérifier le planning", appel à
`verifier_planning()`, affichage du rapport directement dans l'app (pas de
fichier généré) — compteurs 🔴/🟡, puis détail groupé par semaine et par
jour dans des `st.expander`.

### Limite connue

Pour les agents "pause flexible" avec deux mentions "Arrivée" le même jour
dans la vue par agent (avant/après une pause variable), le **mode dégradé**
(sans onglets de préparation) ne garde que la borne la plus tôt comme
référence, pour éviter une fausse alerte sur le retour de pause — un peu
moins strict pour ces agents-là dans ce mode précis. Non-problème en mode
complet (onglets de préparation présents), qui utilise directement les
horaires contractuels exacts via `agent_disponible()`.

### Tests effectués

- `evenement_incomplet()` : 4 cas testés (avec agent, sans agent, "libre"
  sans agent, sans horaire) — tous corrects.
- Surlignage jaune + commentaire vérifié sur une cellule réelle.
- `copier_onglets_preparation_caches()` : vérifié que les onglets sont bien
  recopiés et passent en `veryHidden`.
- `planning_checker.py` en mode complet : testé avec un classeur de
  préparation synthétique — détecte correctement un agent placé hors
  horaire (dépassement + empiètement sur la pause), ne signale rien à tort
  sur un cas valide, et détecte correctement un agent placé un samedi de la
  mauvaise couleur (roulement).
- `app.py` : démarrage Streamlit vérifié sans erreur d'import après le
  branchement du bloc 3.

### Validation en conditions réelles (même session, après premier déploiement)

Premier test réel par Elo sur un planning de novembre généré via l'app
déployée sur Streamlit Cloud.

**Incident 1 — onglets `_prep_*` absents du fichier généré.** Cause :
Streamlit Cloud n'avait pas encore relu le nouveau code au moment de la
génération (l'app tournait sur l'ancienne version malgré le push GitHub).
Pas un bug du code — résolu par un "Reboot" explicite de l'app depuis le
tableau de bord Streamlit Cloud, qui force la relecture de GitHub. **Point
de vigilance pour la suite** : après tout push, un Reboot manuel est
nécessaire avant de considérer la mise à jour comme active (ne pas supposer
qu'un simple push suffit).

**Incident 2 — bug réel trouvé et corrigé : "MF" vs "M & F".** Une fois les
onglets `_prep_*` bien présents, la vérification des habilitations (R5)
générait une soixantaine de fausses alertes 🔴 "section non habilitée" sur
quasiment tous les agents travaillant en M & F. Cause : l'onglet
`Affectations` écrit cette section sous la forme `"MF"` (sans espace ni
esperluette), alors que l'en-tête de colonne du planning l'appelle
`"M & F"` — comparaison de chaînes strictement égales, donc jamais
reconnues comme identiques. **Corrigé** dans `planning_checker.py` par une
nouvelle fonction `canon_section()` (retire espaces/accents/majuscules/
symboles avant comparaison), utilisée dans la règle R5. Revérifié sur le
même fichier réel : les ~60 fausses alertes ont disparu, 0 anomalie rouge
restante (planning de novembre propre, comme attendu), les 20 alertes 🟡
restantes sont toutes des "aucun événement noté" légitimes (mois sans
événement source). Seul `planning_checker.py` a dû être repoussé pour ce
correctif — pas de nouveau Reboot ni de régénération nécessaires côté
`generate_planning_excel_septembre.py`/`app.py`.

**Non encore testé** : introduire volontairement des erreurs (agent hors
horaire, double affectation, congé + planifié...) dans un planning réel
pour confirmer que `planning_checker.py` les détecte toutes sans fausse
alerte gênante — c'est le test qui reste à faire avant de considérer le
bloc 3 pleinement validé.


## 18. SUR L'HORIZON (mis à jour 19/08, après 1er test réel)

> ⚠️ Liste ci-dessous conservée pour l'historique. La liste **à jour** des tâches ouvertes est au **§23** (session du 20/08/2026).

- **Terminer la validation du bloc 3** — introduire volontairement 1-2
  erreurs connues dans un planning réel déjà généré (agent hors horaire,
  double affectation, congé + planifié en même temps...) et confirmer que
  `planning_checker.py` les détecte toutes, sans fausse alerte gênante.
  Après le correctif "MF"/"M & F" (§19), plus aucune fausse alerte connue à
  ce stade, mais un seul test réel n'exclut pas d'autres variantes
  d'écriture non encore rencontrées (ex. dans les autres onglets de
  préparation, ou pour d'autres sections).
- **Relancer `compare_planning.py` contre la référence de mai** — devenu
  d'autant plus important après le correctif Jeunesse (§17.4), qui peut
  changer le nombre d'agents Jeunesse sur de nombreux créneaux à travers le
  mois. Objectif : confirmer qu'aucune régression n'apparaît ailleurs suite à
  ce correctif + à la mise à jour de l'équité (déjà en attente depuis avant
  cette session).
- **Champ "Autres documents" du bloc 1** — capturé dans l'app mais pas
  encore interprété automatiquement (§17.2). À développer dès qu'Elo précise
  quel(s) type(s) de documents il doit couvrir.
- **Parser la colonne J "visite libre" et les nouvelles règles de nommage**
  — bien tester sur un mois complet réel (au-delà des fichiers de test déjà
  utilisés) une fois d'autres mois disponibles, pour confirmer qu'aucun autre
  cas de figure n'a été manqué.
- **Fichier "Planning avec notes agents" (colonnes W à Z)** — toujours pas de
  parseur dédié, toujours pas d'exemple réel disponible. Format attendu déjà
  documenté au §14 (base de règles de lecture à réutiliser le moment venu).
- Générer les plannings d'octobre et novembre 2026 une fois le bloc 2
  validé en conditions réelles sur plusieurs mois consécutifs.
- Continuer la calibration générale du moteur (sujets antérieurs à cette
  session, cf. sections précédentes) en parallèle du travail sur l'app.



---

## Ajout v33 (09/2026) — Contrainte dure de couverture Jeunesse + RDC/Adulte/M&F dans le checker (Bloc 3)

**Demande utilisatrice** : le checker ne signalait aucune alerte quand un agent
Jeunesse manquait par rapport au planning type. Vérification demandée aussi
pour les autres sections (RDC/Adulte/M&F) en cas de "trou".

**Ce qui a changé**

1. `generate_planning_excel_septembre.py` — la liste des onglets recopiés en
   "très masqué" dans le fichier généré (`ONGLETS_PREPARATION_A_RECOPIER`)
   s'enrichit de trois onglets : `Planning_type`, `Besoins_Jeunesse`,
   `Jours_speciaux`. Jusqu'ici seuls Paramètres/Horaires_Des_Agents/
   Affectations/Roulement_Samedi étaient recopiés — le checker ne pouvait
   donc pas connaître le planning type ni les besoins vacances scolaires.

2. `planning_checker.py` — nouvelle règle **R10 (contrainte dure, 🔴)** :
   pour chaque créneau de chaque jour, dans chaque section (RDC, Adulte,
   M & F, Jeunesse), le nombre d'agent·es affecté·es doit correspondre
   **exactement** à ce que prévoit la référence applicable :
   - RDC / Adulte / M & F : toujours le planning type (Vacances Scolaires
     ou non — seule la Jeunesse change de référence pendant les vacances).
   - Jeunesse : le planning type hors vacances scolaires, ou l'onglet
     Besoins_Jeunesse pendant les vacances scolaires (période déterminée par
     le réglage Semaine_N, avec priorité à un jour marqué "vacances" dans
     Jours_speciaux si cet onglet est présent).
   - Agent manquant par rapport à la référence → 🔴 "trou".
   - Agent en trop par rapport à la référence → 🔴 aussi (le moteur de
     calcul ne dépasse jamais ce nombre ; un dépassement en planning modifié
     à la main est donc une vraie anomalie, pas une simple préférence).
   - Reprend telle quelle la logique déjà validée du moteur de calcul
     (mêmes fonctions `parse_planning_type`, `parse_besoins_jeunesse`,
     `parse_jours_speciaux`, `parse_creneau` importées directement de
     `planning_engine_cpsat.py` — pas de réécriture parallèle, donc pas de
     risque de divergence entre "ce que le moteur calcule" et "ce que le
     checker vérifie").

**Compatibilité avec les anciens fichiers** : un planning généré AVANT cette
mise à jour ne contient pas les onglets `Planning_type` / `Besoins_Jeunesse`
cachés — la règle R10 est alors simplement sautée pour ce fichier (message
"Onglet(s) de préparation manquant(s)" déjà existant, pas de plantage). Il
faut régénérer le planning (Bloc 2) avec le code à jour pour bénéficier de
cette nouvelle vérification.

**Tests effectués** : deux petits fichiers Excel synthétiques construits à la
main (hors vacances scolaires, puis vacances scolaires avec Besoins_Jeunesse)
pour vérifier : cas conforme → aucune alerte ; agent manquant en Jeunesse →
alerte 🔴 ; agent manquant en RDC → alerte 🔴 ; agent en trop en Jeunesse →
alerte 🔴. Les 4 cas se comportent comme attendu. Une erreur a été détectée
et corrigée pendant ces tests dans le test lui-même (tableau Besoins_Jeunesse
incomplet côté synthétique, pas dans le code du checker).

**Reste à faire** : valider R10 sur un vrai planning régénéré (septembre ou
octobre 2026) une fois le déploiement effectué, notamment sur des semaines
"vacances scolaires" réelles avec plusieurs tranches horaires Besoins_Jeunesse
différentes dans un même créneau du planning.

Fichiers modifiés : `generate_planning_excel_septembre.py`,
`planning_checker.py`.

### Complément v33 — Onglet Planning_type : visible et verrouillé (pas très masqué)

Suite à une demande complémentaire, l'onglet Planning_type n'est **plus**
recopié en "très masqué" avec les autres onglets de préparation. Il a son
propre traitement :

- Recopié en onglet **visible**, nommé `Planning_type` (sans préfixe
  `_prep_`), consultable par n'importe quel agent qui ouvre le fichier —
  comme une vitre : on voit à travers, mais on n'y touche pas.
- **Verrouillé** (toutes ses cellules), avec protection de feuille activée,
  pour éviter toute modification accidentelle. Pas de mot de passe (même
  logique que le verrouillage des cellules à formule ailleurs dans le
  fichier) : le but est d'éviter la fausse manipulation, pas d'empêcher un
  usage volontaire.
- Nouvelle fonction dédiée : `embarquer_planning_type_visible()` dans
  `generate_planning_excel_septembre.py`, appelée volontairement APRÈS
  `verrouiller_cellules_formules()` (qui reparcourt tous les onglets et
  aurait sinon déverrouillé les cellules sans formule de ce nouvel onglet).
- `planning_checker.py` (règle R10, cf. ci-dessus) va chercher cet onglet
  sous son nom visible `Planning_type` plutôt que sous une version masquée
  `_prep_Planning_type`.
- Les autres onglets de préparation (Paramètres, Horaires_Des_Agents,
  Affectations, Roulement_Samedi, Besoins_Jeunesse, Jours_speciaux) restent
  très masqués comme avant — seul Planning_type change de statut.

Testé unitairement : onglet bien `visible`, protection de feuille active,
cellules verrouillées, valeurs correctement recopiées.

### Correctif v33bis — Planning_type visible : conservation du formatage d'origine

La première version de `embarquer_planning_type_visible()` ne recopiait que
les VALEURS des cellules (comme les autres onglets de préparation, qui eux
sont invisibles donc leur apparence n'a pas d'importance) — résultat :
toutes les couleurs du planning type disparaissaient dans l'onglet recopié.

Corrigé : nouvelle fonction `_copier_feuille_avec_mise_en_forme()` qui copie,
en plus des valeurs :
- la police, les couleurs de fond, les bordures, l'alignement et le format
  numérique de chaque cellule,
- les cellules fusionnées (ex : les 3 colonnes Jeunesse),
- les largeurs de colonnes et hauteurs de lignes,
- le figeage des volets et l'affichage du quadrillage.

Le verrouillage (protection en lecture seule) est appliqué PAR-DESSUS ce
formatage — la protection est un attribut de cellule indépendant de
l'apparence visuelle, donc verrouiller les cellules n'efface plus les
couleurs.

Testé avec `Planning_type_dernier_design.xlsx` (458 cellules colorées, 10
fusions, largeurs de colonnes) : l'onglet recopié reproduit exactement ces
458 cellules colorées et les 10 fusions, tout en restant verrouillé.

Fichier modifié : `generate_planning_excel_septembre.py`.

---

# SESSION DU 20/08/2026 — Fusion de l'addendum dans ce document

Les sections §20 à §23 ci-dessous proviennent d'un addendum écrit le 20/08/2026 (session où l'accès au v33 complet avait été temporairement perdu — voir §22 dernier point) puis fusionné ici pour ne conserver qu'un seul fichier de contexte.

---

## 20. NOUVELLE FONCTIONNALITÉ (en cours) — "Régénérer un planning existant" (session 20/08/2026)

**Cas d'usage** : le planning du mois est déjà généré, des modifications
manuelles et des notes agents (W-Z) ont été ajoutées, et certaines contiennent
trop d'erreurs. Plutôt que de tout regénérer (et perdre les modifs
pertinentes), l'utilisatrice veut pouvoir régénérer **une semaine complète**
au sein du planning existant, en gardant toutes les autres semaines
strictement intactes.

**Décisions de conception actées avec l'utilisatrice** :
- Portée : **une seule semaine à la fois** par lancement (pas de sélection
  à cheval sur deux semaines). Au sein d'une semaine, on peut en théorie
  cibler un ou plusieurs jours (l'implémentation le permet), mais le seul
  cas testé à ce jour est la **semaine complète**.
- Pour le(s) jour(s) régénéré(s) : on **repart de zéro** sur les colonnes
  B à G (RDC/Adulte/M&F/Jeunesse) — pas de tentative de préserver les
  anciennes affectations B-G. En revanche, les colonnes H/I/J
  (Accueil/Animation, Réunion, Absence) et les notes W-Z **ne sont jamais
  effacées** : elles deviennent des contraintes figées pour le calcul.
- Les conflits entre éléments fixes (ex. un agent noté absent ET en réunion
  au même moment) **ne bloquent PAS** le calcul — celui-ci tourne quand
  même (l'agent est alors considéré indisponible sur l'union des deux
  plages). Ces conflits sont uniquement **signalés visuellement** dans le
  fichier de sortie (bordure rouge + commentaire Excel), à charge pour
  l'utilisatrice de trancher elle-même laquelle des deux notes est la
  bonne — le programme ne doit jamais deviner à sa place.
- Alertes de conflit : recherchées **seulement dans la zone régénérée**,
  pas sur tout le fichier à chaque fois.
- Sortie : toujours un **nouveau fichier téléchargeable**, jamais une
  modification en place du fichier uploadé.

**Architecture en 3 briques, chacune testée indépendamment sur un vrai
fichier avant de passer à la suivante** :

1. **`regeneration_lecture.py`** — lecture seule, ne calcule rien.
   - Fonction principale : `lire_planning_pour_regeneration(file_bytes, semaine_num, jours_a_regenerer)`.
   - Lit les onglets cachés `_prep_*` (mêmes fonctions que le Bloc 3) pour
     les règles.
   - Découpe la semaine en jours via `lire_jours_semaine` (Bloc 3).
   - Construit les **"contraintes figées"** du/des jour(s) à régénérer :
     lit H/I/J + notes W-Z déjà combinées, ne garde QUE les occurrences de
     type Accueil/Animation, Réunion, Absence (jamais les occurrences B-G,
     qui seront effacées). Format de sortie identique à
     `parse_evenements()` du moteur CP-SAT (`{'date','cs','ce','nom',
     'nom_affichage','type','agents'}`), pour brancher directement sur
     `solve_day()` sans rien réinventer.
     - `nom` : version "propre" (sans prénoms ni horaire) — `'congé'`
       exactement pour un congé (convention attendue par la vue par
       agent), sinon nettoyée via `_nom_propre()`.
     - `nom_affichage` : texte complet d'origine, utilisé pour les
       messages d'alerte (plus informatif que `nom`).
   - Calcule les heures déjà travaillées cette semaine sur les jours FIXES
     (pour seeder le compteur d'équité hebdomadaire de la brique 2).
     ⚠️ Limite connue : pour une régénération partielle (jours fixes +
     jours régénérés dans la même semaine), il n'existe pas aujourd'hui de
     moyen fiable de reconstituer précisément ce que l'équité "devrait"
     être après le dernier jour régénéré si des jours fixes suivent — non
     bloquant pour une régénération de semaine complète (le cas testé),
     où ce compteur démarre simplement à zéro comme une génération
     normale.
   - Détecte les **conflits qui resteront vrais après régénération**
     (`ConflitFixe` — chevauchement entre deux occurrences H/I/J/W-Z
     FIXES pour un même agent). Volontairement **différent** d'une
     réutilisation brute du Bloc 3 : les conflits impliquant une
     affectation B-G actuelle (qui sera effacée) sont exclus, sinon on
     signalerait des problèmes qui n'existeront plus après régénération.

2. **`regeneration_calcul.py`** — relance le VRAI moteur CP-SAT
   (`solve_day` de `planning_engine_cpsat.py`, aucune logique de calcul
   dupliquée), jour par jour, dans l'ordre chronologique de la semaine, en
   enchaînant le compteur d'équité (`cumul_hebdo`) d'un jour à l'autre.
   - Fonction principale : `regenerer_jours(lecture_resultat, cumul_hebdo_initial=None)`.
   - Ne bloque jamais sur les conflits (cf. décision ci-dessus) — les
     transmet tels quels dans le résultat (`conflits_a_signaler`) pour la
     brique 3.
   - ⚠️ Deux petites fonctions (`_construire_grille_vacances_jour`,
     `_resoudre_besoins_jour`) sont **dupliquées** depuis
     `planning_engine_cpsat.compute_full_planning`, car elles y sont
     définies en fonctions imbriquées (donc non importables telles
     quelles). À garder en tête si leur logique évolue dans le moteur
     principal — amélioration possible plus tard : les sortir du moteur
     pour les rendre réellement partagées.

3. **`regeneration_ecriture.py`** — produit les bytes du nouveau fichier
   Excel.
   - Fonction principale : `ecrire_regeneration(file_bytes, lecture_resultat, calcul_resultat)`
     → retourne `(nouveaux_bytes, jours_infaisables, agent_sheet_reconstruit)`.
   - Réécrit uniquement les colonnes B à G des jours régénérés (même style
     que le générateur habituel : `agent_cell_style`/`fmt_agents`/
     `GREY_BORDER`, réutilisés depuis `generate_planning_excel_septembre.py`).
     Ne touche JAMAIS aux colonnes H à Z ni aux formules qu'elles
     contiennent.
   - Si un jour n'a pas de solution (infaisable), son contenu B-G existant
     est laissé tel quel (jamais écrasé par du vide) et il est listé dans
     `jours_infaisables`.
   - **Reconstruit l'onglet "vue par agent" (`Semaine_N_Agent`)**, mais
     UNIQUEMENT si toute la semaine a été régénérée (`jours_fixes` vide) —
     réutilise `generer_vue_agent()` du générateur principal, replace
     l'onglet à sa position d'origine dans le classeur (sinon il se
     retrouverait en dernière position). Pour une régénération partielle
     (jours fixes + régénérés), cette reconstruction est **sautée**
     (limite actuelle documentée dans le code) : reconstruire sans
     connaître les événements des jours fixes donnerait une vue par agent
     fausse sur ces jours-là.
   - Pose les **alertes visuelles** sur les `ConflitFixe` : bordure rouge
     épaisse + commentaire Excel sur CHACUNE des deux cellules en cause
     (petit triangle rouge visible, cliquable). Choix technique déjà
     discuté et validé avec l'utilisatrice : on n'écrit PAS de texte
     d'avertissement directement dans la cellule H/I/J elle-même, car
     ces cellules contiennent une formule (celle qui combine les notes
     W-Z) — écrire dedans la détruirait.

**Résultat de la validation** (fichier de test `TEST_1___Planning_Novembre2026.xlsx`,
avec 3 erreurs volontairement introduites par l'utilisatrice — Stéphane en
congé toute la journée mercredi mais affecté en M&F, Anne-Françoise absente
et en réunion en même temps mercredi, Macha + Stéphane en Portage et en
rdv médical en même temps jeudi) :
- Semaine complète régénérée avec succès, 5/5 jours faisables.
- Les 3 conflits sont correctement repérés et signalés visuellement,
  exactement sur les bonnes cellules.
- Les autres semaines (2, 3, 4) et tous les onglets cachés/formules restent
  **strictement identiques** (vérifié cellule par cellule).
- La vue par agent reflète bien "Congé" pour Stéphane, et les événements
  s'affichent proprement sans prénoms ni doublon d'horaire.

---

## 21. DEUX CORRECTIFS DE FOND SUR `planning_checker.py` (Bloc 3) — 20/08/2026

Découverts pendant la construction de la fonctionnalité ci-dessus, mais
**concernent aussi le Bloc 3 existant** — déjà livrés et déployés
indépendamment de la nouvelle fonctionnalité.

**Correctif 1 — cases combinant plusieurs événements ("; ").**
Quand une case H/I/J combine plusieurs notes d'agents différents (séparateur
`'; '`, cf. §14), l'ancienne extraction (`extraire_agents_et_fenetre`) ne
lisait que la **toute dernière parenthèse de la case entière**, perdant
silencieusement tout ce qui précède. Concrètement : dans
`"absence (Anne-Françoise); congé (Stéphane)"`, seul Stéphane était détecté,
Anne-Françoise disparaissait complètement de la vérification.
→ Nouvelle fonction `extraire_occurrences_multiples()` : découpe la case sur
`'; '` AVANT extraction, traite chaque segment indépendamment. Utilisée
partout où `extraire_agents_et_fenetre` servait à construire des occurrences
(`construire_occurrences_jour`). `extraire_agents_et_fenetre` elle-même est
conservée (compatibilité), mais documentée comme imprécise dans ce cas de
figure.
Testé sur fichier réel : 0 alerte perdue, 0 fausse alerte ajoutée, les cas
manqués remontent correctement.

**Correctif 2 — le texte associé à chaque agent devait être individuel, pas
celui de toute la case.**
Suite du correctif 1 : même après le découpage, `construire_occurrences_jour`
attribuait encore à CHAQUE agent trouvé dans une case le texte COMPLET de la
case (tous segments confondus) plutôt que son propre segment. Conséquence
concrète découverte en aval (vue par agent) : un congé combiné avec un autre
événement dans la même case n'était plus reconnu comme "congé" (la
comparaison stricte `== 'congé'` échouait puisque le texte contenait aussi
l'autre segment).
→ `extraire_occurrences_multiples()` renvoie maintenant des quadruplets
`(agent, debut, fin, segment)` — le texte SOURCE individuel de chaque agent,
pas celui de la case entière. `construire_occurrences_jour` utilise ce
segment comme `detail` de l'occurrence. Effet de bord positif : les messages
d'alerte du Bloc 3 sont aussi devenus plus lisibles (texte propre à l'agent
concerné, plus de blocs concaténés illisibles dans les messages).
Testé sur fichier réel : 0 régression (mêmes 18 anomalies, textes plus
propres), amélioration du regroupement des occurrences fusionnées.

**Fichiers livrés suite à ces 2 correctifs** : `planning_checker.py` (à
remplacer sur GitHub + **Reboot manuel Streamlit Cloud**, comme toujours).

---

## 22. AUTRES DÉCOUVERTES TECHNIQUES DE LA SESSION (20/08/2026)

- **openpyxl efface le cache des formules à l'enregistrement.** Écrire un
  fichier avec openpyxl (même en ne touchant qu'une partie des cellules)
  fait perdre la **valeur affichée en mémoire** de toutes les formules non
  touchées (ex. H/I/J), même si la formule elle-même reste intacte au
  caractère près. Concrètement : un fichier fraîchement régénéré par l'outil
  peut sembler avoir des cases H/I/J vides si on le relit immédiatement avec
  un script (`data_only=True`) — c'est un artefact, pas une perte de
  données. Excel recalcule tout correctement à l'ouverture ; il suffit
  d'enregistrer une fois dans Excel pour que le Bloc 3 (et tout script) y
  revoie clair. Déjà rencontré côté lecture (fichier jamais ouvert dans
  Excel depuis sa dernière modif = formules H/I/J invisibles) — cette
  session a montré que ça s'applique aussi côté écriture.
- **Colonnes B à G ne sont jamais fusionnées** par le générateur (seules
  H/I/J le sont, via `fusionner_cellules_identiques(colonnes=range(8, 11))`).
  Simplifie beaucoup la réécriture partielle : pas de gestion de fusion à
  prévoir pour B-G.
- **La vue par agent (`Semaine_N_Agent`) n'est pas 100% "live"** : les
  colonnes de section (RDC/Adulte/M&F/Jeunesse) sont bien des formules qui
  se recalculent automatiquement depuis `Semaine_N`, MAIS le marquage gris
  "Congé" et les libellés d'événements (Accueil/Réunion) sont écrits en
  **valeur figée** au moment de la génération — ils ne se mettent PAS à
  jour tout seuls si on ajoute une note W-Z après coup, y compris en dehors
  du contexte de la régénération partielle. Comportement pré-existant de
  l'outil, découvert à cette occasion.
- **Accès aux fichiers du projet en cours de conversation** : la liste des
  fichiers du projet fournie à Claude est prise "en photo" au début d'une
  conversation. Si des fichiers sont ajoutés/modifiés dans le projet APRÈS
  le début de la conversation, Claude peut ne plus y avoir accès tant que la
  conversation continue (rencontré deux fois cette session : perte
  d'accès à `generate_planning_excel_septembre.py`, puis à
  `context_projet_mediatheque_v33.md`). Parade : renvoyer le fichier
  directement dans le chat plutôt que de compter sur la mise à jour
  automatique du projet.

---

## 23. À FAIRE — liste actualisée au 20/08/2026 (remplace la liste du §18 ci-dessus)

- ~~Intégrer les 3 briques (`regeneration_lecture.py`, `regeneration_calcul.py`,
  `regeneration_ecriture.py`) dans une 4ᵉ section de `app.py` ("Régénérer
  un planning existant")~~ → **fait** (constaté le 24/08 en relisant `app.py` :
  upload, sélection semaine + jour(s), bouton, résumé + téléchargement, tout
  y est — la session qui a fait ce travail n'a pas été documentée ici).
- Étendre le support du **jour unique / plusieurs jours au sein d'une
  semaine** (l'implémentation le permet déjà côté briques 1 et 2, jamais
  testé en pratique — seule la semaine complète l'a été). Vérifier
  notamment le comportement de la vue par agent dans ce cas (actuellement
  sautée si jours fixes présents).
- Réfléchir à une solution pour le seeding de l'équité hebdomadaire en cas
  de régénération partielle (jours fixes + régénérés mélangés).
- Root : sortir `_construire_grille_vacances_jour` /
  `_resoudre_besoins_jour` du moteur principal pour éviter la duplication
  actuelle avec `regeneration_calcul.py`.
- Continuer la validation du Bloc 3 avec erreurs injectées (chantier déjà
  en cours avant cette session, cf. v33 §"remaining validation").
- ~~Générer et valider les plannings d'octobre et novembre 2026.~~ →
  **octobre fait** (cf. §24.6) ; **novembre reste à faire**.
- ~~Une fois le v33 récupéré : produire un vrai v34 fusionné~~ → **fait** : ce document EST le v34 fusionné (fusion réalisée le 20/08/2026, plus besoin de conserver l'addendum séparé).

- **Reporté depuis la liste précédente (§18), toujours ouvert** :
  - Relancer `compare_planning.py` contre la référence de mai (confirmer
    qu'aucune régression n'est apparue depuis les derniers correctifs).
  - Champ "Autres documents" du bloc 1 — capturé dans l'app mais pas encore
    interprété automatiquement ; à développer dès qu'Elo précise quel(s)
    type(s) de documents il doit couvrir.
  - Parser la colonne J "visite libre" et les nouvelles règles de nommage —
    bien tester sur un mois complet réel une fois d'autres mois disponibles.

---

## 24. SESSION DU 22/08/2026 — Grille "horaires d'équipes" harmonisée + lecture directe par le moteur (remplace la liste à plat)

### 24.1 Contexte / demande utilisatrice

Elo maintient un fichier collaboratif séparé (`planning_des_agents.xlsx`,
onglet "horaires d'équipes") : une grille de fiches par agent (4 blocs de
service côte à côte — ADULTES, JEUNESSE, MUSIQUE, DIRECTION/ADMINISTRATIF —
avec jusqu'à 5 agents par bloc), bien plus lisible que la liste à plat
`Horaires_Des_Agents` que le moteur lisait jusqu'ici. Jusqu'à cette session,
elle retapait ces horaires à la main dans le fichier de préparation à chaque
mise à jour. Objectif : que le moteur lise cette grille directement.

### 24.2 Harmonisation du fichier source (avant tout changement de code)

Le fichier `planning_des_agents.xlsx` contenait déjà des formules (total
jour, total semaine) mais avec des incohérences :
- des cases visuellement vides contenant en fait un espace insécable
  (`\xa0`), cassant les formules de total (`#VALUE!`) — 8 cas nettoyés ;
- une heure tapée en texte ("08h45" au lieu d'une vraie heure) — 1 cas
  corrigé ;
- des formats de cellule "horloge" au lieu de "durée" par endroits.

Corrections apportées, en conservant le rendu à l'identique (mêmes couleurs,
fusions, largeurs) :
- formule de total jour harmonisée partout : `=(fin_matin-début_matin)+(fin_aprèm-début_aprèm)`
- formule de total semaine harmonisée partout : cumul vendredi (mar-ven) et
  samedi (mar-sam)
- fiche de Lydie vidée (elle a quitté l'équipe)
- **cellules à formule verrouillées** (protection de feuille activée, sans
  mot de passe — comme pour `verrouiller_cellules_formules` ailleurs dans le
  projet), toutes les autres cellules restant modifiables

**⚠️ Piège découvert (à retenir pour tout futur fichier avec hachures) :**
le fichier source utilisait déjà des hachures (`patternType='lightUp'`,
noir sur fond bleu clair) pour marquer les demi-journées non travaillées —
**pas du gris plat**. La première passe de recalcul via `recalc.py`
(LibreOffice, obligatoire pour vérifier les formules) a **aplati ces
hachures en gris plat** en resauvegardant le fichier. Solution retenue :
faire tout le travail de formules/mise en forme avec openpyxl uniquement,
vérifier les formules sur une **copie jetable** passée par `recalc.py`
(jamais livrée), et livrer la version jamais touchée par LibreOffice. Cette
même précaution s'applique maintenant partout où un fichier contient des
`PatternFill` non-solides.

### 24.3 Nouveau parseur dans `planning_engine_cpsat.py`

- `parse_horaires_agents_grille(raw)` : lit la grille directement (4 blocs ×
  5 emplacements d'agent × 5 jours), retourne exactement le même format que
  l'ancien `parse_horaires_agents` : `{agent: {jour: (dm, fm, da, fa)}}` en
  minutes. Éloïse exclue systématiquement (comparaison insensible aux
  accents/casse). Réutilise `hhmm_to_min` telle quelle (gère déjà les objets
  `time` et les chaînes "08h45").
- **`_detecter_onglet_horaires_grille(raw)`** : retrouve l'onglet par sa
  mise en page (cases A6="ADULTES" et H6="JEUNESSE"), pas par son nom.
  Branché dans `load_excel_data()` : si l'onglet canonique
  `"horaires d'équipes"` n'existe pas tel quel, un alias est ajouté vers
  l'onglet détecté. Ainsi, peu importe comment l'onglet est nommé dans le
  fichier de préparation, le reste du code le trouve sous son nom habituel.
  **Découvert nécessaire en testant sur le vrai fichier d'octobre** : Elo
  avait nommé l'onglet `horaires_des_agents` (pas `horaires d'équipes`) —
  un simple alias de casse/ponctuation n'aurait pas suffi, d'où la
  détection par contenu plutôt que par nom.
- Ancien `parse_horaires_agents` (liste à plat) conservé intact, utilisé en
  repli automatique si l'onglet grille n'est trouvé nulle part.

### 24.4 Onglet visible dans le planning généré (comme Planning_type)

Dans `generate_planning_excel_septembre.py` :
- `'Horaires_Des_Agents'` retiré de `ONGLETS_PREPARATION_A_RECOPIER` (la
  liste des onglets recopiés en masqué) ; repli conservé (copie masquée de
  l'ancien format uniquement si la grille est absente du fichier source).
- Nouvelle fonction **`embarquer_horaires_agents_visible(wb, raw)`**,
  jumelle exacte de `embarquer_planning_type_visible` : recopie la grille en
  onglet visible et verrouillé (mise en forme complète conservée, hachures
  comprises), appelée juste après `embarquer_planning_type_visible` (même
  contrainte d'ordre : après `verrouiller_cellules_formules`).

### 24.5 Bug corrigé dans `planning_checker.py`

`charger_donnees_preparation` cherchait la grille sous sa version masquée
(`_prep_horaires d'équipes`), qui n'existe plus puisqu'elle est maintenant
recopiée **visible** (comme Planning_type). Résultat avant correctif : le
Bloc 3 tombait systématiquement en "mode dégradé" et remontait de fausses
alertes 🔴 (ex. horaires d'Agnès/Macha en apparence violés, alors que les
vraies données de contrat n'avaient jamais été lues). Corrigé : la boucle de
chargement traite maintenant `'Planning_type'` ET `"horaires d'équipes"`
comme onglets "sans préfixe" (visibles), au lieu de `'Planning_type'` seul.

### 24.6 Test réalisé — octobre 2026

Fichiers fournis par Elo : `Evenements_Octobre2026.xlsx` (sortie Bloc 1) +
`OCTOBRE_Preparation_Planning_Mediatheque_Modele.xlsx` (onglet horaires
nommé `horaires_des_agents`, cf. §24.3). Après fusion de l'onglet
Événements dans le fichier de préparation :
- génération complète réussie (4 semaines, 30 événements pris en compte) ;
- Bloc 3 (vérification) : **0 anomalie rouge** après le correctif §24.5 (20
  anomalies jaunes bénignes, du type "aucun événement noté ce jour, à
  vérifier si normal") ;
- onglet "horaires d'équipes" bien présent dans le fichier généré : visible,
  verrouillé, hachures intactes.

### 24.7 Fichiers livrés cette session

- `planning_engine_cpsat.py` (nouveau parseur grille + détection par
  contenu + repli)
- `planning_checker.py` (correctif §24.5)
- `generate_planning_excel_septembre.py` (nouvel onglet visible §24.4)
- `horaires_agents_harmonise.xlsx` (fichier source harmonisé, à réintégrer
  dans le fichier de préparation d'Elo — pour rappel, le nom de l'onglet
  n'a plus d'importance grâce à la détection par contenu)
- `Planning_Octobre_2026_TEST.xlsx` (résultat du test bout-en-bout §24.6)

### 24.8 Reste à faire

- Générer/valider le planning de **novembre 2026** avec le même circuit.
- Vérifier avec Elo si elle souhaite renommer son onglet en
  `horaires d'équipes` par cohérence (pas obligatoire, la détection par
  contenu fonctionne quel que soit le nom).
- Reproduire l'harmonisation (§24.2) si le fichier collaboratif source est
  à nouveau modifié par l'équipe et réintroduit des cases invisibles /
  heures en texte.

---

## 25. SESSION DU 24/08/2026 — Paramètres/Affectations/horaires d'équipes visibles et modifiables + correctif Bloc 3/4

### 25.1 Contexte / demande utilisatrice

Cas d'usage de départ : le 15 septembre, Stéphane (jusque-là MF uniquement)
est formé pour la section Adulte, et il faut pouvoir regénérer le planning
(régénération partielle, Bloc 4) en tenant compte de cette nouvelle
habilitation, sans redéposer le fichier de préparation d'origine — juste en
éditant directement le planning déjà généré. Extension immédiate à un
deuxième cas concret : correction d'un horaire d'agent en cours de mois via
la grille "horaires d'équipes" embarquée dans le planning généré.

### 25.2 Écart constaté entre la doc (§24.5) et le code réellement livré

En relisant le vrai `planning_checker.py` fourni par Elo pour cette session,
`charger_donnees_preparation()` ne traitait **que** `'Planning_type'` comme
onglet "sans préfixe" — pas `"horaires d'équipes"`, contrairement à ce que
§24.5 ci-dessus affirme avoir corrigé. Pire : `ONGLETS_PREP_NOMS` ne
contenait même pas d'entrée pour la grille, et le code n'importait pas
`parse_horaires_agents_grille` / `ONGLET_HORAIRES_GRILLE` — seule l'ancienne
liste à plat `_prep_Horaires_Des_Agents` (`parse_horaires_agents`) était
lue, y compris par la régénération partielle (Bloc 4), qui réutilise
`charger_donnees_preparation()` telle quelle. Concrètement : avant cette
session, éditer la grille "horaires d'équipes" du planning généré n'avait
**aucun effet** sur une régénération — ni sur le Bloc 3. Cause exacte de
l'écart avec §24.5 non identifiée (fix perdu entre deux sessions ? jamais
réellement livré malgré la doc ?) — à garder en tête : **vérifier le code
réel plutôt que de faire confiance aveuglément à une note "corrigé" dans ce
document**, il peut être en avance sur ce qui a vraiment été déployé.

### 25.3 Changements de conception actés avec Elo

- `Paramètres` et `Affectations` : retirés du lot "très masqué"
  (`ONGLETS_PREPARATION_A_RECOPIER` dans `generate_planning_excel_septembre.py`),
  désormais embarqués **visibles et librement modifiables** (protection de
  feuille active pour la structure — pas d'insertion/suppression de ligne
  ou colonne — mais AUCUNE cellule verrouillée, y compris l'en-tête).
- `"horaires d'équipes"` : reste embarqué visible, mais **n'est plus en
  lecture seule** — même traitement que Paramètres/Affectations désormais
  (avant : verrouillage total façon "vitre", comme Planning_type qui, lui,
  reste en lecture seule).
- `Planning_type` : **inchangé**, reste visible et verrouillé (Elo n'a pas
  demandé à le rendre modifiable).
- Justification du choix : Paramètres/Affectations/horaires d'équipes sont
  des données qu'Elo doit pouvoir corriger en cours de mois (nouvelle
  habilitation, changement de contrat, absence de dernière minute) et voir
  répercutées par une régénération partielle ; Planning_type est une
  référence de structure, jamais éditée au coup par coup.

### 25.4 Changements de code

**`generate_planning_excel_septembre.py`**
- `ONGLETS_PREPARATION_A_RECOPIER` : `Paramètres` et `Affectations` retirés
  (ne restent en très masqué que `Roulement_Samedi`, `Besoins_Jeunesse`,
  `Jours_speciaux`, + repli `Horaires_Des_Agents`).
- Nouvelle fonction générique `_embarquer_prep_visible_modifiable(wb, raw,
  nom_source)` : recopie un onglet visible, structure protégée, **toutes
  les cellules déverrouillées** (contrairement à
  `embarquer_planning_type_visible`, qui verrouille tout).
- `embarquer_parametres_visible` / `embarquer_affectations_visible` :
  nouvelles, appellent la fonction générique ci-dessus.
- `embarquer_horaires_agents_visible` : réécrite pour appeler la même
  fonction générique au lieu de verrouiller toutes les cellules — seul son
  comportement change, sa signature et son emplacement d'appel dans
  `generer()` restent identiques.
- Les 4 fonctions `embarquer_*` sont appelées dans cet ordre dans
  `generer()`, toujours **après** `verrouiller_cellules_formules(wb)` (même
  contrainte d'ordre que documentée en §24.4) :
  `embarquer_planning_type_visible` → `embarquer_horaires_agents_visible` →
  `embarquer_parametres_visible` → `embarquer_affectations_visible`.

**`planning_checker.py`**
- Import ajouté : `parse_horaires_agents_grille`, `ONGLET_HORAIRES_GRILLE`
  depuis `planning_engine_cpsat`.
- Nouvelle constante `ONGLETS_PREP_SANS_PREFIXE = {'Planning_type',
  'Paramètres', 'Affectations'}` (remplace le test `if nom == 'Planning_type'`
  ponctuel).
- Nouvelle fonction `_detecter_grille_horaires_dans_classeur(wb)` : retrouve
  la grille dans le classeur **déjà généré** par sa mise en page (case
  A6='ADULTES', H6='JEUNESSE'), pas par son nom d'onglet — même principe que
  `_detecter_onglet_horaires_grille` côté fichier de préparation
  (`planning_engine_cpsat.py`), mais réimplémentée localement (fonction
  d'origine privée, non importée) : tolérant si Elo renomme cet onglet à sa
  façon dans le planning généré.
- `charger_donnees_preparation()` : cherche désormais Paramètres/Affectations
  sous leur nom visible en priorité (repli `_prep_...` pour les fichiers
  générés avant ce changement) ; pour les horaires agents, cherche la grille
  par contenu en priorité, repli sur `_prep_Horaires_Des_Agents` (liste à
  plat) si absente. Le calcul de `manquants` a été ajusté en conséquence
  (ne signale plus `Horaires_Des_Agents` comme manquant si la grille est
  trouvée).

**`app.py`** : message d'aide du Bloc 4 mis à jour — Paramètres, Affectations
ET "horaires d'équipes" présentés comme directement modifiables (pas besoin
de déverrouiller) ; seul 'Planning_type' reste mentionné dans la partie
"à déverrouiller d'abord" (cohérent avec §25.3 : lui seul reste verrouillé).

### 25.5 Piège hachures — revécu, pas juste documenté

Le fichier `horaires_agents.xlsx` fourni par Elo pour cette session contient
28 cellules en hachures (`patternType='lightUp'`, marquage des demi-journées
non travaillées) ET 105 formules (totaux jour/semaine). Passer l'ensemble du
classeur par `recalc.py` (LibreOffice, nécessaire pour restaurer les valeurs
calculées des formules du reste du fichier — Semaine_1..5 etc, qui perdent
leur cache à chaque sauvegarde openpyxl) aplatit ces hachures en gris plat
uni — exactement le piège déjà documenté en §24.2, mais qui n'avait jusqu'ici
été résolu que pour le fichier `horaires_agents_harmonise.xlsx` **seul**
(sans autre onglet à formules dans le même classeur). Cette fois, impossible
d'éviter complètement `recalc.py` (les formules Semaine_X en ont besoin) —
solution appliquée : recalculer le classeur complet, puis **patcher
directement le XML** (`xl/styles.xml`) après coup pour ré-écrire la seule
définition de remplissage concernée (retour de
`patternType="solid"` + gris vers `patternType="lightUp"` +
`FF000000`/`FFD9E1F2`), sans repasser par openpyxl (qui aurait de nouveau
effacé le cache des formules). Combiné dans la même passe XML que la
restauration habituelle du `veryHidden` (aussi cassé par LibreOffice, cf.
pratique déjà en place). À généraliser : **tout classeur combinant à la fois
des formules à recalculer ET des remplissages non-solides quelque part**
demande cette double correction post-`recalc.py` (état des onglets +
styles), pas seulement l'une ou l'autre.

### 25.6 Fichiers livrés cette session

- `generate_planning_excel_septembre.py` (§25.4, puis §25.8)
- `planning_checker.py` (§25.4)
- `app.py` (message d'aide Bloc 4, mis à jour et complété en §25.4)
- `regeneration_lecture.py` (message d'erreur reformulé, cosmétique)
- `Planning_Septembre_2026.xlsx` (regénéré avec Paramètres/Affectations/
  horaires d'équipes visibles et modifiables, hachures et onglets
  très masqués intacts, dates lisibles, formules de la grille verrouillées,
  contenu du planning lui-même strictement inchangé)

### 25.7 Reste à faire

- Décider si `Planning_type` doit lui aussi devenir modifiable un jour, ou
  s'il reste volontairement en lecture seule (pas demandé à ce stade).
- Vérifier que le repli `_prep_Paramètres`/`_prep_Affectations` fonctionne
  bien sur un vieux fichier (mai, octobre) lors d'une prochaine régénération
  — testé seulement en lecture de code cette session, pas sur un fichier
  réel antérieur à ce changement.
- Élucider si possible pourquoi le correctif décrit en §24.5 n'était pas
  présent dans le code réellement en usage (cf. §25.2) — pour éviter de
  refaire la même désynchronisation doc/code.
### 25.8 Correctifs du 24/08 (suite) — affichage dates Paramètres + verrouillage formules horaires d'équipes

Deux petits défauts remontés par Elo après avoir testé le fichier livré en
§25.6, corrigés à la fois sur `Planning_Septembre_2026.xlsx` **et** dans le
code (pour que ça ne se reproduise pas sur les prochaines générations) :

**1. Dates en `#######` dans Paramètres (colonne B, tableau Présence
Vacataire).** Cause : l'ancien pipeline (`copier_onglets_preparation_caches`,
copie valeurs seules) ne recopiait pas le format numérique d'origine ; les
dates récupéraient le format par défaut d'openpyxl
(`yyyy-mm-dd h:mm:ss`, 19 caractères), trop large pour une colonne de 13 —
défaut invisible tant que l'onglet était très masqué, visible seulement
maintenant qu'il ne l'est plus. Note : `_embarquer_prep_visible_modifiable`
(§25.4, via `_copier_feuille_avec_mise_en_forme`) recopie déjà le format
d'origine cellule par cellule, donc ce bug précis ne se serait normalement
pas reproduit sur une génération neuve — mais rien ne garantit que le
format du fichier de préparation d'Elo soit toujours propre (fichier
retouché à la main d'un mois sur l'autre). **Correctif** : nouvelle fonction
`_normaliser_dates_visibles(ws)`, appelée systématiquement à la fin de
`_embarquer_prep_visible_modifiable` (donc pour Paramètres, Affectations
ET horaires d'équipes) — force `d-mmm-yy` (convention déjà utilisée par
Elo, ex. `5-sept-26`) sur toute cellule contenant une vraie date
(`datetime.date`/`datetime.datetime`, jamais `datetime.time` ni
`timedelta` — donc aucun effet sur les colonnes d'horaires de la grille),
et élargit à 16 la colonne concernée si elle est plus étroite.

**2. Grille "horaires d'équipes" pas assez verrouillée.** Toutes les
cellules étaient déverrouillées (cohérent avec la demande du 24/08 matin —
pouvoir corriger un horaire), mais cela incluait aussi les ~105 cellules de
formule (totaux jour/semaine), exposées à un écrasement accidentel en
tapant à côté d'une case de saisie. **Correctif** :
`_embarquer_prep_visible_modifiable` accepte désormais un paramètre
`verrouiller_formules` (défaut `False`) ; `embarquer_horaires_agents_visible`
l'appelle avec `verrouiller_formules=True` — après le déverrouillage
général, les cellules dont la valeur est une formule (`ArrayFormula` ou
chaîne commençant par `=`) sont reverrouillées. Sans effet sur
Paramètres/Affectations (aucune formule dans ces deux onglets).

**Fichiers retouchés** : `generate_planning_excel_septembre.py` (les deux
correctifs) ; `Planning_Septembre_2026.xlsx` regénéré avec les deux
correctifs appliqués manuellement (même résultat que ce que produirait
désormais le code), en repassant par le même protocole habituel
(`recalc.py` puis double restauration XML `veryHidden` + hachures, cf.
§25.5) — testé et vérifié : 105 cellules verrouillées, 28 hachures
intactes, dates lisibles avec largeur de colonne correcte, valeurs
calculées et contenu du planning inchangés par ailleurs.

`planning_checker.py` et `app.py` : **non retouchés** cette fois — ces deux
défauts concernent uniquement l'affichage/la protection du fichier généré,
pas la lecture des données par le Bloc 3/4.

### 25.9 Premier vrai test du Bloc 4 par Elo — bug de casse sur 'Planning_type'

Premier test réel (pas juste relecture de code) du Bloc 4 par Elo, avec le
fichier de septembre livré en §25.8 : erreur *"Onglet(s) de préparation
manquant(s) : Planning_type"* alors que l'onglet était bien présent.

**Cause** : l'onglet s'appelle en réalité `planning_type`, tout en
minuscules, dans ce fichier — pas `Planning_type`. `ONGLETS_PREP_SANS_PREFIXE`
et le test `nom in wb.sheetnames` de `charger_donnees_preparation()` sont
sensibles à la casse, donc ne le trouvaient pas. Ce n'est pas une
régression de cette session : ce bug existait déjà dans le code d'origine
(avant toute intervention de Claude), simplement jamais testé en
conditions réelles jusqu'ici — la preuve : `embarquer_planning_type_visible()`
dans `generate_planning_excel_septembre.py` tolère déjà les deux casses
en LECTURE du fichier de préparation
(`raw.get('Planning_type') or raw.get('planning_type')`), signe que ce
problème de casse était déjà connu d'un côté du pipeline, mais pas de
l'autre. Le fichier de test d'Elo a probablement été généré par une version
du générateur antérieure à l'harmonisation en casse capitalisée.

**Correctif** (`planning_checker.py` uniquement) : nouvelle fonction
`_trouver_onglet_insensible_casse(wb, nom)`, recherche insensible à la
casse dans `wb.sheetnames` ; utilisée dans `charger_donnees_preparation()`
à la place du test `nom in wb.sheetnames` pour les onglets de
`ONGLETS_PREP_SANS_PREFIXE` (Planning_type, Paramètres, Affectations).
Testé directement sur le fichier `TEST-Planning_Septembre_2026-24_08.xlsx`
fourni par Elo : les 3 onglets sont désormais trouvés, plus aucun
"manquant".

**Observation annexe, non résolue** : dans ce même fichier de test,
`Paramètres` et `Affectations` sont à l'état `hidden` (masqué simple) plutôt
que `visible` — alors qu'ils avaient été livrés visibles. Sans effet sur
le bug ci-dessus (la détection d'onglet ne dépend pas de l'état), donc pas
creusé cette session ; cause possible : réenregistrement du fichier par un
tableur (Excel/LibreOffice) entre la livraison et le test. À surveiller si
Elo signale ne plus voir ces onglets par défaut à l'ouverture.

**Fichier livré** : `planning_checker.py` uniquement — pas de nouvelle
version de `Planning_Septembre_2026.xlsx` cette fois (le fichier d'Elo n'a
pas été modifié, seule la lecture côté app était en cause).

### 25.10 Généralisation demandée par Elo — insensibilité à la casse partout

Elo a demandé explicitement, après le correctif ponctuel du §25.9 : *"il ne
faut pas que la lecture soit sensible à la casse, majuscule = minuscule
dans tout le fichier"*. Vérification faite : le correctif précédent ne
touchait QUE la recherche de `Planning_type`/`Paramètres`/`Affectations`
dans `charger_donnees_preparation()` — plusieurs autres recherches
d'onglets, dans les mêmes fichiers, restaient sensibles à la casse.
Recensement complet et correction :

- **`planning_checker.py`** : repli `_prep_...` (Paramètres/Affectations/
  Planning_type et Horaires_Des_Agents) désormais aussi insensible à la
  casse ; recherche des onglets `Semaine_N` (`re.match` avec
  `re.IGNORECASE` ajouté) et `Semaine_N_Agent` (via
  `_trouver_onglet_insensible_casse`, nouvelle fonction, exportée pour
  réutilisation par les 2 fichiers suivants).
- **`regeneration_lecture.py`** : recherche de l'onglet `Semaine_N` à
  régénérer, désormais insensible à la casse (importe
  `_trouver_onglet_insensible_casse` depuis `planning_checker.py`).
- **`regeneration_ecriture.py`** : recherche de l'onglet `Semaine_N` (pour
  y écrire les jours régénérés) et de l'ancien onglet `Semaine_N_Agent`
  (pour le supprimer avant reconstruction), toutes deux insensibles à la
  casse désormais. **Nuance volontaire** : le nom du NOUVEL onglet
  "vue par agent" que `generer_vue_agent()` crée reste écrit avec la casse
  canonique (`Semaine_N_Agent`, majuscules) — l'insensibilité à la casse
  s'applique à la LECTURE (retrouver un onglet existant, quelle que soit
  sa casse), jamais à l'ÉCRITURE (un fichier fraîchement généré a toujours
  une casse cohérente et connue).

**Périmètre non couvert par ce correctif, précisé à Elo** : les
comparaisons de CONTENU (noms de section RDC/Adulte/M&F/Jeunesse, jours de
la semaine, catégories d'agent...) passent déjà, depuis l'origine, par
`normalize()`/`canon_section()` (minuscules + accents retirés) — donc déjà
insensibles à la casse, rien à corriger là. Seuls les noms d'ONGLET
posaient problème. Non touché non plus, faute d'accès au fichier :
`planning_engine_cpsat.py` (`load_excel_data`, lecture du fichier de
préparation d'origine, pas du planning déjà généré) — à vérifier lors
d'une prochaine session si ce fichier est fourni.

**Fichiers livrés** : `planning_checker.py`, `regeneration_lecture.py`,
`regeneration_ecriture.py`. Testé sur `TEST-Planning_Septembre_2026-24_08.xlsx` :
tous les onglets (Semaine_1, Semaine_1_Agent, planning_type minuscule,
_prep_Roulement_Samedi, Paramètres, Affectations) retrouvés correctement.

