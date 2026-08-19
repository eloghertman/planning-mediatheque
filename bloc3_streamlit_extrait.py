"""
Extrait à intégrer dans app.py — Bloc 3 : Vérifier un planning déjà rempli.

⚠️ Le app.py actuellement dans le projet importe encore planning_engine /
excel_writer (l'ancienne version), alors que le fichier de contexte v30
indique que app.py a été réécrit le 18/08 pour ne plus utiliser que
planning_engine_cpsat / generate_planning_excel_septembre. Le app.py du
projet semble donc être une version plus ancienne que celle réellement
utilisée. Collez ce bloc dans VOTRE app.py actuel (celui utilisé pour de
vrai), à la suite du bloc 2 — ne remplacez pas tout le fichier avec ceci.

Ajoutez en haut du fichier :
    from planning_checker import verifier_planning, resumer
"""

import streamlit as st
from planning_checker import verifier_planning, resumer


def afficher_bloc_verification():
    st.markdown("---")
    st.markdown("### 🔎 Étape — Vérifier un planning déjà rempli")
    st.markdown(
        "Déposez un planning déjà généré (et éventuellement modifié à la main "
        "par vous ou par les agents). L'app relit toutes les cases et signale "
        "les contradictions avec les règles obligatoires."
    )

    fichier_verif = st.file_uploader(
        "Fichier Excel du planning à vérifier",
        type=['xlsx'],
        key="uploader_verification",
    )

    if fichier_verif is not None:
        with st.spinner("Vérification en cours..."):
            try:
                anomalies = verifier_planning(fichier_verif.read())
            except Exception as e:
                st.error(f"❌ Impossible de vérifier ce fichier : {str(e)}")
                with st.expander("Détails de l'erreur"):
                    import traceback
                    st.code(traceback.format_exc())
                return

        n_rouge, n_jaune = resumer(anomalies)

        if not anomalies:
            st.success("✅ Aucune anomalie détectée sur les règles vérifiées.")
            return

        col1, col2 = st.columns(2)
        with col1:
            st.metric("🔴 Impossibilités", n_rouge)
        with col2:
            st.metric("🟡 À vérifier", n_jaune)

        # Regroupement par semaine puis par jour
        semaines = {}
        for a in anomalies:
            semaines.setdefault(a.semaine, {}).setdefault(a.jour or '(général)', []).append(a)

        for semaine, jours in semaines.items():
            with st.expander(f"📅 {semaine}", expanded=True):
                for jour, liste in jours.items():
                    st.markdown(f"**{jour}**" if jour else "**Général**")
                    for a in sorted(liste, key=lambda x: 0 if x.gravite == 'rouge' else 1):
                        icone = "🔴" if a.gravite == 'rouge' else "🟡"
                        st.markdown(f"{icone} {a.message}")


# Dans le corps principal de app.py, après le bloc 2 (génération) :
# afficher_bloc_verification()
