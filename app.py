"""
Application web — Planning Médiathèque
Déployée sur Streamlit Cloud
"""

import streamlit as st
import io
from planning_engine import compute_full_planning
from excel_writer import generate_excel

# ─────────────────────────────────────────────
#  CONFIGURATION PAGE
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Planning Médiathèque",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS personnalisé
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 800;
        color: #1A2F4A;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        font-size: 1rem;
        color: #7F8C8D;
        margin-bottom: 2rem;
    }
    .step-box {
        background: #EBF5FB;
        border-left: 4px solid #2E86C1;
        padding: 1rem 1.2rem;
        border-radius: 6px;
        margin-bottom: 1rem;
    }
    .step-num {
        font-size: 1.3rem;
        font-weight: 700;
        color: #2E86C1;
    }
    .success-box {
        background: #D5F5E3;
        border-left: 4px solid #1E8449;
        padding: 1rem 1.2rem;
        border-radius: 6px;
    }
    .warning-box {
        background: #FEF9E7;
        border-left: 4px solid #F39C12;
        padding: 1rem 1.2rem;
        border-radius: 6px;
    }
    .info-card {
        background: white;
        border: 1px solid #D5D8DC;
        border-radius: 8px;
        padding: 1.2rem;
        margin: 0.5rem 0;
    }
    div[data-testid="stDownloadButton"] button {
        background-color: #1E8449 !important;
        color: white !important;
        font-size: 1.1rem !important;
        padding: 0.6rem 2rem !important;
        border-radius: 8px !important;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  SIDEBAR — GUIDE
# ─────────────────────────────────────────────

with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/library.png", width=80)
    st.markdown("## 📖 Guide d'utilisation")
    st.markdown("""
    **Chaque mois, en 4 étapes :**

    **① Préparez votre Excel — chaque mois**
    - Mettez à jour l'onglet `Événements` (congés, réunions, absences)
    - Dans `Paramètres` : changez **Mois**, **Année**
    - Dans `Paramètres` : mettez à jour **Samedi_S1 à Samedi_S5** (ROUGE ou BLEU pour chaque samedi du mois) ⚠️ indispensable
    - Mettez à jour `SP_MinMax` si les quotas SP changent
    - Ajustez `Horaires_Des_Agents` si des horaires ont changé
    - Vérifiez `Roulement_Samedi` si l'équipe évolue

    **② Uploadez le fichier**
    - Cliquez sur *Browse files*
    - Sélectionnez votre fichier `.xlsx`

    **③ Générez**
    - Cliquez sur **Générer le planning**
    - L'algorithme calcule les 4 semaines

    **④ Téléchargez**
    - Récupérez le fichier Excel avec les plannings remplis
    - Les onglets `Semaine_1` à `Semaine_4` sont mis à jour

    ---
    **❓ Règles appliquées**
    - ✅ Horaires de chaque agent respectés
    - ✅ Événements / congés pris en compte
    - ✅ Besoins Jeunesse par créneau
    - ✅ Roulement samedi Rouge/Bleu
    - ✅ Pas de vacataire seul (sauf 12h-14h)
    - ✅ Sections autorisées par agent
    - ✅ Vacataires uniquement mercredi et samedi
    """)

    st.markdown("---")
    st.markdown("**Version :** 1.0  \n**Médiathèque** — Planning SP")


# ─────────────────────────────────────────────
#  PAGE PRINCIPALE
# ─────────────────────────────────────────────

st.markdown('<div class="main-title">📅 Planning Service Public — Médiathèque</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Générez automatiquement le planning mensuel à partir de votre fichier Excel</div>', unsafe_allow_html=True)

# ── ÉTAPE 1 : UPLOAD ──
st.markdown("---")
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📂 Étape 1 — Uploadez votre fichier Excel")
    st.markdown("Choisissez votre fichier `Solveur_Planning_Vxx.xlsx` avec les données du mois à planifier.")

    uploaded_file = st.file_uploader(
        "Glissez-déposez ou cliquez pour choisir",
        type=['xlsx', 'xls'],
        help="Le fichier doit contenir les onglets : Paramètres, Événements, Horaires_Des_Agents, Affectations, etc."
    )

with col2:
    st.markdown("### 💡 Données attendues")
    st.markdown("""
    <div class="info-card">
    Votre fichier doit contenir :<br><br>
    📋 <b>Paramètres</b> — mois, année<br>
    📅 <b>Événements</b> — congés, réunions<br>
    🕐 <b>Horaires_Des_Agents</b><br>
    👥 <b>Affectations</b> — sections<br>
    📊 <b>SP_MinMax</b> — quotas SP<br>
    🔄 <b>Roulement_Samedi</b><br>
    👶 <b>Besoins_Jeunesse</b>
    </div>
    """, unsafe_allow_html=True)


# ── ÉTAPE 2 : APERÇU DES DONNÉES ──
if uploaded_file is not None:
    st.markdown("---")
    st.markdown("### 🔍 Étape 2 — Vérification des données")

    try:
        file_bytes = uploaded_file.read()
        file_buf   = io.BytesIO(file_bytes)

        with st.spinner("Lecture du fichier Excel..."):
            from planning_engine import (
                load_excel_data, parse_parametres, parse_affectations,
                parse_roulement_samedi, parse_evenements, get_weeks_of_month,
                compute_full_planning
            )
            raw      = load_excel_data(file_buf)
            params   = parse_parametres(raw)
            affecta  = parse_affectations(raw)
            roulem   = parse_roulement_samedi(raw)

        col1, col2, col3 = st.columns(3)

        with col1:
            mois  = str(params.get('Mois', '?')).capitalize()
            annee = int(params.get('Année', 2026))
            st.metric("📅 Mois planifié", f"{mois} {annee}")

        with col2:
            nb_agents = len([a for a in affecta if 'vacataire' not in a.lower()])
            st.metric("👥 Agents permanents", nb_agents)

        with col3:
            nb_vac = len([a for a in affecta if 'vacataire' in a.lower()])
            st.metric("🔄 Vacataires", nb_vac)

        # Tableau des agents
        with st.expander("👥 Voir la liste des agents et leurs sections", expanded=False):
            import pandas as pd
            rows = []
            for agent, sects in affecta.items():
                rouge_bleu = roulem.get(agent, '—')
                rows.append({
                    'Agent': agent,
                    'Section principale': sects[0] if sects else '?',
                    'Sections secondaires': ', '.join(sects[1:]) if len(sects) > 1 else '—',
                    'Samedi': rouge_bleu,
                })
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

        # Événements du mois
        MOIS_MAP = {
            'janvier':1,'février':2,'mars':3,'avril':4,'mai':5,'juin':6,
            'juillet':7,'août':8,'septembre':9,'octobre':10,'novembre':11,'décembre':12
        }
        mois_num = MOIS_MAP.get(str(params.get('Mois','')).lower(), 1)

        evts = parse_evenements(raw, mois_num, annee)
        total_evts = sum(len(v) for v in evts.values())

        if total_evts > 0:
            with st.expander(f"📌 Événements du mois ({total_evts} événements)", expanded=False):
                rows_ev = []
                for date_str, ev_list in sorted(evts.items()):
                    for ev in ev_list:
                        from planning_engine import min_to_hhmm
                        rows_ev.append({
                            'Date': date_str,
                            'Début': min_to_hhmm(ev['debut']),
                            'Fin':   min_to_hhmm(ev['fin']),
                            'Événement': ev['nom'],
                            'Agents': ', '.join(ev['agents']) if ev['agents'] else '—',
                        })
                df_ev = pd.DataFrame(rows_ev)
                st.dataframe(df_ev, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ Aucun événement trouvé pour ce mois dans l'onglet Événements.")

        # ── ÉTAPE 3 : GÉNÉRATION ──
        st.markdown("---")
        st.markdown("### ⚙️ Étape 3 — Génération du planning")

        col_btn, col_info = st.columns([1, 2])

        with col_btn:
            generate = st.button(
                "🚀 Générer le planning",
                type="primary",
                use_container_width=True,
            )

        with col_info:
            st.markdown("""
            <div class="warning-box">
            ⏱️ La génération prend environ <b>5 à 15 secondes</b>.<br>
            Le planning respecte toutes vos règles et contraintes.
            </div>
            """, unsafe_allow_html=True)

        if generate:
            with st.spinner("⚙️ Calcul du planning en cours... Analyse des disponibilités, contraintes, événements..."):
                try:
                    file_buf2 = io.BytesIO(file_bytes)
                    weeks_data, metadata = compute_full_planning(file_buf2)

                    file_buf3 = io.BytesIO(file_bytes)
                    output_buf = generate_excel(file_buf3, weeks_data, metadata)

                    st.markdown("---")
                    st.markdown("### ✅ Étape 4 — Téléchargement")

                    st.markdown("""
                    <div class="success-box">
                    ✅ <b>Planning généré avec succès !</b><br>
                    Les 4 semaines ont été calculées en respectant toutes les contraintes.
                    Cliquez sur le bouton pour télécharger votre fichier Excel.
                    </div>
                    """, unsafe_allow_html=True)

                    # Résumé par semaine
                    for wd in weeks_data:
                        wn   = wd['week_num']
                        wdat = wd['week_dates']
                        sam  = wd['samedi_type']
                        dates = sorted([d for d in wdat.values() if d is not None])
                        label = f"{dates[0].day} au {dates[-1].day}" if dates else "?"
                        col_a, col_b, col_c = st.columns([1, 2, 1])
                        with col_a:
                            st.markdown(f"**Semaine {wn}**")
                        with col_b:
                            mois_cap = str(metadata['mois']).capitalize()
                            st.markdown(f"{label} {mois_cap} {metadata['annee']}")
                        with col_c:
                            badge = "🔴 ROUGE" if sam == 'ROUGE' else "🔵 BLEU"
                            st.markdown(f"Samedi : {badge}")

                    st.markdown("<br>", unsafe_allow_html=True)

                    mois_cap = str(metadata['mois']).capitalize()
                    filename = f"Planning_{mois_cap}_{metadata['annee']}.xlsx"

                    st.download_button(
                        label=f"⬇️  Télécharger le planning — {mois_cap} {metadata['annee']}",
                        data=output_buf.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération : {str(e)}")
                    with st.expander("Détails de l'erreur"):
                        import traceback
                        st.code(traceback.format_exc())

    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier : {str(e)}")
        with st.expander("Détails"):
            import traceback
            st.code(traceback.format_exc())

else:
    # Placeholder quand pas de fichier
    st.markdown("""
    <br>
    <div style="text-align:center; padding: 3rem; background:#F8F9FA; border-radius:12px; border: 2px dashed #D5D8DC;">
        <div style="font-size:4rem">📂</div>
        <div style="font-size:1.2rem; color:#7F8C8D; margin-top:1rem">
            Uploadez votre fichier Excel pour commencer
        </div>
        <div style="font-size:0.9rem; color:#AAB7B8; margin-top:0.5rem">
            Format accepté : .xlsx
        </div>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#AAB7B8; font-size:0.8rem'>"
    "Planning Médiathèque — Généré automatiquement • "
    "Toutes les règles de planification sont appliquées"
    "</div>",
    unsafe_allow_html=True
)
