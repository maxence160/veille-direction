import html
import io
import json
import os
import urllib.parse
from datetime import datetime
import feedparser
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Configuration de la page
st.set_page_config(
    page_title="Présidentielle 2027 — La course, en un coup d'œil",
    page_icon="⚖️",
    layout="wide",
)

# Injection de styles pour reproduire la DA éditoriale / presse
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,600;0,700;1,400;1,600&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    /* Fond général crème et police principale */
    .stApp {
        background-color: #F4F1EA;
        color: #1C1917;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    /* En-tête éditorial */
    .top-meta {
        display: flex;
        justify-content: space-between;
        font-size: 0.85rem;
        color: #57534E;
        margin-bottom: 18px;
    }
    .header-title {
        font-family: 'Lora', serif;
        font-size: 3.2rem;
        font-weight: 700;
        line-height: 1.15;
        letter-spacing: -0.5px;
        color: #1C1917;
        margin-bottom: 8px;
    }
    .header-title em {
        font-style: italic;
        font-weight: 600;
        color: #8B261E;
    }
    .header-lead {
        font-size: 1.05rem;
        color: #57534E;
        max-width: 780px;
        margin-bottom: 20px;
        line-height: 1.5;
    }
    .milestones {
        display: flex;
        gap: 24px;
        flex-wrap: wrap;
        font-size: 0.9rem;
        color: #57534E;
        padding-bottom: 16px;
    }
    .milestones strong {
        color: #1C1917;
        font-weight: 700;
    }
    .main-separator {
        border-top: 2px solid #1C1917;
        margin-bottom: 24px;
    }

    /* Encadré d'avertissement institutionnel */
    .callout {
        background-color: rgba(255, 255, 255, 0.55);
        border: 1px solid #D6D3CD;
        border-left: 4px solid #9A7B56;
        border-radius: 4px;
        padding: 16px 20px;
        font-size: 0.88rem;
        color: #1C1917;
        line-height: 1.6;
        margin-bottom: 28px;
    }

    /* Fiches candidats */
    .candidate-box {
        background: rgba(255, 255, 255, 0.6);
        border: 1px solid #D6D3CD;
        border-radius: 4px;
        padding: 16px 18px;
        margin-bottom: 12px;
    }
    .candidate-name {
        font-family: 'Lora', serif;
        font-weight: 700;
        font-size: 1.15rem;
        color: #1C1917;
        margin-bottom: 4px;
    }
    .candidate-role {
        font-size: 0.88rem;
        color: #57534E;
    }

    /* Boutons et éléments d'interaction */
    div.stButton > button:first-child {
        background-color: #1C1917 !important;
        color: #F4F1EA !important;
        border: none !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        padding: 8px 18px !important;
        transition: background-color 0.2s ease;
    }
    div.stButton > button:first-child:hover {
        background-color: #8B261E !important;
        color: #FFFFFF !important;
    }

    /* Style des onglets Streamlit */
    .stTabs [data-baseweb="tab-list"] {
        border-bottom: 1px solid #1C1917;
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Lora', serif !serif;
        font-size: 1.1rem;
        font-weight: 600;
        color: #57534E;
        background-color: transparent;
        padding: 10px 20px;
        border: 1px solid transparent;
        border-bottom: none;
        border-radius: 4px 4px 0 0;
    }
    .stTabs [aria-selected="true"] {
        color: #1C1917 !important;
        background-color: #F4F1EA !important;
        border-color: #1C1917 !important;
        border-bottom: 1px solid #F4F1EA !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

DB_FILE = "veille_data.json"
SEEN_FILE = "deja_vus.json"

# Requêtes ciblées : présidentielle, propositions législatives et actualité agro-alimentaire
QUERIES = [
    'présidentielle "proposition de loi"',
    'présidentielle programme candidat',
    'présidentielle "agroalimentaire" OR "alimentation"',
    'présidentielle fiscalité agriculture',
]


def charger_donnees():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def sauvegarder_donnees(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def charger_vus():
    if os.path.exists(SEEN_FILE):
        with open(SEEN_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def sauvegarder_vus(vus):
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(list(vus), f, ensure_ascii=False, indent=2)


def lancer_actualisation():
    deja_vus = charger_vus()
    donnees_existantes = charger_donnees()
    nouveaux_ajouts = 0

    for q in QUERIES:
        query_encoded = urllib.parse.quote(q)
        url_flux = f"https://news.google.com/rss/search?q={query_encoded}&hl=fr&gl=FR&ceid=FR:fr"
        flux = feedparser.parse(url_flux)

        for entry in flux.entries:
            lien = entry.link
            if lien not in deja_vus:
                deja_vus.add(lien)
                source = (
                    entry.source.get("title", "Presse")
                    if hasattr(entry, "source")
                    else "Presse"
                )
                date_pub = entry.get(
                    "published", datetime.now().strftime("%d %b %Y %H:%M")
                )
                titre_propre = html.unescape(entry.title)

                donnees_existantes.insert(
                    0,
                    {
                        "Date": date_pub,
                        "Source": source,
                        "Titre": titre_propre,
                        "Lien": lien,
                    },
                )
                nouveaux_ajouts += 1

    sauvegarder_vus(deja_vus)
    sauvegarder_donnees(donnees_existantes)
    return nouveaux_ajouts


def exporter_excel(df_export):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Veille Législative"

    headers = [
        "Date",
        "Média / Source",
        "Proposition / Annonce de loi",
        "Lien de consultation",
    ]
    ws.append(headers)

    header_fill = PatternFill(
        start_color="1C1917", end_color="1C1917", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D6D3CD"),
        right=Side(style="thin", color="D6D3CD"),
        top=Side(style="thin", color="D6D3CD"),
        bottom=Side(style="thin", color="D6D3CD"),
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row in df_export[["Date", "Source", "Titre", "Lien"]].itertuples(
        index=False
    ):
        ws.append(list(row))

    for r_idx in range(2, len(df_export) + 2):
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if c_idx == 4:
                cell.font = Font(color="8B261E", underline="single")
                cell.hyperlink = cell.value

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 70)

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()


# --- Structure Visuelle Supérieure ---
st.markdown(
    f"""
<div class="top-meta">
    <span>Répertoire indépendant, non affilié</span>
    <span>{datetime.now().strftime('%d %B %Y')}</span>
</div>
<h1 class="header-title">Présidentielle 2027<br>— <em>la course, en un coup d'œil</em></h1>
<p class="header-lead">Un point de repère sur les candidatures déclarées et les propositions de loi qui font l'actualité, avec des liens vers les sources pour aller plus loin.</p>
<div class="milestones">
    <span><strong>1er tour :</strong> 18 avril 2027</span>
    <span><strong>2d tour :</strong> 2 mai 2027</span>
    <span><strong>Seuil d'entrée :</strong> 500 parrainages d'élus, 30 départements min.</span>
</div>
<div class="main-separator"></div>
<div class="callout">
    <strong>Ceci est un instantané de veille continue.</strong> La page reprend l'état des candidatures et une sélection de propositions de loi relayées dans la presse et les comptes-rendus publics (Assemblée nationale, Sénat, presse spécialisée et nationale).
</div>
""",
    unsafe_allow_html=True,
)

# Onglets
tab_candidats, tab_lois = st.tabs(["Candidats", "Propositions de loi"])

# --- Onglet 1 : Candidats ---
with tab_candidats:
    st.markdown(
        """
    <h3 style="font-family:'Lora',serif; font-size:1.5rem; font-weight:700; margin-bottom:6px;">Candidatures officialisées & déclarées</h3>
    <p style="color:#57534E; font-size:0.95rem; margin-bottom:20px;">Panorama des personnalités ayant engagé une démarche officielle ou confirmé leur candidature.</p>
    """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    candidats = [
        (
            "Édouard Philippe",
            "Horizons",
            "Candidature officialisée en septembre 2024",
        ),
        (
            "Marine Le Pen",
            "Rassemblement National",
            "Candidature annoncée pour le mouvement",
        ),
        (
            "Jean-Luc Mélenchon",
            "La France Insoumise",
            "Hypothèse de candidature réitérée",
        ),
        (
            "Laurent Wauquiez",
            "La Droite Républicaine",
            "Chef de file à l'Assemblée nationale",
        ),
        (
            "Fabien Roussel",
            "Parti Communiste Français",
            "Démarche autonome confirmée",
        ),
        (
            "François Ruffin",
            "Debout !",
            "Positionnement pour un rassemblement populaire",
        ),
    ]

    for idx, (nom, parti, statut) in enumerate(candidats):
        cible = col1 if idx % 2 == 0 else col2
        with cible:
            st.markdown(
                f"""
            <div class="candidate-box">
                <div class="candidate-name">{nom}</div>
                <div class="candidate-role"><strong>{parti}</strong> — {statut}</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

# --- Onglet 2 : Propositions de loi ---
with tab_lois:
    col_act, col_recherche, col_exp = st.columns([1.2, 2.5, 1.3])

    with col_act:
        if st.button("🔄 Actualiser la veille", use_container_width=True):
            with st.spinner("Récupération des flux..."):
                nb = lancer_actualisation()
                if nb > 0:
                    st.success(f"{nb} nouveau(x) texte(s) ajouté(s).")
                else:
                    st.info("Aucune nouvelle parution.")

    donnees = charger_donnees()
    df = pd.DataFrame(donnees) if donnees else pd.DataFrame()

    with col_recherche:
        terme = st.text_input(
            "Filtrer les annonces :",
            placeholder="Mots-clés (ex: alimentation, impôt, régulation...)",
            label_visibility="collapsed",
        )

    if not df.empty:
        if terme:
            df = df[
                df["Titre"].str.contains(terme, case=False, na=False)
                | df["Source"].str.contains(terme, case=False, na=False)
            ]

        with col_exp:
            fichier_excel = exporter_excel(df)
            st.download_button(
                label="📥 Télécharger l'Excel",
                data=fichier_excel,
                file_name=f"veille_lois_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.caption(f"{len(df)} proposition(s) ou texte(s) répertorié(s)")

        st.dataframe(
            df[["Date", "Source", "Titre", "Lien"]],
            column_config={
                "Lien": st.column_config.LinkColumn(
                    "Source", display_text="Consulter le texte ↗"
                ),
                "Date": st.column_config.TextColumn(
                    "Date de parution", width="small"
                ),
                "Source": st.column_config.TextColumn("Média", width="small"),
                "Titre": st.column_config.TextColumn(
                    "Proposition de loi / Annonce", width="large"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.info(
            "La base de données est vide. Cliquez sur 'Actualiser la veille'"
            " pour lancer l'aspiration des premières propositions de loi."
        )
