import html
import io
import re
import unicodedata
import urllib.parse
from datetime import datetime
import dateutil.parser
import feedparser
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Présidentielle 2027 — La course, en un coup d'œil",
    page_icon="⚖️",
    layout="wide",
)

DATE_MIN = datetime(2026, 1, 1, 0, 0, 0)

# Référentiel strict : Prénom + Nom complet obligatoire pour éviter les confusions
REFERENTIEL_CANDIDATS = [
    ("Édouard Philippe", ["edouard philippe"], "Horizons"),
    ("Marine Le Pen", ["marine le pen"], "Rassemblement National"),
    ("Jordan Bardella", ["jordan bardella"], "Rassemblement National"),
    (
        "Jean-Luc Mélenchon",
        ["jean-luc melenchon", "jean luc melenchon"],
        "La France Insoumise",
    ),
    ("François Ruffin", ["francois ruffin"], "Debout !"),
    ("Laurent Wauquiez", ["laurent wauquiez"], "La Droite Républicaine"),
    ("Gabriel Attal", ["gabriel attal"], "Renaissance"),
    ("Gérald Darmanin", ["gerald darmanin"], "Renaissance"),
    ("Fabien Roussel", ["fabien roussel"], "Parti Communiste Français"),
    ("Bernard Cazeneuve", ["bernard cazeneuve"], "La Convention"),
    ("David Lisnard", ["david lisnard"], "Nouvelle Énergie"),
    ("Marine Tondelier", ["marine tondelier"], "Les Écologistes"),
]

# Requêtes ciblées : Présidentielle obligatoire
QUERIES = [
    (
        'présidentielle ("taxe soda" OR "taxe sucre" OR "boissons sucrées" OR'
        ' fiscalité OR impôt OR TVA) after:2026-01-01'
    ),
    (
        'présidentielle (consigne OR "bouteilles plastique" OR réemploi OR'
        " emballage OR eau OR \"prélèvements d'eau\" OR écologie)"
        " after:2026-01-01"
    ),
    (
        'présidentielle (EGAlim OR Descrozaille OR "négociations commerciales"'
        ' OR "grande distribution" OR marges OR "prix planchers")'
        " after:2026-01-01"
    ),
    (
        'présidentielle (Nutri-score OR nutrition OR "santé publique" OR sucre'
        ' OR obésité OR "recettes allégées") after:2026-01-01'
    ),
    (
        'présidentielle ("publicité alimentaire" OR "publicité enfants" OR'
        ' "marketing alimentaire" OR parrainage) after:2026-01-01'
    ),
    (
        'présidentielle (programme OR proposition OR annonce OR réforme)'
        " agroalimentaire after:2026-01-01"
    ),
    (
        'présidentielle (programme OR proposition OR annonce) ("Édouard'
        ' Philippe" OR "Marine Le Pen" OR "Jean-Luc Mélenchon" OR "Laurent'
        ' Wauquiez" OR "Gabriel Attal" OR "François Ruffin") after:2026-01-01'
    ),
]

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,600;0,700;1,400;1,600&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    .stApp {
        background-color: #F4F1EA;
        color: #1C1917;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

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
        max-width: 820px;
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

    /* Carte candidat simplifiée sans citations */
    .candidate-card {
        background: rgba(255, 255, 255, 0.65);
        border: 1px solid #D6D3CD;
        border-left: 4px solid #8B261E;
        border-radius: 4px;
        padding: 14px 18px;
        margin-bottom: 8px;
    }
    .candidate-name {
        font-family: 'Lora', serif;
        font-weight: 700;
        font-size: 1.2rem;
        color: #1C1917;
    }
    .candidate-party {
        display: inline-block;
        background-color: #1C1917;
        color: #F4F1EA;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 3px;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .stTabs [data-baseweb="tab-list"] {
        border-bottom: 1px solid #1C1917;
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Lora', serif;
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


def normaliser_chaine(texte):
    nfkd = unicodedata.normalize("NFKD", texte)
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower()


def parser_date(date_str):
    try:
        dt = dateutil.parser.parse(date_str)
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None


def identifier_auteur(texte):
    texte_norm = normaliser_chaine(texte)
    for nom_officiel, variantes, _ in REFERENTIEL_CANDIDATS:
        for v in variantes:
            if re.search(r"\b" + re.escape(v) + r"\b", texte_norm):
                return nom_officiel
    return "Débat présidentiel / Collectif"


def classifier_thematique(texte):
    t = texte.lower()
    if any(
        k in t
        for k in [
            "plastique",
            "consigne",
            "eau",
            "prélèvement",
            "emballage",
            "climat",
            "écologie",
            "réemploi",
        ]
    ):
        return "Environnement"
    elif any(
        k in t for k in ["taxe", "impôt", "fiscal", "tva", "redevance", "budget"]
    ):
        return "Fiscalité"
    elif any(
        k in t
        for k in [
            "egalim",
            "descrozaille",
            "négociation",
            "distribution",
            "marge",
            "prix",
            "grande distribution",
        ]
    ):
        return "Relations commerciales"
    elif any(
        k in t
        for k in [
            "nutri-score",
            "nutrition",
            "sucre",
            "santé",
            "obésité",
            "édulcorant",
            "sel",
        ]
    ):
        return "Health & Nutrition"
    elif any(
        k in t
        for k in [
            "publicité",
            "marketing",
            "enfants",
            "mineurs",
            "réclame",
            "communication",
        ]
    ):
        return "Marketing"
    return "Relations commerciales"


@st.cache_data(ttl=600, show_spinner=False)
def recuperer_propositions():
    articles = []
    titres_vus = set()

    for q in QUERIES:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(q)}&hl=fr&gl=FR&ceid=FR:fr"
        try:
            flux = feedparser.parse(url)
            for entry in flux.entries:
                titre = html.unescape(entry.title)
                if " - " in titre:
                    titre = titre.rsplit(" - ", 1)[0]

                titre_cle = normaliser_chaine(titre)
                if titre_cle in titres_vus:
                    continue

                dt = parser_date(entry.get("published", ""))
                if dt and dt < DATE_MIN:
                    continue

                titres_vus.add(titre_cle)
                source = (
                    entry.source.get("title", "Presse")
                    if hasattr(entry, "source")
                    else "Presse"
                )

                auteur = identifier_auteur(titre)
                theme = classifier_thematique(titre)

                articles.append({
                    "Date": dt.strftime("%d %b %Y %H:%M") if dt else "Récemment",
                    "Auteur": auteur,
                    "Thématique": theme,
                    "Proposition": titre,
                    "Source": source,
                    "Lien": entry.link,
                    "_dt": dt.isoformat() if dt else "2026-01-01T00:00:00",
                })
        except Exception:
            continue

    articles.sort(key=lambda x: x.get("_dt", ""), reverse=True)
    return articles


propositions = recuperer_propositions()
df_total = pd.DataFrame(propositions) if propositions else pd.DataFrame()

if "filtre_candidat" not in st.session_state:
    st.session_state["filtre_candidat"] = "Tous"
if "filtre_theme" not in st.session_state:
    st.session_state["filtre_theme"] = "Tous"


def exporter_excel(df_export):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Propositions Présidentielle"

    headers = [
        "Date",
        "Candidat / Porteur",
        "Thématique",
        "Média / Source",
        "Proposition",
        "Lien",
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

    for row in df_export[[
        "Date",
        "Auteur",
        "Thématique",
        "Source",
        "Proposition",
        "Lien",
    ]].itertuples(index=False):
        ws.append(list(row))

    for r_idx in range(2, len(df_export) + 2):
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if c_idx == 6:
                cell.font = Font(color="8B261E", underline="single")
                cell.hyperlink = cell.value

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 70)

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()


# En-tête éditorial
st.markdown(
    f"""
<div class="top-meta">
    <span>Direction des Affaires Publiques & Réglementaires</span>
    <span>{datetime.now().strftime('%d %B %Y')} — Synchronisation continue</span>
</div>
<h1 class="header-title">Présidentielle 2027<br>— <em>la course, en un coup d'œil</em></h1>
<p class="header-lead">Observatoire systématique des propositions de campagne et prises de position présidentielles impactant le secteur agroalimentaire.</p>
<div class="milestones">
    <span><strong>1er tour :</strong> 18 avril 2027</span>
    <span><strong>2d tour :</strong> 2 mai 2027</span>
</div>
<div class="main-separator"></div>
""",
    unsafe_allow_html=True,
)

tab_candidats, tab_propositions = st.tabs(["Candidats", "Propositions"])

# --- Onglet 1 : Candidats (Épuré sans citations) ---
with tab_candidats:
    st.markdown(
        "<h3 style='font-family:Lora,serif; font-size:1.4rem;"
        " font-weight:700;'>Candidats déclarés & acteurs clés</h3>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Cliquez sur un candidat pour filtrer toutes ses propositions dans"
        " l'onglet dédié."
    )

    c1, c2 = st.columns(2)
    for idx, (nom, _, parti) in enumerate(REFERENTIEL_CANDIDATS):
        cible = c1 if idx % 2 == 0 else c2
        with cible:
            st.markdown(
                f"""
            <div class="candidate-card">
                <div class="candidate-name">{nom}</div>
                <span class="candidate-party">{parti}</span>
            </div>
            """,
                unsafe_allow_html=True,
            )
            if st.button(
                f"🔍 Voir les propositions de {nom}",
                key=f"btn_{nom}",
                use_container_width=True,
            ):
                st.session_state["filtre_candidat"] = nom
                st.rerun()

# --- Onglet 2 : Propositions ---
with tab_propositions:
    st.markdown("**Filtrer par thématique prioritaire :**")
    themes = [
        "Tous",
        "Environnement",
        "Fiscalité",
        "Relations commerciales",
        "Health & Nutrition",
        "Marketing",
    ]
    cols_th = st.columns(len(themes))

    for idx, th in enumerate(themes):
        with cols_th[idx]:
            label = f"🔴 {th}" if st.session_state["filtre_theme"] == th else th
            if st.button(label, key=f"theme_{th}", use_container_width=True):
                st.session_state["filtre_theme"] = th
                st.rerun()

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    df_filtre = df_total.copy()

    # Application du filtre candidat actif
    if st.session_state["filtre_candidat"] != "Tous":
        df_filtre = df_filtre[
            df_filtre["Auteur"] == st.session_state["filtre_candidat"]
        ]
        col_info, col_reset = st.columns([3, 1])
        with col_info:
            st.info(
                "Filtrage actif sur le candidat :"
                f" **{st.session_state['filtre_candidat']}**"
            )
        with col_reset:
            if st.button(
                "✖️ Réinitialiser le filtre candidat", use_container_width=True
            ):
                st.session_state["filtre_candidat"] = "Tous"
                st.rerun()

    # Application du filtre thématique actif
    if st.session_state["filtre_theme"] != "Tous":
        df_filtre = df_filtre[
            df_filtre["Thématique"] == st.session_state["filtre_theme"]
        ]

    # Recherche texte et téléchargement
    col_search, col_dl = st.columns([3, 1.2])
    with col_search:
        mot_cle = st.text_input(
            "Recherche plein texte :",
            placeholder=(
                "Rechercher : plastique, soda, eau, distributeurs,"
                " taxation..."
            ),
            label_visibility="collapsed",
        )
        if mot_cle:
            df_filtre = df_filtre[
                df_filtre["Proposition"].str.contains(
                    mot_cle, case=False, na=False
                )
                | df_filtre["Auteur"].str.contains(
                    mot_cle, case=False, na=False
                )
                | df_filtre["Source"].str.contains(
                    mot_cle, case=False, na=False
                )
            ]

    with col_dl:
        if not df_filtre.empty:
            fichier_excel = exporter_excel(df_filtre)
            st.download_button(
                label="📥 Télécharger l'Excel",
                data=fichier_excel,
                file_name=(
                    "propositions_presidentielle_"
                    f"{datetime.now().strftime('%Y%m%d')}.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.caption(
        f"{len(df_filtre)} proposition(s) répertoriée(s) — classées de la plus"
        " récente à la plus ancienne"
    )

    if not df_filtre.empty:
        st.dataframe(
            df_filtre[[
                "Date",
                "Auteur",
                "Thématique",
                "Source",
                "Proposition",
                "Lien",
            ]],
            column_config={
                "Lien": st.column_config.LinkColumn(
                    "Lien presse", display_text="Consulter ↗"
                ),
                "Date": st.column_config.TextColumn("Date", width="small"),
                "Auteur": st.column_config.TextColumn(
                    "Qui l'a dit / Porteur", width="medium"
                ),
                "Thématique": st.column_config.TextColumn(
                    "Pôle thématique", width="small"
                ),
                "Source": st.column_config.TextColumn("Média", width="small"),
                "Proposition": st.column_config.TextColumn(
                    "Détail de la proposition / Annonce", width="large"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.info("Aucune proposition ne correspond à cette combinaison de filtres.")
