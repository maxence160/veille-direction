import html
import io
import json
import os
import urllib.parse
from datetime import datetime
import dateutil.parser
import feedparser
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as st

# Configuration de la page
st.set_page_config(
    page_title="Présidentielle 2027 — La course, en un coup d'œil",
    page_icon="⚖️",
    layout="wide",
)

# Date pivot : 1er janvier 2026
DATE_MIN = datetime(2026, 1, 1, 0, 0, 0)

# CSS éditorial reprenant exactement la DA
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

DB_FILE = "veille_data.json"
CANDIDATS_FILE = "candidats_data.json"

QUERIES = [
    'présidentielle "proposition de loi"',
    'présidentielle "projet de loi"',
    'présidentielle "agroalimentaire" OR "alimentation"',
    'présidentielle "taxe soda" OR "sucre"',
    'présidentielle agriculture fiscalité',
]


# --- Récupération en direct des candidats déclarés ---
def charger_candidats_en_ligne():
    try:
        url = "https://fr.wikipedia.org/w/api.php?action=parse&page=Candidatures_%C3%A0_l%27%C3%A9lection_pr%C3%A9sidentielle_fran%C3%A7aise_de_2027&prop=wikitext&format=json"
        res = requests.get(url, headers={"User-Agent": "VeilleDirection/1.0"}).json()
        wikitext = res["parse"]["wikitext"]["*"]

        # Extraction des sections de candidatures
        lignes = wikitext.split("\n")
        candidats = []
        capture = False

        for l in lignes:
            if "== Candidats déclarés ==" in l:
                capture = True
                continue
            if capture and l.startswith("==") and "Candidats" not in l:
                break
            if capture and l.startswith("=== "):
                nom = l.replace("===", "").strip().replace("[[", "").replace("]]", "")
                if "|" in nom:
                    nom = nom.split("|")[-1]
                candidats.append({"nom": nom, "parti": "Candidature déclarée"})

        if candidats:
            with open(CANDIDATS_FILE, "w", encoding="utf-8") as f:
                json.dump(candidats, f, ensure_ascii=False)
            return candidats
    except Exception:
        pass

    # Données par défaut si l'API est temporairement inaccessible
    return [
        {"nom": "Édouard Philippe", "parti": "Horizons — Démarche confirmée"},
        {"nom": "Marine Le Pen", "parti": "Rassemblement National — Déclarée"},
        {"nom": "Jean-Luc Mélenchon", "parti": "La France Insoumise — Pôle de gauche"},
        {"nom": "Laurent Wauquiez", "parti": "La Droite Républicaine — Assemblée nationale"},
        {"nom": "Fabien Roussel", "parti": "Parti Communiste Français — Candidature autonome"},
        {"nom": "François Ruffin", "parti": "Debout ! — Mouvement populaire"},
    ]


# --- Récupération des textes post-2026 ---
def charger_donnees():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def sauvegarder_donnees(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def parser_date(date_str):
    try:
        dt = dateutil.parser.parse(date_str)
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None


def actualiser_veille():
    existants = charger_donnees()
    liens_vus = {a["Lien"] for a in existants}
    nouveaux = []

    for q in QUERIES:
        query_enc = urllib.parse.quote(q)
        url = f"https://news.google.com/rss/search?q={query_enc}&hl=fr&gl=FR&ceid=FR:fr"
        flux = feedparser.parse(url)

        for entry in flux.entries:
            lien = entry.link
            if lien in liens_vus:
                continue

            date_brute = entry.get("published", "")
            date_obj = parser_date(date_brute)

            # Filtre strict : après le 1er janvier 2026
            if date_obj and date_obj < DATE_MIN:
                continue

            liens_vus.add(lien)
            source = entry.source.get("title", "Presse") if hasattr(entry, "source") else "Presse"
            titre = html.unescape(entry.title)

            # Nettoyage de la source dans le titre
            if " - " in titre:
                titre = titre.rsplit(" - ", 1)[0]

            date_lisible = date_obj.strftime("%d %b %Y %H:%M") if date_obj else "2026"

            nouveaux.append({
                "Date": date_lisible,
                "Source": source,
                "Titre": titre,
                "Lien": lien,
                "_dt": date_obj.isoformat() if date_obj else "",
            })

    total = nouveaux + existants
    # Tri antichronologique
    total.sort(key=lambda x: x.get("_dt", ""), reverse=True)
    sauvegarder_donnees(total)
    return len(nouveaux)


def exporter_excel(df_export):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Propositions Législatives"

    headers = ["Date", "Média / Source", "Proposition / Annonce de loi", "Lien source"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1C1917", end_color="1C1917", fill_type="solid")
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

    for row in df_export[["Date", "Source", "Titre", "Lien"]].itertuples(index=False):
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


# --- En-tête ---
st.markdown(
    f"""
<div class="top-meta">
    <span>Répertoire institutionnel de veille stratégique</span>
    <span>{datetime.now().strftime('%d %B %Y')}</span>
</div>
<h1 class="header-title">Présidentielle 2027<br>— <em>la course, en un coup d'œil</em></h1>
<p class="header-lead">Point de repère sur les candidatures et sélection des propositions de loi relayées dans l'actualité politique et agroalimentaire.</p>
<div class="milestones">
    <span><strong>1er tour :</strong> 18 avril 2027</span>
    <span><strong>2d tour :</strong> 2 mai 2027</span>
    <span><strong>Filtre temporel :</strong> Décrets et annonces post-1er janvier 2026</span>
</div>
<div class="main-separator"></div>
""",
    unsafe_allow_html=True,
)

tab_candidats, tab_lois = st.tabs(["Candidats", "Propositions de loi"])

# --- Onglet Candidats ---
with tab_candidats:
    col_t, col_b = st.columns([3, 1])
    with col_t:
        st.markdown("<h3 style='font-family:Lora,serif; font-size:1.4rem; font-weight:700;'>Candidatures déclarées recensées</h3>", unsafe_allow_html=True)
    with col_b:
        if st.button("🔄 Actualiser les candidats", use_container_width=True):
            charger_candidats_en_ligne()
            st.rerun()

    candidats_liste = charger_candidats_en_ligne()
    c1, c2 = st.columns(2)
    for idx, c in enumerate(candidats_liste):
        col_dest = c1 if idx % 2 == 0 else c2
        with col_dest:
            st.markdown(
                f"""
            <div class="candidate-box">
                <div class="candidate-name">{c['nom']}</div>
                <div class="candidate-role">{c['parti']}</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

# --- Onglet Propositions de loi ---
with tab_lois:
    col_act, col_search, col_dl = st.columns([1.2, 2.5, 1.3])

    with col_act:
        if st.button("🔄 Actualiser les lois", use_container_width=True):
            with st.spinner("Analyse du web et des parutions depuis le 01/01/2026..."):
                nb = actualiser_veille()
                if nb > 0:
                    st.success(f"{nb} actualité(s) récente(s) intégrée(s).")
                else:
                    st.info("Aucune nouvelle publication par rapport à la base.")

    data_lois = charger_donnees()
    df_lois = pd.DataFrame(data_lois) if data_lois else pd.DataFrame()

    with col_search:
        mot_recherche = st.text_input(
            "Filtrer par mot-clé :",
            placeholder="Filtrer : taxe, alimentation, PME, fiscalité...",
            label_visibility="collapsed",
        )

    if not df_lois.empty:
        if mot_recherche:
            df_lois = df_lois[
                df_lois["Titre"].str.contains(mot_recherche, case=False, na=False)
                | df_lois["Source"].str.contains(mot_recherche, case=False, na=False)
            ]

        with col_dl:
            fichier_excel = exporter_excel(df_lois)
            st.download_button(
                label="📥 Télécharger l'Excel",
                data=fichier_excel,
                file_name=f"veille_lois_2026_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.caption(f"{len(df_lois)} annonce(s) et proposition(s) parue(s) après le 01/01/2026")

        st.dataframe(
            df_lois[["Date", "Source", "Titre", "Lien"]],
            column_config={
                "Lien": st.column_config.LinkColumn("Source", display_text="Consulter ↗"),
                "Date": st.column_config.TextColumn("Parution", width="small"),
                "Source": st.column_config.TextColumn("Média", width="small"),
                "Titre": st.column_config.TextColumn("Proposition / Mesure", width="large"),
            },
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.info("Aucun article répertorié. Cliquez sur 'Actualiser les lois' pour lancer la première collecte automatique.")
