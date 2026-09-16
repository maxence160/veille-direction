import html
import io
import json
import os
import re
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
DB_LOIS = "veille_data.json"
DB_CANDIDATS = "candidats_presse.json"

QUERIES_CANDIDATS = [
    'présidentielle "se déclare candidat"',
    'présidentielle "annonce sa candidature"',
    'présidentielle "candidat officiel"',
    'présidentielle "candidature déclarée"',
    'présidentielle "investi par"',
    'présidentielle primaire parti',
]

QUERIES_LOIS = [
    'présidentielle "proposition de loi"',
    'présidentielle "projet de loi"',
    'présidentielle "mesure de campagne"',
    'présidentielle "programme économique"',
    'présidentielle fiscalité entreprises',
    'présidentielle "agroalimentaire"',
    'présidentielle "taxe soda" OR "taxe sucre"',
    'présidentielle agriculture alimentation',
    'présidentielle "industrie agroalimentaire"',
    'présidentielle emballage plastique consigne',
    'présidentielle régulation distribution marges',
    'présidentielle EGAlim négociations commerciales',
]

PARTIS_CONNUS = [
    ("Horizons", "Horizons"),
    ("Renaissance", "Renaissance"),
    ("Rassemblement National", "Rassemblement National"),
    ("RN", "Rassemblement National"),
    ("La France Insoumise", "La France Insoumise"),
    ("LFI", "La France Insoumise"),
    ("Parti Socialiste", "Parti Socialiste"),
    ("PS", "Parti Socialiste"),
    ("Les Républicains", "Les Républicains"),
    ("LR", "Les Républicains"),
    ("Droite Républicaine", "La Droite Républicaine"),
    ("Écologistes", "Les Écologistes (EELV)"),
    ("EELV", "Les Écologistes (EELV)"),
    ("Parti Communiste", "PCF"),
    ("PCF", "PCF"),
    ("Reconquête", "Reconquête"),
    ("Debout", "Debout !"),
    ("MoDem", "MoDem"),
    ("UDI", "UDI"),
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
        font-size: 1.1rem;
        color: #1C1917;
        margin-bottom: 4px;
    }
    .candidate-role {
        font-size: 0.85rem;
        color: #57534E;
    }
    .candidate-source {
        display: inline-block;
        margin-top: 6px;
        font-size: 0.8rem;
        color: #8B261E;
        text-decoration: none;
        font-weight: 600;
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


def parser_date(date_str):
    try:
        dt = dateutil.parser.parse(date_str)
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None


def detecter_parti(texte):
    for declinaison, label in PARTIS_CONNUS:
        if re.search(r"\b" + re.escape(declinaison) + r"\b", texte, re.IGNORECASE):
            return label
    return "Parti / Mouvement en attente de précision"


def charger_fichier(chemin):
    if os.path.exists(chemin):
        with open(chemin, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def sauvegarder_fichier(chemin, data):
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# Mise à jour automatique en arrière-plan (au chargement)
@st.cache_data(ttl=1800, show_spinner=False)
def sync_presse_automatique():
    existants_lois = charger_fichier(DB_LOIS)
    liens_vus_lois = {a["Lien"] for a in existants_lois}
    nouvelles_lois = []

    for q in QUERIES_LOIS:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(q)}&hl=fr&gl=FR&ceid=FR:fr"
        flux = feedparser.parse(url)

        for entry in flux.entries:
            lien = entry.link
            if lien in liens_vus_lois:
                continue

            dt = parser_date(entry.get("published", ""))
            # Rejet de tout ce qui précède le 01/01/2026
            if dt and dt < DATE_MIN:
                continue

            liens_vus_lois.add(lien)
            source = entry.source.get("title", "Presse") if hasattr(entry, "source") else "Presse"
            titre = html.unescape(entry.title)
            if " - " in titre:
                titre = titre.rsplit(" - ", 1)[0]

            nouvelles_lois.append({
                "Date": dt.strftime("%d %b %Y %H:%M") if dt else "2026",
                "Source": source,
                "Titre": titre,
                "Lien": lien,
                "_dt": dt.isoformat() if dt else "2026-01-01T00:00:00",
            })

    total_lois = nouvelles_lois + existants_lois
    # Classement strict : plus récent en haut
    total_lois.sort(key=lambda x: x.get("_dt", ""), reverse=True)
    sauvegarder_fichier(DB_LOIS, total_lois)

    existants_cand = charger_fichier(DB_CANDIDATS)
    liens_vus_cand = {c["Lien"] for c in existants_cand}
    nouveaux_cand = []

    for q in QUERIES_CANDIDATS:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(q)}&hl=fr&gl=FR&ceid=FR:fr"
        flux = feedparser.parse(url)

        for entry in flux.entries:
            lien = entry.link
            if lien in liens_vus_cand:
                continue

            dt = parser_date(entry.get("published", ""))
            if dt and dt < DATE_MIN:
                continue

            liens_vus_cand.add(lien)
            source = entry.source.get("title", "Presse") if hasattr(entry, "source") else "Presse"
            titre = html.unescape(entry.title)
            if " - " in titre:
                titre = titre.rsplit(" - ", 1)[0]

            parti = detecter_parti(titre)

            nouveaux_cand.append({
                "Annonce": titre,
                "Parti": parti,
                "Source": source,
                "Lien": lien,
                "Date": dt.strftime("%d %b %Y") if dt else "2026",
                "_dt": dt.isoformat() if dt else "2026-01-01T00:00:00",
            })

    total_cand = nouveaux_cand + existants_cand
    total_cand.sort(key=lambda x: x.get("_dt", ""), reverse=True)
    sauvegarder_fichier(DB_CANDIDATS, total_cand)

    return len(nouvelles_lois), len(nouveaux_cand)


# Exécution automatique immédiate à l'ouverture
sync_presse_automatique()

data_lois = charger_fichier(DB_LOIS)
data_cand = charger_fichier(DB_CANDIDATS)


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


# En-tête
st.markdown(
    f"""
<div class="top-meta">
    <span>Répertoire institutionnel de veille stratégique</span>
    <span>{datetime.now().strftime('%d %B %Y')} — Synchronisé en direct</span>
</div>
<h1 class="header-title">Présidentielle 2027<br>— <em>la course, en un coup d'œil</em></h1>
<p class="header-lead">Point de repère sur les candidatures et sélection élargie des propositions de loi relayées dans l'actualité politique et agroalimentaire.</p>
<div class="milestones">
    <span><strong>1er tour :</strong> 18 avril 2027</span>
    <span><strong>2d tour :</strong> 2 mai 2027</span>
    <span><strong>Filtre strict :</strong> Parutions presse post-1er janvier 2026</span>
</div>
<div class="main-separator"></div>
""",
    unsafe_allow_html=True,
)

tab_candidats, tab_lois = st.tabs(["Candidats", "Propositions de loi"])

# Onglet 1 : Candidats
with tab_candidats:
    st.markdown("<h3 style='font-family:Lora,serif; font-size:1.4rem; font-weight:700;'>Candidatures & investitures relevées dans la presse</h3>", unsafe_allow_html=True)
    st.caption("Articles de presse récents mentionnant des déclarations officielles de candidatures et mouvements de partis (post-2026).")

    if data_cand:
        c1, c2 = st.columns(2)
        for idx, c in enumerate(data_cand):
            col_dest = c1 if idx % 2 == 0 else c2
            with col_dest:
                st.markdown(
                    f"""
                <div class="candidate-box">
                    <div class="candidate-name">{c['Annonce']}</div>
                    <div class="candidate-role"><strong>Parti détecté :</strong> {c['Parti']} — <em>{c['Source']} ({c['Date']})</em></div>
                    <a class="candidate-source" href="{c['Lien']}" target="_blank">Lire la déclaration ↗</a>
                </div>
                """,
                    unsafe_allow_html=True,
                )
    else:
        st.info("Recherche des déclarations récentes en cours...")

# Onglet 2 : Propositions de loi
with tab_lois:
    df_lois = pd.DataFrame(data_lois) if data_lois else pd.DataFrame()

    col_search, col_dl = st.columns([3, 1.2])
    with col_search:
        mot_cle = st.text_input(
            "Filtrer par mot-clé :",
            placeholder="Filtrer : soda, agriculture, emballage, distribution, marge...",
            label_visibility="collapsed",
        )

    if not df_lois.empty:
        if mot_cle:
            df_lois = df_lois[
                df_lois["Titre"].str.contains(mot_cle, case=False, na=False)
                | df_lois["Source"].str.contains(mot_cle, case=False, na=False)
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

        st.caption(f"{len(df_lois)} annonce(s) et proposition(s) répertoriée(s) — classées de la plus récente à la plus ancienne")

        st.dataframe(
            df_lois[["Date", "Source", "Titre", "Lien"]],
            column_config={
                "Lien": st.column_config.LinkColumn("Source", display_text="Consulter ↗"),
                "Date": st.column_config.TextColumn("Horodatage", width="small"),
                "Source": st.column_config.TextColumn("Média", width="small"),
                "Titre": st.column_config.TextColumn("Proposition / Mesure", width="large"),
            },
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.info("Aucune parution post-2026 répertoriée pour le moment.")
