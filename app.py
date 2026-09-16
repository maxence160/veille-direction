import html
import json
import os
import urllib.parse
from datetime import datetime
import feedparser
import pandas as pd
import streamlit as st

# Configuration de la page
st.set_page_config(
    page_title="Veille Présidentielle", page_icon="🗳️", layout="wide"
)

DB_FILE = "veille_data.json"
SEEN_FILE = "deja_vus.json"

QUERIES = [
    "présidentielle proposition loi",
    "présidentielle programme candidat",
    "présidentielle mesure campagne",
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
                    else "Média"
                )
                date_pub = entry.get(
                    "published", datetime.now().strftime("%Y-%m-%d %H:%M")
                )

                donnees_existantes.insert(
                    0,
                    {
                        "Date": date_pub,
                        "Source": source,
                        "Titre": html.unescape(entry.title),
                        "Lien": lien,
                    },
                )
                nouveaux_ajouts += 1

    sauvegarder_vus(deja_vus)
    sauvegarder_donnees(donnees_existantes)
    return nouveaux_ajouts


# --- Interface Web ---
st.title("🗳️ Veille Stratégique — Présidentielle")
st.caption("Tableau de bord de suivi des annonces et propositions politiques")

# Barre d'actions en haut
col1, col2 = st.columns([1, 4])

with col1:
    if st.button("🔄 Actualiser les données", type="primary"):
        with st.spinner("Recherche des dernières annonces..."):
            nb = lancer_actualisation()
            if nb > 0:
                st.success(f"{nb} nouvelle(s) annonce(s) ajoutée(s) !")
            else:
                st.info("Aucune nouvelle annonce détectée.")

# Chargement et affichage des données
articles = charger_donnees()

if articles:
    df = pd.DataFrame(articles)

    # Zone de recherche pour la PDG
    recherche = st.text_input(
        "🔍 Filtrer par mot-clé (ex: retraites, fiscalité, santé...)"
    )
    if recherche:
        df = df[
            df["Titre"].str.contains(recherche, case=False, na=False)
            | df["Source"].str.contains(recherche, case=False, na=False)
        ]

    st.write(f"**{len(df)} annonces répertoriées**")

    # Tableau interactif avec liens cliquables
    st.dataframe(
        df,
        column_config={
            "Lien": st.column_config.LinkColumn(
                "Lien direct", display_text="Ouvrir l'article"
            ),
            "Date": st.column_config.TextColumn("Date de publication"),
            "Source": st.column_config.TextColumn("Média / Source", width="medium"),
            "Titre": st.column_config.TextColumn("Proposition / Annonce", width="large"),
        },
        hide_index=True,
        use_container_width=True,
    )
else:
    st.warning("Aucune donnée enregistrée. Cliquez sur 'Actualiser les données' pour lancer la première collecte.")