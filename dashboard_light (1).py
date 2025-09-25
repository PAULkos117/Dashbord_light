
import os
import time
from datetime import datetime
from zipfile import ZipFile

import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook

# Google Drive
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# OpenAI
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Dashboard Light", page_icon="⚡", layout="centered")
st.title("⚡ Dashboard GPT Ultra-Light")
st.caption("4 clics suffisent pour lancer un projet GPT avec suivi Excel et Google Drive.")

# =========================
# Connexion Drive
# =========================
drive = None
if st.button("1️⃣ Se connecter à Google Drive"):
    try:
        gauth = GoogleAuth()
        gauth.LocalWebserverAuth()
        drive = GoogleDrive(gauth)
        st.success("Connecté à Google Drive ✅")
    except Exception as e:
        st.error(f"Erreur de connexion Drive : {e}")

# =========================
# Import Excel
# =========================
uploaded_excel = None
if st.button("2️⃣ Charger Excel de suivi"):
    try:
        if drive:
            # Cherche un fichier par défaut dans Drive
            q = "title='Exemple_Suivi_Projet_GPT.xlsx' and trashed=false"
            files = drive.ListFile({'q': q}).GetList()
            if files:
                f = files[0]
                f.GetContentFile("Exemple_Suivi_Projet_GPT.xlsx")
                uploaded_excel = "Exemple_Suivi_Projet_GPT.xlsx"
                st.success("Excel chargé depuis Drive ✅")
            else:
                st.warning("Pas trouvé dans Drive, upload manuel recommandé.")
        else:
            st.info("Upload manuel requis")
            file = st.file_uploader("Uploader un fichier Excel", type=["xlsx"])
            if file:
                uploaded_excel = "Suivi.xlsx"
                with open(uploaded_excel, "wb") as f:
                    f.write(file.getbuffer())
                st.success("Excel chargé ✅")
    except Exception as e:
        st.error(f"Erreur lors du chargement : {e}")

# =========================
# Paramètres Projet
# =========================
title = st.text_input("Nom du projet GPT", "Grand Traité Universel")
total_pages = st.number_input("Nombre de pages", 1, 10000, 1000)
words_per_page = st.number_input("Mots par page", 50, 2000, 700)
pages_per_lot = st.number_input("Pages par lot", 1, 100, 5)

style = f"""
• Chaque page ≈ {words_per_page} mots
• Notes de bas de page à la fin de chaque page
• Sous-titres quand nécessaire
• Pas de résumé intermédiaire
• Langue : français clair et structuré
"""

# =========================
# Lancer GPT
# =========================
api_key = os.getenv("OPENAI_API_KEY", "")
if st.button("3️⃣ Lancer GPT 🚀"):
    if not api_key:
        st.error("Ajoute ta clé OPENAI_API_KEY (variable d'environnement).")
    elif OpenAI is None:
        st.error("Librairie openai manquante. Installe-la avec `pip install openai`.")
    else:
        client = OpenAI(api_key=api_key)
        lots = [(i, min(i+pages_per_lot-1, total_pages)) for i in range(1, total_pages+1, pages_per_lot)]
        os.makedirs("outputs_light", exist_ok=True)
        st.info("Génération en cours...")

        system_prompt = f"Tu es un écrivain. Respecte strictement {words_per_page} mots par page et le format demandé."

        results = []
        for idx, (s,e) in enumerate(lots, start=1):
            user_prompt = f"Rédige les pages {s} à {e} de '{title}'. Contraintes:\n{style}"
            try:
                resp = client.chat.completions.create(
                    model="gpt-4o-mini",
                    temperature=0.3,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                )
                content = resp.choices[0].message.content
                fname = f"outputs_light/{title.replace(' ','_')}_p{s}_p{e}.txt"
                with open(fname, "w", encoding="utf-8") as f:
                    f.write(content)
                results.append(fname)
            except Exception as err:
                results.append(f"Erreur: {err}")

            st.progress(idx/len(lots))

        # ZIP final
        zip_name = f"{title.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
        zip_path = os.path.join("outputs_light", zip_name)
        with ZipFile(zip_path, "w") as z:
            for r in results:
                if os.path.exists(r):
                    z.write(r, arcname=os.path.basename(r))

        with open(zip_path, "rb") as f:
            st.download_button("⬇️ Télécharger le ZIP final", f, file_name=zip_name, mime="application/zip")

        st.success("Projet généré et sauvegardé ✅")
