
import os
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import altair as alt

# Google Drive
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

st.set_page_config(page_title="Dashboard Multi-Projets (Drive)", page_icon="📊", layout="wide")

st.title("📊 Dashboard Multi-Projets — Connecté à Google Drive")
st.caption("Importer / exporter votre fichier Excel directement depuis/vers Google Drive (via PyDrive2)")

# -------------------- Auth Google Drive --------------------
st.sidebar.header("🔐 Connexion Google Drive")

@st.cache_resource(show_spinner=False)
def init_drive():
    gauth = GoogleAuth()
    # Requiert un credentials.json dans le même dossier (OAuth Client ID)
    # Le premier lancement ouvrira une fenêtre de login Google dans le navigateur
    gauth.LocalWebserverAuth()
    return GoogleDrive(gauth)

drive = None
if st.sidebar.button("Se connecter à Google Drive"):
    try:
        drive = init_drive()
        st.sidebar.success("Connecté à Google Drive ✅")
    except Exception as e:
        st.sidebar.error(f"Erreur de connexion: {e}")

# -------------------- Import depuis Drive --------------------
st.sidebar.header("📥 Import depuis Drive")
default_filename = st.sidebar.text_input("Nom du fichier Excel (ex: Dashboard_MultiProjets_v56.xlsx)", "Dashboard_MultiProjets_v56.xlsx")
file_id_input = st.sidebar.text_input("OU ID du fichier Drive (optionnel)", "")

def get_file_from_drive(drive, filename=None, file_id=None):
    if file_id:
        f = drive.CreateFile({'id': file_id})
        local_name = f['title']
        f.GetContentFile(local_name)
        return local_name
    # sinon recherche par nom
    q = f"title='{filename}' and trashed=false"
    files = drive.ListFile({'q': q}).GetList()
    if not files:
        return None
    f = files[0]
    local_name = filename
    f.GetContentFile(local_name)
    return local_name

st.sidebar.header("📁 Fichier local / Upload")
uploaded = st.sidebar.file_uploader("…ou importer un Excel local", type=["xlsx"])
use_default_local = st.sidebar.checkbox("Utiliser le fichier local par défaut s'il existe", value=False)

# -------------------- Chargement dataframe --------------------
def load_dataframe(file):
    # Les 2 premières lignes sont la légende, les en-têtes sont sur la ligne 3
    df = pd.read_excel(file, sheet_name="Planning_Projets", skiprows=2)
    expected = ["Projet", "Responsable", "Date début", "Date fin", "État", "Progression (%)"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        st.warning(f"Colonnes manquantes: {missing}. Le tableau doit contenir {expected}.")
    # Types
    for c in ["Date début", "Date fin"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    if "Progression (%)" in df.columns:
        df["Progression (%)"] = (
            df["Progression (%)"].astype(str).str.replace("%","",regex=False).str.strip().replace({"":np.nan}).astype(float)
        )
    if "Alerte 🚨" not in df.columns:
        df["Alerte 🚨"] = ""
    if "Évolution 📈" not in df.columns:
        df["Évolution 📈"] = ""
    return df

df = None
source_info = st.empty()

col_imp1, col_imp2, col_imp3 = st.columns([1,1,1])
with col_imp1:
    if st.button("📥 Importer depuis Drive"):
        if drive is None:
            st.error("Connectez-vous à Google Drive d'abord (bouton dans la barre latérale).")
        else:
            local_path = None
            try:
                local_path = get_file_from_drive(drive, filename=default_filename, file_id=file_id_input or None)
            except Exception as e:
                st.error(f"Erreur import Drive: {e}")
            if local_path and os.path.exists(local_path):
                df = load_dataframe(local_path)
                source_info.info(f"Chargé depuis Drive: {os.path.basename(local_path)} ✅")
            else:
                st.warning("Fichier introuvable dans Drive. Vérifiez le nom ou l'ID.")

with col_imp2:
    if uploaded is not None:
        df = load_dataframe(uploaded)
        source_info.info("Chargé depuis un upload local ✅")

with col_imp3:
    if use_default_local and os.path.exists("Dashboard_MultiProjets_v56.xlsx"):
        df = load_dataframe("Dashboard_MultiProjets_v56.xlsx")
        source_info.info("Chargé depuis le fichier local par défaut ✅")

if df is None:
    st.info("👉 Importez un fichier (Drive ou local) pour continuer.")
    st.stop()

# -------------------- Calculs et KPI --------------------
def calc_retard_row(row, today=None):
    today = today or pd.Timestamp.today().normalize()
    start, end, prog = row.get("Date début"), row.get("Date fin"), row.get("Progression (%)", 0) or 0
    if pd.isna(start) or pd.isna(end):
        return np.nan, np.nan
    if today <= start:
        attendu = 0
    elif today >= end:
        attendu = 100
    else:
        total = (end - start).days
        done = (today - start).days
        attendu = int((done / total) * 100) if total > 0 else 0
    retard = max(0, attendu - float(prog or 0))
    return attendu, retard

if {"Date début","Date fin","Progression (%)"}.issubset(df.columns):
    temp = df.copy()
    temp[["Progression attendue (%)","Retard (%)"]] = temp.apply(lambda r: pd.Series(calc_retard_row(r)), axis=1)
    df["Progression attendue (%)"] = temp["Progression attendue (%)"]
    df["Retard (%)"] = temp["Retard (%)"]
else:
    df["Progression attendue (%)"] = np.nan
    df["Retard (%)"] = np.nan

total = len(df)
n_prevus = (df["État"] == "Prévu").sum() if "État" in df.columns else 0
n_encours = (df["État"] == "En cours").sum() if "État" in df.columns else 0
n_termines = (df["État"] == "Terminé").sum() if "État" in df.columns else 0
n_bloques = (df["État"] == "Bloqué").sum() if "État" in df.columns else 0
avg_prog = float(df["Progression (%)"].mean()) if "Progression (%)" in df.columns else 0.0
n_retards = (df["Retard (%)"] > 0).sum()

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total projets", total)
c2.metric("🟨 Prévu", int(n_prevus))
c3.metric("🟩 En cours", int(n_encours))
c4.metric("🟦 Terminé", int(n_termines))
c5.metric("🟥 Bloqué", int(n_bloques))
c6.metric("📊 Prog. moyenne", f"{avg_prog:.0f}%")

st.divider()

# -------------------- Filtres --------------------
st.subheader("🔎 Filtres")
fc1, fc2, fc3 = st.columns(3)
responsables = sorted([x for x in df["Responsable"].dropna().unique()]) if "Responsable" in df.columns else []
etats = sorted([x for x in df["État"].dropna().unique()]) if "État" in df.columns else []
alertes = ["⚠️ Retard","✅ OK"]
sel_resp = fc1.multiselect("Responsable(s)", responsables, default=responsables[:3] if len(responsables)>3 else responsables)
sel_etat = fc2.multiselect("État(s)", etats, default=etats)
sel_alert = fc3.selectbox("Alerte", ["(Tous)"] + alertes, index=0)

df_filtered = df.copy()
if sel_resp:
    df_filtered = df_filtered[df_filtered["Responsable"].isin(sel_resp)]
if sel_etat:
    df_filtered = df_filtered[df_filtered["État"].isin(sel_etat)]
if sel_alert != "(Tous)":
    if sel_alert == "⚠️ Retard":
        df_filtered = df_filtered[df_filtered["Retard (%)"] > 0]
    else:
        df_filtered = df_filtered[(df_filtered["Retard (%)"].fillna(0) == 0)]

st.subheader("📋 Projets (filtrés)")
st.dataframe(df_filtered, use_container_width=True, hide_index=True)

# -------------------- Visualisations --------------------
st.subheader("📈 Visualisations")

left, right = st.columns(2)

with left:
    st.markdown("**Top 5 retards (%)**")
    top5 = df.sort_values("Retard (%)", ascending=False).head(5)[["Projet","Retard (%)"]].dropna()
    if not top5.empty:
        chart = alt.Chart(top5).mark_bar().encode(
            x=alt.X("Retard (%)", title="Retard (%)"),
            y=alt.Y("Projet", sort="-x", title="Projet"),
            tooltip=["Projet","Retard (%)"]
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("Aucun retard détecté.")

with right:
    st.markdown("**Répartition par Responsable × État**")
    if {"Responsable","État"}.issubset(df.columns) and len(df):
        pivot = (df.groupby(["Responsable","État"])["Projet"].count().reset_index(name="Nombre"))
        stacked = alt.Chart(pivot).mark_bar().encode(
            x=alt.X("sum(Nombre)", stack="zero", title="Nombre de projets"),
            y=alt.Y("Responsable", sort="-x"),
            color=alt.Color("État", legend=alt.Legend(title="État")),
            tooltip=["Responsable","État","Nombre"]
        ).properties(height=300)
        st.altair_chart(stacked, use_container_width=True)
    else:
        st.info("Colonnes Responsable / État absentes.")

st.divider()

# -------------------- Edition --------------------
st.subheader("✏️ Édition rapide")
with st.expander("Ajouter un projet"):
    fcol1, fcol2 = st.columns(2)
    p_projet = fcol1.text_input("Projet")
    p_resp = fcol2.text_input("Responsable")
    gcol1, gcol2, gcol3 = st.columns(3)
    p_debut = gcol1.date_input("Date début", value=datetime.today())
    p_fin = gcol2.date_input("Date fin", value=datetime.today())
    p_etat = gcol3.selectbox("État", ["Prévu","En cours","Terminé","Bloqué"])
    p_prog = st.slider("Progression (%)", 0, 100, 0)
    if st.button("➕ Ajouter à la table"):
        new_row = {
            "Projet": p_projet,
            "Responsable": p_resp,
            "Date début": pd.to_datetime(p_debut),
            "Date fin": pd.to_datetime(p_fin),
            "État": p_etat,
            "Progression (%)": p_prog,
        }
        # calc
        start = pd.to_datetime(new_row["Date début"])
        end = pd.to_datetime(new_row["Date fin"])
        if pd.isna(start) or pd.isna(end):
            attendu = np.nan; retard = np.nan
        else:
            today = pd.Timestamp.today().normalize()
            if today <= start: attendu = 0
            elif today >= end: attendu = 100
            else:
                total = (end - start).days
                done = (today - start).days
                attendu = int((done/total)*100) if total>0 else 0
            retard = max(0, attendu - float(new_row["Progression (%)"] or 0))
        new_row["Progression attendue (%)"] = attendu
        new_row["Retard (%)"] = retard
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        st.success("Projet ajouté (non encore sauvegardé).")

st.subheader("🛠️ Data editor (modifications en place)")
edited = st.data_editor(
    df,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    column_config={
        "Progression (%)": st.column_config.NumberColumn(format="%.0f"),
        "Progression attendue (%)": st.column_config.NumberColumn(format="%.0f"),
        "Retard (%)": st.column_config.NumberColumn(format="%.0f"),
    }
)
df = edited.copy()

# -------------------- Exports --------------------
st.subheader("💾 Exports")

def to_excel_bytes(df_out):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Plage légende vide lignes 1-2
        # puis en-têtes ligne 3
        startrow = 2  # 0-based => ligne 3
        df_out.to_excel(writer, sheet_name="Planning_Projets", startrow=startrow, index=False)
        dash = writer.book.create_sheet("Dashboard")
        dash["A1"] = "Tableau généré par Streamlit (Drive)"
    output.seek(0)
    return output

colx, coly, colz = st.columns(3)

with colx:
    if st.button("⬇️ Exporter en Excel (local)"):
        xls = to_excel_bytes(df)
        st.download_button(
            "Télécharger Excel mis à jour",
            data=xls,
            file_name=f"Dashboard_MultiProjets_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with coly:
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Exporter en CSV (local)", data=csv, file_name="planning_export.csv", mime="text/csv")

with colz:
    folder_id = st.text_input("ID dossier Drive pour l'export (optionnel)", "")
    if st.button("📤 Exporter vers Google Drive"):
        if drive is None:
            st.error("Connectez-vous à Google Drive d'abord.")
        else:
            # Sauvegarder un fichier temporaire local puis uploader
            local_name = f"Dashboard_MultiProjets_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            with open(local_name, "wb") as f:
                f.write(to_excel_bytes(df).read())
            try:
                meta = {'title': os.path.basename(local_name)}
                if folder_id:
                    meta['parents'] = [{'id': folder_id}]
                fdr = drive.CreateFile(meta)
                fdr.SetContentFile(local_name)
                fdr.Upload()
                st.success(f"✅ Exporté dans Drive : {fdr['title']}")
            except Exception as e:
                st.error(f"Erreur d'upload Drive: {e}")

st.success("Interface Drive prête ✔️ — Pensez à placer credentials.json à côté du script.")
