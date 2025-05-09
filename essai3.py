import streamlit as st
import pandas as pd
from fuzzywuzzy import process
from io import BytesIO

# --- MENU DE NAVIGATION ---
st.sidebar.title("Navigation")
page = st.sidebar.radio("Aller à :", ["Accueil", "Mapping entre Fichiers"])

# Charger le fichier CSS
def load_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css("style.css")

# --- PAGE D'ACCUEIL ---
if page == "Accueil":
    st.markdown(
    """
    <div style="display: flex; align-items: center;">
        <img src="https://www.pngall.com/wp-content/uploads/15/Excel-Logo-PNG-Cutout.png" alt="Logo Excel" width="60" style="margin-right: 15px;">
        <h1 style="margin: 0;">Mapping entre Fichiers Excel</h1>
    </div>
    <p style="margin-top: 10px;">Utilisez le menu à gauche pour accéder à l'interface de mapping entre deux fichiers Excel.</p>
    """,
    unsafe_allow_html=True
)

# --- INTERFACE DE MAPPING ---
elif page == "Mapping entre Fichiers":
    st.title("Mapping entre Fichiers Excel")

    uploaded_file1 = st.file_uploader("Téléchargez le premier fichier Excel", type=["xlsx"])
    uploaded_file2 = st.file_uploader("Téléchargez le second fichier Excel", type=["xlsx"])

    df1, df2 = None, None

    if uploaded_file1 and uploaded_file2:
        df1 = pd.read_excel(uploaded_file1)
        df2 = pd.read_excel(uploaded_file2)

        st.subheader("Sélectionnez les colonnes pour le mappage")
        col_mapping_file1 = st.selectbox("Colonne du premier fichier :", options=df1.columns)
        col_mapping_file2 = st.selectbox("Colonne du second fichier :", options=df2.columns)

        df1['colonne_normalisee'] = df1[col_mapping_file1].astype(str).str.replace(r'[-_]', '', regex=True).str.lower()
        df2['colonne_normalisee'] = df2[col_mapping_file2].astype(str).str.replace(r'[-_]', '', regex=True).str.lower()

        def match_values_with_score(value, list_values):
            if pd.isna(value):
                return None, 0
            match, score = process.extractOne(value, list_values)
            return match, score

        df1[['valeur_matching', 'matching_score']] = df1['colonne_normalisee'].apply(
            lambda x: pd.Series(match_values_with_score(x, df2['colonne_normalisee'].tolist()))
        )

        seuil_acceptation = 50
        df1['match_acceptable'] = df1['matching_score'] >= seuil_acceptation

        st.subheader("Colonnes à inclure dans les résultats")
        columns_to_include = st.multiselect("Colonnes de df2 :", options=df2.columns.tolist(), default=df2.columns.tolist())

        result = pd.merge(
            df1,
            df2[columns_to_include],
            left_on='valeur_matching',
            right_on='colonne_normalisee',
            how='left',
            suffixes=('_fichier1', '_fichier2')
        )

        st.subheader("Résultats du Matching")
        st.write(result)

        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Résultats Matching')
            return output.getvalue()

        st.download_button(
            label="📥 Télécharger les résultats",
            data=convert_df_to_excel(result),
            file_name="resultat_matching.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Veuillez télécharger les deux fichiers Excel pour commencer.")
