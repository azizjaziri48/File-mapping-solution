import streamlit as st
import pandas as pd
from fuzzywuzzy import process
from io import BytesIO

# Titre de l'application
st.title("Matching entre Fichiers Excel avec Sélection de Colonnes pour le Mappage")

# Téléchargement des fichiers Excel
uploaded_file1 = st.file_uploader("Téléchargez le premier fichier Excel", type=["xlsx"])
uploaded_file2 = st.file_uploader("Téléchargez le second fichier Excel", type=["xlsx"])

# Initialisation des DataFrames
df1, df2 = None, None

if uploaded_file1 and uploaded_file2:
    # Lecture des fichiers
    df1 = pd.read_excel(uploaded_file1)
    df2 = pd.read_excel(uploaded_file2)
    
    # Sélection des colonnes pour le mapping
    st.subheader("Sélectionnez les colonnes pour le mappage")
    
    col_mapping_file1 = st.selectbox("Sélectionnez une colonne du premier fichier pour le mappage :", options=df1.columns)
    col_mapping_file2 = st.selectbox("Sélectionnez une colonne du second fichier pour le mappage :", options=df2.columns)
    
    # Normalisation des colonnes sélectionnées
    df1['colonne_normalisee'] = df1[col_mapping_file1].astype(str).str.replace(r'[-_]', '', regex=True).str.lower()
    df2['colonne_normalisee'] = df2[col_mapping_file2].astype(str).str.replace(r'[-_]', '', regex=True).str.lower()
    
    # Fonction pour trouver les meilleures correspondances avec un score
    def match_values_with_score(value, list_values):
        if pd.isna(value):
            return None, 0
        match, score = process.extractOne(value, list_values)
        return match, score
    
    # Appliquer le matching flou
    df1[['valeur_matching', 'matching_score']] = df1['colonne_normalisee'].apply(
        lambda x: pd.Series(match_values_with_score(x, df2['colonne_normalisee'].tolist()))
    )

    # Filtrer les correspondances avec un score élevé (optionnel)
    seuil_acceptation = 85
    df1['match_acceptable'] = df1['matching_score'] >= seuil_acceptation

    # Permettre à l'utilisateur de sélectionner les colonnes à restituer dans df2
    st.subheader("Sélectionnez les colonnes du second fichier à inclure dans les résultats")
    columns_to_include = st.multiselect(
        "Colonnes disponibles dans le second fichier (df2) :", 
        options=df2.columns.tolist(), 
        default=df2.columns.tolist()  # Par défaut, toutes les colonnes sont sélectionnées
    )

    # Fusionner les DataFrames selon les correspondances trouvées
    result = pd.merge(
        df1, 
        df2[columns_to_include + ['colonne_normalisee']],  # Inclure uniquement les colonnes sélectionnées
        left_on='valeur_matching', 
        right_on='colonne_normalisee', 
        how='left', 
        suffixes=('_fichier1', '_fichier2')
    )

    # Affichage du résultat dans l'interface
    st.subheader("Résultats du Matching")
    st.write(result)

    # Fonction pour télécharger le fichier Excel résultant
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Résultats Matching')
        processed_data = output.getvalue()
        return processed_data

    # Bouton pour télécharger les résultats
    st.download_button(
        label="Télécharger le fichier Excel des résultats",
        data=convert_df_to_excel(result),
        file_name="resultat_matching.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Veuillez télécharger les deux fichiers Excel pour commencer.")
