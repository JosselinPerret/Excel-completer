import pandas as pd
import re
import streamlit as st
import os

def extract_coverage_from_report(text_file_path):
    coverage_data = {}
    
    with open(text_file_path, 'r', encoding='utf-8', errors='ignore') as file:
        text_content = file.read()
    
    # Patterns pour trouver les informations de coverage dans le rapport
    # Cherche les patterns comme "Test Summary for U10 (MC14519)" suivi de "Totals:" et "100.00%"
    component_blocks = re.findall(r'Test Summary for ([A-Z]+\d+).*?Totals:.*?(\d+\.\d+)%', 
                                  text_content, re.DOTALL)
    
    for component, coverage in component_blocks:
        coverage_data[component] = f"{float(coverage):.2f}%"
    
    return coverage_data

def update_excel_with_coverage(excel_file_path, coverage_data, output_file=None):
    if excel_file_path.endswith('.csv'):
        df = pd.read_csv(excel_file_path)
    else:
        df = pd.read_excel(excel_file_path)
    
    if "COVERAGE %" not in df.columns:
        df["COVERAGE %"] = None
    
    for index, row in df.iterrows():
        component = row["COMP."]
        if component in coverage_data:
            df.at[index, "COVERAGE %"] = coverage_data[component]
        else:
            df.at[index, "COVERAGE %"] = "0%"
    
    if output_file is None:
        base_name = os.path.splitext(excel_file_path)[0]
        output_file = f"{base_name}_updated.xlsx"
    
    df.to_excel(output_file, index=False)
    
    return df, output_file

def main():
    st.set_page_config(page_title="Coverage Excel Update", layout="wide")
    
    st.title("Coverage Excel Update")
    
    st.markdown("""
    Cette application permet de mettre à jour un fichier Excel avec des données de couverture de test.
    """)
    
    # Création de deux colonnes pour les fichiers
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Sélectionner le fichier Excel/CSV")
        excel_file = st.file_uploader("Choisissez un fichier Excel ou CSV", type=["xlsx", "csv"])
        
    with col2:
        st.subheader("Sélectionner le rapport de couverture")
        text_file = st.file_uploader("Choisissez un fichier texte de rapport", type=["txt"])
    
    if excel_file is not None and text_file is not None:
        # Sauvegarde des fichiers téléchargés temporairement
        excel_temp_path = os.path.join(os.getcwd(), excel_file.name)
        text_temp_path = os.path.join(os.getcwd(), text_file.name)
        
        with open(excel_temp_path, "wb") as f:
            f.write(excel_file.getbuffer())
            
        with open(text_temp_path, "wb") as f:
            f.write(text_file.getbuffer())
            
        # Bouton pour traiter les fichiers
        if st.button("Traiter les fichiers", type="primary"):
            with st.spinner("Traitement des fichiers en cours..."):
                try:
                    coverage_data = extract_coverage_from_report(text_temp_path)
                    
                    if not coverage_data:
                        st.error("Aucune donnée de couverture trouvée dans le fichier texte")
                    else:
                        updated_df, output_file = update_excel_with_coverage(excel_temp_path, coverage_data)
                        
                        # Affichage d'un aperçu du résultat
                        st.success(f"Traitement terminé! Fichier sauvegardé sous: {output_file}")
                        
                        # Afficher l'aperçu du DataFrame
                        st.subheader("Aperçu du fichier mis à jour")
                        st.dataframe(updated_df)
                        
                        # Téléchargement du fichier mis à jour
                        with open(output_file, "rb") as f:
                            st.download_button(
                                label="Télécharger le fichier mis à jour",
                                data=f,
                                file_name=os.path.basename(output_file),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                except Exception as e:
                    st.error(f"Erreur lors du traitement: {str(e)}")
        
        # Nettoyage des fichiers temporaires
        st.text("Note: Les fichiers téléchargés seront automatiquement supprimés après utilisation.")

if __name__ == "__main__":
    main()
