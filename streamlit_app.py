import pandas as pd
import re
import streamlit as st
import os

def extract_coverage_from_report(text_file):
    coverage_data = {}
    
    text_content = text_file.getvalue().decode('utf-8', errors='ignore')
    
    # Patterns pour trouver les informations de coverage dans le rapport
    # Cherche les patterns comme "Test Summary for U10 (MC14519)" suivi de "Totals:" et "100.00%"
    component_blocks = re.findall(r'Test Summary for ([A-Z]+\d+).*?Totals:.*?(\d+\.\d+)%', 
                                  text_content, re.DOTALL)
    
    for component, coverage in component_blocks:
        coverage_data[component] = float(coverage)
    
    return coverage_data

def update_excel_with_coverage(excel_file, coverage_data):
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file)
    else:
        df = pd.read_excel(excel_file)
    
    if "COVERAGE %" not in df.columns:
        df["COVERAGE %"] = None
    
    for index, row in df.iterrows():
        component = row["COMP."]
        if component in coverage_data:
            df.at[index, "COVERAGE %"] = f"{coverage_data[component]:.2f}%"
        else:
            df.at[index, "COVERAGE %"] = "0%"
    
    return df

def main():
    st.set_page_config(page_title="Coverage Excel Update", layout="wide")
    
    st.title("Coverage Excel Update")
    
    st.markdown("""
    Cette application permet de mettre à jour un fichier Excel avec des données de couverture de test.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Sélectionner le fichier Excel/CSV")
        excel_file = st.file_uploader("Choisissez un fichier Excel ou CSV", type=["xlsx", "csv"])
        
    with col2:
        st.subheader("Sélectionner le rapport de couverture")
        text_file = st.file_uploader("Choisissez un fichier texte de rapport", type=["txt"])
    
    if excel_file is not None and text_file is not None:
        import tempfile
            
        if st.button("Traiter les fichiers", type="primary"):
            with st.spinner("Traitement des fichiers en cours..."):
                try:
                    coverage_data = extract_coverage_from_report(text_file)
                    
                    if not coverage_data:
                        st.error("Aucune donnée de couverture trouvée dans le fichier texte")
                    else:
                        updated_df = update_excel_with_coverage(excel_file, coverage_data)
                        
                        st.success(f"Traitement terminé!")
                        
                        st.subheader("Aperçu du fichier mis à jour")
                        st.dataframe(updated_df)
                        
                        import io
                        import openpyxl
                        
                        if excel_file.name.endswith('.xlsx'):
                            excel_file.seek(0)
                            
                            wb = openpyxl.load_workbook(excel_file)
                            sheet_name = wb.sheetnames[0]
                            sheet = wb[sheet_name]
                            
                            header_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
                            try:
                                coverage_col_idx = header_row.index("COVERAGE %") + 1  # +1 car openpyxl est indexé à partir de 1
                            except ValueError:
                                coverage_col_idx = len(header_row) + 1
                                sheet.cell(row=1, column=coverage_col_idx, value="COVERAGE %")
                            
                            for i, row in enumerate(sheet.iter_rows(min_row=2)):
                                if i < len(updated_df):
                                    comp = updated_df.iloc[i]["COMP."]
                                    coverage_str = updated_df.iloc[i]["COVERAGE %"]
                                    
                                    if coverage_str != "0%":
                                        try:
                                            coverage_value = float(coverage_str.replace('%', ''))
                                            cell = sheet.cell(row=i+2, column=coverage_col_idx, value=coverage_value/100.0)
                                            
                                            cell.number_format = '0.00%'
                                        except (ValueError, AttributeError):
                                            sheet.cell(row=i+2, column=coverage_col_idx, value=coverage_str)
                                    else:
                                        cell = sheet.cell(row=i+2, column=coverage_col_idx, value=0)
                                        cell.number_format = '0.00%'
                            
                            output = io.BytesIO()
                            wb.save(output)
                            output.seek(0)
                            
                            file_name = excel_file.name.replace('.xlsx', '_updated.xlsx')
                            
                            st.download_button(
                                label="Télécharger le fichier mis à jour",
                                data=output,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            # Pour les CSV, utiliser pandas
                            csv_buffer = io.StringIO()
                            updated_df.to_csv(csv_buffer, index=False)
                            
                            file_name = excel_file.name.replace('.csv', '_updated.csv')
                            
                            st.download_button(
                                label="Télécharger le fichier mis à jour",
                                data=csv_buffer.getvalue(),
                                file_name=file_name,
                                mime="text/csv"
                            )
                        
                except Exception as e:
                    st.error(f"Erreur lors du traitement: {str(e)}")
        
        st.text("Note: Les fichiers sont traités directement en mémoire et ne sont pas sauvegardés sur le serveur.")

if __name__ == "__main__":
    main()
