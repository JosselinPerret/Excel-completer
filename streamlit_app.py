import pandas as pd
import re
import streamlit as st
import os
import openpyxl
import io
from openpyxl.styles import PatternFill

def extract_data_from_report(text_file):
    coverage_data = {}
    notest_data = {}
    pmsg_not_used = []
    pass_tests = {}  # Nouveau dictionnaire pour stocker les tests PASS
    component_test_counts = {}  # Pour compter le nombre de tests par composant
    
    text_content = text_file.getvalue().decode('utf-8', errors='ignore')
    
    # 1. Extraire les données de couverture (Test Summary)
    component_blocks = re.findall(r'Test Summary for ([A-Z]+\d+).*?Totals:.*?(\d+\.\d+)%', 
                                  text_content, re.DOTALL)
    
    for component, coverage in component_blocks:
        coverage_data[component] = float(coverage)
    
    # 2. Extraire les composants NOTEST et PMSG not used depuis la section Untested Devices
    untested_section = re.search(r'Untested Devices(.*?)(?=General Summary Report|\Z)', text_content, re.DOTALL)
    
    if untested_section:
        untested_content = untested_section.group(1)
        
        # Chercher les lignes avec "COMPONENT IS TESTED IN PARALLEL WITH" suivi de NOTEST
        notest_matches = re.findall(r'([A-Z]+\d+)\s+\(COMPONENT IS TESTED IN PARALLEL WITH ([A-Z]+\d+)\)\s+NOTEST', untested_content)
        for component, tested_with in notest_matches:
            comment = f"COMPONENT IS TESTED IN PARALLEL WITH {tested_with}"
            notest_data[component] = comment
        
        # Chercher les composants avec "PMSG is not used"
        pmsg_matches = re.findall(r'([A-Z]+\d+)\s+\(PMSG is not used\)', untested_content)
        pmsg_not_used.extend(pmsg_matches)
    
    # 3. Extraire les blocs de test avec PASS/FAIL
    test_blocks = re.findall(r'\*([A-Z]+\d+(?:_[A-Z]+\d*)*)\s+Units.*?(?:PASS|FAIL)\s+LBF', text_content, re.DOTALL)
    
    # Compter les occurrences de chaque composant principal
    for test_id in test_blocks:
        # Extrayez le composant principal (avant le premier underscore)
        main_component = test_id.split('_')[0]
        if main_component not in component_test_counts:
            component_test_counts[main_component] = 0
        component_test_counts[main_component] += 1
    
    # Trouver les composants qui n'ont qu'un seul test et qui sont PASS
    for test_block in re.finditer(r'\*([A-Z]+\d+(?:_[A-Z]+\d*)*)\s+Units.*?((?:PASS|FAIL)\s+LBF)', text_content, re.DOTALL):
        test_id = test_block.group(1)
        result = test_block.group(2).strip()
        
        main_component = test_id.split('_')[0]
        if component_test_counts[main_component] == 1 and result.startswith("PASS"):
            pass_tests[main_component] = True
    
    return {
        "coverage": coverage_data,
        "notest": notest_data,
        "pmsg_not_used": pmsg_not_used,
        "pass_tests": pass_tests  # Ajout des tests PASS uniques
    }

def update_excel_with_data(excel_file, report_data):
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file)
    else:
        df = pd.read_excel(excel_file)
    
    # S'assurer que les colonnes nécessaires existent
    if "COVERAGE %" not in df.columns:
        df["COVERAGE %"] = None
    if "PPVS" not in df.columns:
        df["PPVS"] = None
    if "REMARQUE" not in df.columns:
        df["REMARQUE"] = None
    
    coverage_data = report_data["coverage"]
    notest_data = report_data["notest"]
    pmsg_not_used = report_data["pmsg_not_used"]
    pass_tests = report_data.get("pass_tests", {})  # Récupérer les tests PASS uniques
    
    # Créer un dictionnaire pour stocker les informations de formatage
    format_info = {}
    
    # Créer des ensembles pour suivre les composants déjà traités dans le rapport
    processed_components = set()
    
    for index, row in df.iterrows():
        if "COMP." not in row or pd.isna(row["COMP."]):
            continue
            
        component = row["COMP."]
        
        # 1. Traitement des composants avec couverture
        if component in coverage_data:
            df.at[index, "COVERAGE %"] = f"{coverage_data[component]:.2f}%"
            df.at[index, "PPVS"] = "OK"
            # Mémoriser les cellules à formater en vert
            format_info[index] = {"column": "PPVS", "color": "green"}
            processed_components.add(component)
            
        # 2. Traitement des composants NOTEST (SOUS-TEST) - composants testés en parallèle
        elif component in notest_data:
            df.at[index, "PPVS"] = "SOUS-TEST"
            df.at[index, "REMARQUE"] = notest_data[component]
            # Laisser la colonne COVERAGE % inchangée si elle a déjà une valeur
            if pd.isna(df.at[index, "COVERAGE %"]) or df.at[index, "COVERAGE %"] is None:
                df.at[index, "COVERAGE %"] = None
            # Mémoriser les cellules à formater en jaune clair
            format_info[index] = {"column": "PPVS", "color": "yellow"}
            processed_components.add(component)
            
        # 3. Traitement des composants PMSG not used (NOTEST)
        elif component in pmsg_not_used:
            df.at[index, "PPVS"] = "NOTEST"
            # Laisser la colonne COVERAGE % inchangée si elle a déjà une valeur
            if pd.isna(df.at[index, "COVERAGE %"]) or df.at[index, "COVERAGE %"] is None:
                df.at[index, "COVERAGE %"] = None
            # Mémoriser les cellules à formater en rouge clair
            format_info[index] = {"column": "PPVS", "color": "red"}
            processed_components.add(component)
        
        # 4. Traitement des composants avec un test unique PASS
        elif component in pass_tests:
            df.at[index, "PPVS"] = "OK"
            # Mémoriser les cellules à formater en vert
            format_info[index] = {"column": "PPVS", "color": "green"}
            processed_components.add(component)
    
    # Afficher les statistiques des données traitées
    st.write(f"Composants avec couverture trouvés: {len(coverage_data)}")
    st.write(f"Composants SOUS-TEST trouvés: {len(notest_data)}")
    st.write(f"Composants NOTEST (PMSG not used) trouvés: {len(pmsg_not_used)}")
    st.write(f"Composants avec test unique PASS trouvés: {len(pass_tests)}")
    st.write(f"Composants traités dans le tableau Excel: {len(processed_components)}")
    
    return df, format_info

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
                    # Réinitialiser le curseur du fichier texte
                    text_file.seek(0)
                    report_data = extract_data_from_report(text_file)
                    
                    # Afficher un résumé des données extraites
                    st.write("Données extraites du rapport:")
                    st.write(f"- Composants avec couverture: {len(report_data['coverage'])}")
                    st.write(f"- Composants NOTEST (SOUS-TEST): {len(report_data['notest'])}")
                    st.write(f"- Composants PMSG not used (NOTEST): {len(report_data['pmsg_not_used'])}")
                    st.write(f"- Composants avec test unique PASS: {len(report_data.get('pass_tests', {}))}")
                    
                    if not report_data["coverage"] and not report_data["notest"] and not report_data["pmsg_not_used"]:
                        st.error("Aucune donnée trouvée dans le fichier texte")
                    else:
                        # Réinitialiser le curseur du fichier Excel
                        excel_file.seek(0)
                        updated_df, format_info = update_excel_with_data(excel_file, report_data)
                        
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
                            
                            # Trouver les indices de colonnes pour COVERAGE %, PPVS et REMARQUE
                            coverage_col_idx = None
                            ppvs_col_idx = None
                            remarque_col_idx = None
                            
                            try:
                                coverage_col_idx = header_row.index("COVERAGE %") + 1  # +1 car openpyxl est indexé à partir de 1
                            except ValueError:
                                coverage_col_idx = len(header_row) + 1
                                sheet.cell(row=1, column=coverage_col_idx, value="COVERAGE %")
                                
                            try:
                                ppvs_col_idx = header_row.index("PPVS") + 1
                            except ValueError:
                                ppvs_col_idx = len(header_row) + 2
                                sheet.cell(row=1, column=ppvs_col_idx, value="PPVS")
                                
                            try:
                                remarque_col_idx = header_row.index("REMARQUE") + 1
                            except ValueError:
                                remarque_col_idx = len(header_row) + 3
                                sheet.cell(row=1, column=remarque_col_idx, value="REMARQUE")
                            
                            # Style de remplissage pour chaque couleur
                            from openpyxl.styles import PatternFill
                            green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # Vert clair
                            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # Jaune clair
                            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")    # Rouge clair
                            
                            # Mettre à jour les cellules
                            for i, row in enumerate(sheet.iter_rows(min_row=2)):
                                if i < len(updated_df):
                                    # Mettre à jour la colonne COVERAGE %
                                    coverage_str = updated_df.iloc[i]["COVERAGE %"]
                                    if coverage_str is not None and pd.notna(coverage_str):
                                        if str(coverage_str) != "0%":
                                            try:
                                                coverage_value = float(str(coverage_str).replace('%', ''))
                                                cell = sheet.cell(row=i+2, column=coverage_col_idx, value=coverage_value/100.0)
                                                cell.number_format = '0.00%'
                                            except (ValueError, AttributeError):
                                                sheet.cell(row=i+2, column=coverage_col_idx, value=coverage_str)
                                        else:
                                            cell = sheet.cell(row=i+2, column=coverage_col_idx, value=0)
                                            cell.number_format = '0.00%'
                                    
                                    # Mettre à jour la colonne PPVS
                                    ppvs_value = updated_df.iloc[i]["PPVS"]
                                    if ppvs_value is not None and pd.notna(ppvs_value):
                                        cell = sheet.cell(row=i+2, column=ppvs_col_idx, value=ppvs_value)
                                        
                                        # Appliquer le formatage selon format_info
                                        if i in format_info:
                                            if format_info[i]["color"] == "green":
                                                cell.fill = green_fill
                                            elif format_info[i]["color"] == "yellow":
                                                cell.fill = yellow_fill
                                            elif format_info[i]["color"] == "red":
                                                cell.fill = red_fill
                                    
                                    # Mettre à jour la colonne REMARQUE
                                    remarque_value = updated_df.iloc[i]["REMARQUE"]
                                    if remarque_value is not None and pd.notna(remarque_value):
                                        sheet.cell(row=i+2, column=remarque_col_idx, value=remarque_value)
                                        sheet.cell(row=i+2, column=remarque_col_idx, value=remarque_value)
                            
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
