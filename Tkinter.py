import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

def xlsx_to_csv(excel_file):
    df = pd.read_excel(excel_file)
    
    csv_file = excel_file.replace('.xlsx', '.csv')
    df.to_csv(csv_file, index=False)
    
    return csv_file

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
        output_file = excel_file_path.replace('.xlsx', '_updated.xlsx')
    
    df.to_excel(output_file, index=False)
    
    return df

def main():
    root = tk.Tk()
    root.title("Coverage Excel Update")
    root.geometry("600x400")
    
    excel_path = tk.StringVar()
    text_path = tk.StringVar()
    status_message = tk.StringVar()
    status_message.set("Selectionnez les fichiers pour commencer")
    
    file_frame = tk.Frame(root, pady=20)
    file_frame.pack(fill="x")
    
    tk.Label(file_frame, text="Excel/CSV:").grid(row=0, column=0, padx=5, sticky="w")
    tk.Entry(file_frame, textvariable=excel_path, width=50).grid(row=0, column=1, padx=5)
    tk.Button(file_frame, text="Parcourir", command=lambda: excel_path.set(tk.filedialog.askopenfilename(
        filetypes=[("Fichiers Excel", "*.xlsx"), ("Fichiers CSV", "*.csv")]))).grid(row=0, column=2, padx=5)
    
    tk.Label(file_frame, text="Coverage Report:").grid(row=1, column=0, padx=5, sticky="w")
    tk.Entry(file_frame, textvariable=text_path, width=50).grid(row=1, column=1, padx=5)
    tk.Button(file_frame, text="Parcourir", command=lambda: text_path.set(tk.filedialog.askopenfilename(
        filetypes=[("Fichiers texte", "*.txt")]))).grid(row=1, column=2, padx=5)

    def process_files():
        try:
            if not excel_path.get() or not text_path.get():
                status_message.set("Veuillez sélectionner les deux fichiers")
                return
                
            coverage_data = extract_coverage_from_report(text_path.get())
            
            if not coverage_data:
                status_message.set("Aucune donnée de couverture trouvée dans le fichier texte")
                return
                
            updated_df = update_excel_with_coverage(excel_path.get(), coverage_data)
            
            status_message.set(f"Succès! Fichier Excel mis à jour enregistré sous '{excel_path.get().replace('.xlsx', '_updated.xlsx')}'")
        except Exception as e:
            status_message.set(f"Erreur: {str(e)}")
    
    tk.Button(root, text="Traiter les fichiers", command=process_files, 
                bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), 
                pady=10, padx=20).pack(pady=20)
    
    status_frame = tk.Frame(root, pady=10)
    status_frame.pack(fill="x")
    tk.Label(status_frame, text="Statut:").pack(anchor="w", padx=20)
    tk.Label(status_frame, textvariable=status_message, wraplength=500,
                justify="left", fg="blue").pack(anchor="w", padx=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()