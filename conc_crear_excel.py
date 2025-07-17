import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import pandas as pd
import json
from openpyxl import load_workbook

def afegir_json_a_excel():
    # 1. Agafar el text del JSON
    text = text_area.get("1.0", tk.END).strip()
    if not text:
        messagebox.showwarning("Atenció", "El camp de text està buit.")
        return

    try:
        data = json.loads(text)

        # Transformem les llistes de dicts o strings en textos llegibles
        def llista_dicts_a_text(llista):
            return "\n".join([f"{d['titol']} ({d['experiencia']}): {d['funcions']}" for d in llista])

        def llista_a_text(llista):
            return "\n".join(llista)

        dades_planes = {
            "Nom del projecte": data.get("nom_projecte"),
            "Ubicació": data.get("ubicacio_projecte"),
            "Pressupost": data.get("pressupost_licitacio_PEM"),
            "Data licitació": data.get("data_licitacio"),
            "Perfils tècnics": llista_dicts_a_text(data.get("perfils_tecnics_requerits", [])),
            "Termini execució": data.get("termini_execucio"),
            "Requisits legals/tècnics": llista_a_text(data.get("requisits_legals_tecnics", [])),
            "Documentació a aportar": llista_a_text(data.get("documentacio_aportar", []))
        }

        df = pd.DataFrame([dades_planes])

    except Exception as e:
        messagebox.showerror("Error JSON", f"No s'ha pogut llegir el JSON:\n{e}")
        return

    # 2. Seleccionar fitxer Excel
    excel_path = filedialog.askopenfilename(
        title="Selecciona l'Excel on afegir les dades",
        filetypes=[("Fitxers Excel", "*.xlsx")]
    )
    if not excel_path:
        return

    # 3. Obrir l’Excel i demanar el nom de la pestanya (full)
    try:
        book = load_workbook(excel_path)
        fulls = book.sheetnames
    except Exception as e:
        messagebox.showerror("Error Excel", f"No s'ha pogut obrir l'Excel:\n{e}")
        return

    # Demanar a l'usuari el full (pestanya)
    full = simpledialog.askstring("Nom de la pestanya", f"Tria una pestanya d'aquest Excel:\n{', '.join(fulls)}")
    if full not in fulls:
        messagebox.showerror("Error", f"La pestanya '{full}' no existeix a l'Excel.")
        return

    try:
        # Llegeix les columnes existents del full d'Excel
        full_df = pd.read_excel(excel_path, sheet_name=full, nrows=1)
        cols_excel = list(full_df.columns)

        # Afegir columnes buides si no hi són al DataFrame
        for col in cols_excel:
            if col not in df.columns:
                df[col] = ""

        # Reordenar les columnes segons l'ordre del full d'Excel
        df = df[cols_excel]

        # Escriure les dades a l'Excel a partir de la següent fila disponible
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = book[full].max_row
            df.to_excel(writer, sheet_name=full, index=False, header=False, startrow=startrow)

        messagebox.showinfo("Fet!", "Les dades s'han afegit correctament a l'Excel.")
        text_area.delete("1.0", tk.END)
    except Exception as e:
        messagebox.showerror("Error", f"No s'han pogut afegir les dades:\n{e}")

# Interfície gràfica
root = tk.Tk()
root.title("Enganxar JSON i afegir a Excel")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

label = tk.Label(frame, text="Enganxa aquí el teu JSON:")
label.pack()

text_area = scrolledtext.ScrolledText(frame, width=80, height=20)
text_area.pack()

btn_afegir = tk.Button(frame, text="📤 Afegir al Excel", command=afegir_json_a_excel)
btn_afegir.pack(pady=10)

root.mainloop()