# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------
# Quadre de Disponibilitat de Professors
#
# Genera un quadre visual dels professors disponibles en diferents dies 
# i hores per a grups específics per facilitar la planificació de sortides.
#
# Autor: Josep M.
# Versió: 3.0
# Data: 14 d'abril de 2026
# Llicència: Creative Commons BY-NC-SA 4.0
#
# Canvis:
# - v3.0: Afegida capçalera d'autoria i llicència segons la plantilla.
# - v2.0: Versió funcional amb selecció de múltiples dies i hores.
# ----------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import webbrowser
import csv

# --- Funcions Lògiques ---

def obtenir_cursos_de_fitxer(ruta_fitxer):
    if not ruta_fitxer or "Cap fitxer" in ruta_fitxer:
        return []
    try:
        with open(ruta_fitxer, 'r', encoding='utf-8') as f:
            linies = f.readlines()
        cursos = set()
        for linia in linies[1:]:
            parts = linia.strip().split(',')
            if len(parts) > 1:
                grup = parts[1].strip().strip('"')
                if grup: cursos.add(grup)
        return sorted(list(cursos))
    except Exception as e:
        messagebox.showerror("Error", f"No s'ha pogut llegir el fitxer: {e}")
        return []

def generar_dades_quadre(ruta_fitxer, dies_sel, hores_sel, cursos_sel):
    dia_to_num = {"dilluns": "1", "dimarts": "2", "dimecres": "3", "dijous": "4", "divendres": "5"}
    hora_to_num = {"1a": "1", "2a": "2", "3a": "3", "4a": "4", "5a": "5", "6a": "6", "7a": "7"}
    
    try:
        with open(ruta_fitxer, 'r', encoding='utf-8') as f:
            linies = f.readlines()
    except Exception as e:
        return None, f"Error: {e}"

    dades = {h_num: {d_num: [] for d_num in [dia_to_num[d.lower()] for d in dies_sel]} 
             for h_num in [hora_to_num[h] for h in hores_sel]}
    
    dies_numerics = [dia_to_num[d.lower()] for d in dies_sel]
    hores_numeriques = [hora_to_num[h] for h in hores_sel]

    for linia in linies[1:]:
        parts = linia.strip().split(',')
        if len(parts) >= 7:
            grup = parts[1].strip().strip('"')
            prof = parts[2].strip().strip('"')
            dia = parts[5].strip()
            hora = parts[6].strip()
            
            if grup in cursos_sel and dia in dies_numerics and hora in hores_numeriques:
                if prof not in dades[hora][dia]:
                    dades[hora][dia].append(prof)
    
    return dades, ""

def formatar_text_quadre(dades, dies_sel, hores_sel):
    dia_to_num = {"dilluns": "1", "dimarts": "2", "dimecres": "3", "dijous": "4", "divendres": "5"}
    hora_to_num = {"1a": "1", "2a": "2", "3a": "3", "4a": "4", "5a": "5", "6a": "6", "7a": "7"}
    
    col_width = 25
    hora_width = 10
    output = []
    
    header = f"{'HORA':<{hora_width}} | " + " | ".join([f"{d:^{col_width}}" for d in dies_sel])
    separator = "-" * (hora_width + 3 + (col_width + 3) * len(dies_sel))
    output.append(header)
    output.append(separator)

    for h_ui in hores_sel:
        h_num = hora_to_num[h_ui]
        max_profs = max([len(dades[h_num][dia_to_num[d.lower()]]) for d in dies_sel])
        
        if max_profs == 0:
            row = f"{h_ui:<{hora_width}} | " + " | ".join([f"{'---':^{col_width}}" for _ in dies_sel])
            output.append(row)
        else:
            for i in range(max_profs):
                row_label = f"{h_ui:<{hora_width}} | " if i == 0 else f"{' ':<{hora_width}} | "
                cells = []
                for d in dies_sel:
                    d_num = dia_to_num[d.lower()]
                    llista_p = sorted(dades[h_num][d_num])
                    prof_nom = llista_p[i] if i < len(llista_p) else ""
                    cells.append(f"{prof_nom:^{col_width}}")
                output.append(row_label + " | ".join(cells))
        output.append(separator)
    
    return "\n".join(output)

# --- GUI ---

def carregar_fitxer():
    ruta = filedialog.askopenfilename(title="Selecciona Horaris.TXT", filetypes=(("Text", "*.txt"), ("Tots", "*.*")))
    if ruta:
        ruta_var.set(ruta)
        cursos = obtenir_cursos_de_fitxer(ruta)
        llista_cursos.delete(0, tk.END)
        for c in cursos: llista_cursos.insert(tk.END, c)

def executar():
    ruta = ruta_var.get()
    dies = [llista_dies.get(i) for i in llista_dies.curselection()]
    hores = [llista_hores.get(i) for i in llista_hores.curselection()]
    cursos = [llista_cursos.get(i) for i in llista_cursos.curselection()]

    if not ruta or "Cap fitxer" in ruta or not dies or not hores or not cursos:
        messagebox.showwarning("Atenció", "Selecciona fitxer, dies, hores i grups.")
        return

    dades, err = generar_dades_quadre(ruta, dies, hores, cursos)
    if err:
        messagebox.showerror("Error", err)
        return

    res_text = formatar_text_quadre(dades, dies, hores)
    text_res.config(state=tk.NORMAL)
    text_res.delete('1.0', tk.END)
    text_res.insert(tk.END, f"GRUPS: {', '.join(cursos)}\n\n" + res_text)
    text_res.config(state=tk.DISABLED)
    # Guardem dades per a l'exportació
    app_state['dades'] = dades
    app_state['dies'] = dies
    app_state['hores'] = hores
    app_state['cursos'] = cursos

def guardar_resultats():
    if 'dades' not in app_state:
        messagebox.showwarning("Atenció", "Primer genera el quadre.")
        return

    ruta_salva = filedialog.asksaveasfilename(
        title="Guardar com...",
        filetypes=(("Markdown", "*.md"), ("Excel/CSV", "*.csv"), ("Text", "*.txt")),
        defaultextension=".md"
    )
    
    if not ruta_salva: return

    try:
        if ruta_salva.endswith('.csv'):
            with open(ruta_salva, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['HORA'] + app_state['dies'])
                dia_to_num = {"dilluns": "1", "dimarts": "2", "dimecres": "3", "dijous": "4", "divendres": "5"}
                hora_to_num = {"1a": "1", "2a": "2", "3a": "3", "4a": "4", "5a": "5", "6a": "6", "7a": "7"}
                
                for h_ui in app_state['hores']:
                    h_num = hora_to_num[h_ui]
                    max_profs = max([len(app_state['dades'][h_num][dia_to_num[d.lower()]]) for d in app_state['dies']])
                    if max_profs == 0:
                        writer.writerow([h_ui] + ['---'] * len(app_state['dies']))
                    else:
                        for i in range(max_profs):
                            row = [h_ui if i == 0 else ""]
                            for d in app_state['dies']:
                                d_num = dia_to_num[d.lower()]
                                llista = sorted(app_state['dades'][h_num][d_num])
                                row.append(llista[i] if i < len(llista) else "")
                            writer.writerow(row)
        else:
            contingut = text_res.get("1.0", tk.END)
            with open(ruta_salva, 'w', encoding='utf-8') as f:
                f.write(contingut)
        
        messagebox.showinfo("Èxit", "Fitxer guardat correctament.")
    except Exception as e:
        messagebox.showerror("Error en guardar", str(e))

root = tk.Tk()
root.title("Quadre Disponibilitat v03")
root.geometry("1100x800")

app_state = {}
ruta_var = tk.StringVar(value="Cap fitxer seleccionat.")

# --- Peu de pàgina (Footer) - Empaqueta primer a baix ---
footer_frame = tk.Frame(root, height=30, bg="#f0f0f0")
footer_frame.pack(side="bottom", fill="x")

autor = "Josep M."
llicencia_text = "CC BY-NC-SA 4.0"
url_llicencia = "https://creativecommons.org/licenses/by-nc-sa/4.0/"
info_text = f"Autor: {autor}  |  Llicència: {llicencia_text}"

footer_label = tk.Label(footer_frame, text=info_text, fg="blue", cursor="hand2", bg="#f0f0f0")
f_footer = font.Font(family="Arial", size=9, underline=True)
footer_label.configure(font=f_footer)
footer_label.bind("<Button-1>", lambda e: webbrowser.open_new(url_llicencia))
footer_label.pack(pady=5)

# Top
f_top = ttk.Frame(root, padding=10)
f_top.pack(fill=tk.X)
ttk.Button(f_top, text="Carregar Horaris.TXT", command=carregar_fitxer).pack(side=tk.LEFT)
ttk.Entry(f_top, textvariable=ruta_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

# Config
f_mid = ttk.Frame(root, padding=10)
f_mid.pack(fill=tk.X)

f_dia = ttk.LabelFrame(f_mid, text=" Dies (Mult.) ", padding=5)
f_dia.pack(side=tk.LEFT, padx=5)
llista_dies = tk.Listbox(f_dia, selectmode=tk.MULTIPLE, height=5, width=15, exportselection=False)
for d in ["Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres"]: llista_dies.insert(tk.END, d)
llista_dies.pack()

f_hores = ttk.LabelFrame(f_mid, text=" Hores (Mult.) ", padding=5)
f_hores.pack(side=tk.LEFT, padx=5)
llista_hores = tk.Listbox(f_hores, selectmode=tk.MULTIPLE, height=7, width=10, exportselection=False)
for h in ["1a", "2a", "3a", "4a", "5a", "6a", "7a"]: llista_hores.insert(tk.END, h)
llista_hores.pack()

f_curs = ttk.LabelFrame(f_mid, text=" Grups de la sortida ", padding=5)
f_curs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
llista_cursos = tk.Listbox(f_curs, selectmode=tk.MULTIPLE, height=7, exportselection=False)
scr_c = ttk.Scrollbar(f_curs, command=llista_cursos.yview)
llista_cursos.config(yscrollcommand=scr_c.set)
llista_cursos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scr_c.pack(side=tk.RIGHT, fill=tk.Y)

# Botons Acció
f_btns = ttk.Frame(root, padding=5)
f_btns.pack(fill=tk.X, padx=15)
ttk.Button(f_btns, text="GENERAR QUADRE v03", command=executar).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
ttk.Button(f_btns, text="GUARDAR RESULTATS (.md / .csv)", command=guardar_resultats).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

# Resultats (Amb expand=True per ocupar la resta)
f_res = ttk.LabelFrame(root, text=" Quadre Resultant (Files=Hores / Columnes=Dies) ", padding=10)
f_res.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
x_scroll = ttk.Scrollbar(f_res, orient=tk.HORIZONTAL)
y_scroll = ttk.Scrollbar(f_res, orient=tk.VERTICAL)
text_res = tk.Text(f_res, wrap=tk.NONE, state=tk.DISABLED, font=("Courier New", 9), height=15,
                   xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
x_scroll.config(command=text_res.xview)
y_scroll.config(command=text_res.yview)
y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
text_res.pack(fill=tk.BOTH, expand=True)

root.mainloop()
