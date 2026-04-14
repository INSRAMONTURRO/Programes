# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------
# Rànquing de Professors per a Sortides
#
# Analitza la disponibilitat de professors i genera un rànquing d'aquells 
# amb menys afectació per a una sortida escolar segons els grups seleccionats.
#
# Autor: Josep M.
# Versió: 3.0
# Data: 14 d'abril de 2026
# Llicència: Creative Commons BY-NC-SA 4.0
#
# Canvis:
# - v3.0: Afegida capçalera d'autoria i llicència segons la plantilla.
# - v2.0: Versió anterior funcional.
# ----------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import webbrowser
from collections import Counter

# --- Funcions Lògiques ---

def obtenir_cursos_de_fitxer(ruta_fitxer):
    if not ruta_fitxer or "Cap fitxer seleccionat" in ruta_fitxer:
        return []
    
    try:
        with open(ruta_fitxer, 'r', encoding='utf-8') as f:
            linies = f.readlines()
    except Exception as e:
        messagebox.showerror("Error", f"No s'ha pogut llegir el fitxer: {e}")
        return []

    cursos = set()
    for linia in linies[1:]:
        parts = linia.strip().split(',')
        if len(parts) > 1:
            grup_raw = parts[1].strip().strip('"')
            if grup_raw: 
                cursos.add(grup_raw)
    
    return sorted(list(cursos))


def analitzar_disponibilitat(ruta_fitxer, dia_seleccionat, hores_seleccionades, cursos_seleccionats):
    dia_a_numero = {"dilluns": "1", "dimarts": "2", "dimecres": "3", "dijous": "4", "divendres": "5"}
    dia_numeric = dia_a_numero.get(dia_seleccionat.lower())
    hora_a_numero = {"1a": "1", "2a": "2", "3a": "3", "4a": "4", "5a": "5", "6a": "6", "7a": "7"}

    # Paraules clau per identificar reunions, càrrecs o tasques no lectives
    REUNION_KEYWORDS = ['RN', 'JUNTA', 'CD_', 'JUN.', 'TUT_', 'COORD', 'DIRECC', 'CAP_', 'PREP']

    try:
        with open(ruta_fitxer, 'r', encoding='utf-8') as f:
            linies = f.readlines()
    except Exception as e:
        return f"Error en llegir el fitxer: {e}"

    professors_relevants = set()
    for linia in linies[1:]:
        parts = linia.strip().split(',')
        if len(parts) >= 6:
            grup = parts[1].strip().strip('"')
            professor = parts[2].strip().strip('"')
            dia = parts[5].strip()
            if dia == dia_numeric and grup in cursos_seleccionats and professor:
                professors_relevants.add(professor)

    if not professors_relevants:
        return "No s'han trobat professors per als cursos seleccionats en el dia indicat."

    horari_professors = {prof: {} for prof in professors_relevants}
    for linia in linies[1:]:
        parts = linia.strip().split(',')
        if len(parts) >= 8:
            grup = parts[1].strip().strip('"')
            professor = parts[2].strip().strip('"')
            activitat = parts[3].strip().strip('"')
            dia = parts[5].strip()
            hora = parts[6].strip()
            if dia == dia_numeric and professor in professors_relevants:
                horari_professors[professor][hora] = (activitat, grup)

    resultats_per_hora = {hora: {'lliures': [], 'classe_seleccionada': [], 'classe_altres': [], 'guardia': [], 'reunions': []} for hora in hores_seleccionades}
    recompte_afectacio = Counter()
    
    for hora_ui in hores_seleccionades:
        hora_num = hora_a_numero.get(hora_ui)
        for prof in sorted(list(professors_relevants)):
            info_activitat = horari_professors[prof].get(hora_num)
            
            if not info_activitat:
                resultats_per_hora[hora_ui]['lliures'].append(prof)
            else:
                activitat, grup_activitat = info_activitat
                activitat_upper = activitat.upper()

                if "GUARDIA" in activitat_upper:
                    resultats_per_hora[hora_ui]['guardia'].append(prof)
                elif any(keyword in activitat_upper for keyword in REUNION_KEYWORDS):
                    resultats_per_hora[hora_ui]['reunions'].append(prof)
                    recompte_afectacio[prof] += 1 
                else:
                    if grup_activitat in cursos_seleccionats:
                        resultats_per_hora[hora_ui]['classe_seleccionada'].append(prof)
                        recompte_afectacio[prof] += 1 
                    else:
                        resultats_per_hora[hora_ui]['classe_altres'].append((prof, grup_activitat))

    output = []
    output.append(f"INFORME D'IDONEÏTAT PER A SORTIDES ({dia_seleccionat.upper()})\n")
    output.append(f"Cursos: {', '.join(cursos_seleccionats)}\n")
    output.append("="*80 + "\n")

    output.append("--- RÀNQUING 10 PROFESSORS AMB MENYS AFECTACIÓ (RECOMANATS) ---")
    top_10 = recompte_afectacio.most_common(10)
    for i, (prof, punts) in enumerate(top_10, 1):
        output.append(f"{i}. {prof:.<40} {punts} hores coincidents")
    
    output.append("\n" + "="*80 + "\n")

    for hora_ui in sorted(resultats_per_hora.keys()):
        output.append(f"--- HORA: {hora_ui} ---")
        if resultats_per_hora[hora_ui]['classe_seleccionada']:
            output.append(f"  Amb els alumnes de la sortida: { ', '.join(resultats_per_hora[hora_ui]['classe_seleccionada'])}")
        if resultats_per_hora[hora_ui]['reunions']:
            output.append(f"  En reunió o càrrecs: { ', '.join(resultats_per_hora[hora_ui]['reunions'])}")
        if resultats_per_hora[hora_ui]['lliures']:
            output.append(f"  Sense classe assignada: { ', '.join(resultats_per_hora[hora_ui]['lliures'])}")
        output.append("")

    return '\n'.join(output)

# --- Interfície Gràfica ---

def seleccionar_fitxer_horaris():
    ruta = filedialog.askopenfilename(title="Selecciona Horaris.TXT", filetypes=(("Text", "*.txt"), ("Tots", "*.*")))
    if ruta:
        ruta_fitxer_var.set(ruta)
        cursos = obtenir_cursos_de_fitxer(ruta)
        llista_cursos.delete(0, tk.END)
        for curs in cursos: llista_cursos.insert(tk.END, curs)

def seleccionar_totes_les_hores(event=None):
    if len(llista_hores.curselection()) == llista_hores.size():
        llista_hores.selection_clear(0, tk.END)
    else:
        llista_hores.selection_set(0, tk.END)

def iniciar_analisi():
    ruta = ruta_fitxer_var.get()
    dia = dia_var.get()
    hores = [llista_hores.get(i) for i in llista_hores.curselection()]
    cursos = [llista_cursos.get(i) for i in llista_cursos.curselection()]

    if "Cap fitxer" in ruta or not hores or not cursos:
        messagebox.showwarning("Atenció", "Assegura't de seleccionar fitxer, hores i cursos.")
        return

    res = analitzar_disponibilitat(ruta, dia, hores, cursos)
    text_resultats.config(state=tk.NORMAL)
    text_resultats.delete('1.0', tk.END)
    text_resultats.insert(tk.END, res)
    text_resultats.config(state=tk.DISABLED)

app = tk.Tk()
app.title("Gestor de Sortides Escolars")
app.geometry("850x850")

ruta_fitxer_var = tk.StringVar(value="Cap fitxer seleccionat.")
dia_var = tk.StringVar(value="dilluns")

# --- Peu de pàgina (Footer) - Empaqueta primer a baix ---
footer_frame = tk.Frame(app, height=30, bg="#f0f0f0")
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

# Frame Dalt
f_top = ttk.Frame(app, padding=10)
f_top.pack(fill=tk.X)
ttk.Button(f_top, text="1. Carregar Horaris.TXT", command=seleccionar_fitxer_horaris).pack(side=tk.LEFT)
ttk.Entry(f_top, textvariable=ruta_fitxer_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

# Frame Selecció
f_mid = ttk.Frame(app, padding=10)
f_mid.pack(fill=tk.X)

# Dia
f_d = ttk.LabelFrame(f_mid, text=" Dia ", padding=5)
f_d.pack(side=tk.LEFT, padx=5, fill=tk.Y)
ttk.Combobox(f_d, textvariable=dia_var, values=["dilluns", "dimarts", "dimecres", "dijous", "divendres"], state="readonly", width=12).pack()

# Hores
f_h = ttk.LabelFrame(f_mid, text=" Hores (Totes) ", padding=5)
f_h.pack(side=tk.LEFT, padx=5, fill=tk.Y)
llista_hores = tk.Listbox(f_h, selectmode=tk.MULTIPLE, height=7, exportselection=False)
for h in ["1a", "2a", "3a", "4a", "5a", "6a", "7a"]: llista_hores.insert(tk.END, h)
llista_hores.pack()
llista_hores.bind("<Double-1>", seleccionar_totes_les_hores)

# Cursos
f_c = ttk.LabelFrame(f_mid, text=" Grups de la Sortida ", padding=5)
f_c.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
llista_cursos = tk.Listbox(f_c, selectmode=tk.MULTIPLE, height=7, exportselection=False)
scr_c = ttk.Scrollbar(f_c, command=llista_cursos.yview)
llista_cursos.config(yscrollcommand=scr_c.set)
llista_cursos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scr_c.pack(side=tk.RIGHT, fill=tk.Y)

# Botó Analitzar
ttk.Button(app, text="GENERAR LLISTAT D'IDONEÏTAT", command=iniciar_analisi).pack(pady=10, fill=tk.X, padx=15)

# Resultats (Amb expand=True per ocupar la resta)
f_res = ttk.LabelFrame(app, text=" Resultats i Rànquing ", padding=10)
f_res.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
text_resultats = tk.Text(f_res, wrap=tk.WORD, state=tk.DISABLED, bg="#ffffff", font=("Courier", 10), height=15)
text_resultats.pack(fill=tk.BOTH, expand=True)

app.mainloop()
