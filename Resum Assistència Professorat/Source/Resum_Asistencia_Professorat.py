# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------
# Analitzador d'Assistència Complet
#
# Unifica l'anàlisi de retards i faltes de fitxatge de professors a partir
# de fitxers Excel, generant informes resumits (Excel) i detallats (Markdown).
#
# Autor: Josep M. Sardà Caimel
# Versió: 1.3
# Data: 01 de Febrer de 2026
# Llicència: Creative Commons Reconeixament-NoComercial-CompartirIgual 4.0 Internacional (CC BY-NC-SA 4.0)
#
# Canvis:
# - v1.3: Afegida opció per desar informes en una carpeta específica.
# - v1.2: Corregit error de 'if_sheet_exist' en escriure Excel, gestionant ExcelWriter centralitzat.
# - v1.1: Corregit error de lògica en filtres de retards, advertència de Pandas i layout de la GUI.
# - v1.0: Versió inicial unificada.
# ----------------------------------------------------------------------

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import os
from datetime import datetime
import re
import webbrowser


# --- Funcions d'Anàlisi per a Resums (Excel) ---

def generar_resum_retards_excel(df_complet, minuts_limit, writer):
    """
    Analitza el DataFrame per identificar professors amb diferències de minuts
    inferiors al límit i desa el resultat en una pestanya de l'ExcelWriter.
    """
    if 'Professor' not in df_complet.columns or 'Minuts Diferència' not in df_complet.columns:
        messagebox.showerror("Error de Dades", "Les columnes 'Professor' i/o 'Minuts Diferència' no existeixen.")
        return False

    df_menys_de_lim = df_complet[df_complet['Minuts Diferència'] < minuts_limit].copy()
    df_menys_de_lim['Minuts Diferència'] = pd.to_numeric(df_menys_de_lim['Minuts Diferència'])

    resultat = df_menys_de_lim.groupby('Professor').agg(
        Cops_Menys_Limit_Minuts=('Professor', 'count'),
        Suma_Total_Minuts=('Minuts Diferència', 'sum')
    ).round(2)

    try:
        resultat.to_excel(writer, sheet_name='Resum_Retards')
        return True
    except Exception as e:
        messagebox.showerror("Error d'Escriptura", f"S'ha produït un error en desar el resum de retards a l'Excel: {e}")
        return False


def generar_resum_fitxatges_excel(df_complet, writer):
    """
    Analitza el DataFrame per identificar faltes de fitxatge i desa el recompte
    en una pestanya de l'ExcelWriter.
    """
    if 'Professor' not in df_complet.columns or 'Tipus Incident' not in df_complet.columns:
        messagebox.showerror("Error de Dades", "Les columnes 'Professor' i/o 'Tipus Incident' no existeixen.")
        return False

    df_entrada = df_complet[df_complet['Tipus Incident'].str.contains('Entrada no fitxada|Entrada no|No entrada', case=False, na=False)]
    df_sortida = df_complet[df_complet['Tipus Incident'].str.contains('Sortida no fitxada|Sortida no|No sortida', case=False, na=False)]

    entrada_counts = df_entrada.groupby('Professor').size().to_dict() if not df_entrada.empty else {}
    sortida_counts = df_sortida.groupby('Professor').size().to_dict() if not df_sortida.empty else {}

    all_professors = set(list(entrada_counts.keys()) + list(sortida_counts.keys()))
    if not all_professors:
        df_resultat = pd.DataFrame(columns=['Cops_Entrada_No_Fitxada', 'Cops_Sortida_No_Fitxada'])
        df_resultat.index.name = 'Professor'
    else:
        df_resultat = pd.DataFrame(index=list(all_professors))
        df_resultat.index.name = 'Professor'
        df_resultat['Cops_Entrada_No_Fitxada'] = df_resultat.index.map(entrada_counts).fillna(0).astype(int)
        df_resultat['Cops_Sortida_No_Fitxada'] = df_resultat.index.map(sortida_counts).fillna(0).astype(int)

    try:
        df_resultat.to_excel(writer, sheet_name='Resum_Faltes_Fitxatge')
        return True
    except Exception as e:
        messagebox.showerror("Error d'Escriptura", f"S'ha produït un error en desar el resum de fitxatges a l'Excel: {e}")
        return False

# --- Funcions d'Anàlisi per a Informes Detallats (Markdown) ---

def generar_informe_detallat_retards(df_complet, fulls_seleccionats, minuts_acumulats_limit, minuts_max_retard, nom_arxiu_sortida):
    """
    Genera un informe detallat en Markdown sobre els retards dels professors
    que superen un llindar de minuts acumulats.
    """
    df_complet['Minuts Diferència'] = pd.to_numeric(df_complet['Minuts Diferència'], errors='coerce')
    df_complet['Data'] = pd.to_datetime(df_complet['Data'], errors='coerce')

    retards = df_complet[(df_complet['Minuts Diferència'] > 0) & (df_complet['Minuts Diferència'] <= minuts_max_retard)].copy()
    
    retards_per_professor = retards.groupby('Professor')['Minuts Diferència'].sum()
    professors_filtrats = retards_per_professor[retards_per_professor > minuts_acumulats_limit]
    
    try:
        with open(nom_arxiu_sortida, 'w', encoding='utf-8') as f:
            f.write("# Informe Detallat de Retards\n\n")
            f.write(f"**Pestanyes analitzades:** {', '.join(fulls_seleccionats)}\n")
            if not df_complet.empty and not df_complet['Data'].isnull().all():
                f.write(f"**Període:** De {df_complet['Data'].min().strftime('%d-%m-%Y')} a {df_complet['Data'].max().strftime('%d-%m-%Y')}\n\n")
            
            if professors_filtrats.empty:
                f.write(f"No s'han trobat professors amb retards acumulats superiors a {minuts_acumulats_limit} minuts.\n")
            else:
                for professor, total_minuts in professors_filtrats.items():
                    f.write(f"## Professor: {professor}\n")
                    f.write(f"**Total de minuts de retard acumulats:** {total_minuts:.0f}\n\n")
                    f.write("### Detalls dels retards:\n")
                    detalls_professor = retards[retards['Professor'] == professor]
                    for _, row in detalls_professor.iterrows():
                        f.write(f"- **Data:** {row['Data'].strftime('%d-%m-%Y')}, **Minuts:** {row['Minuts Diferència']:.0f}\n")
                    f.write("\n")
        return True
    except Exception as e:
        messagebox.showerror("Error d'Escriptura", f"S'ha produït un error en generar l'informe detallat de retards: {e}")
        return False

def generar_informe_detallat_fitxatges(df_complet, fulls_seleccionats, incidents_limit, nom_arxiu_sortida):
    """
    Genera un informe detallat en Markdown sobre les incidències de fitxatge
    dels professors que superen un llindar d'incidències.
    """
    df_complet['Data'] = pd.to_datetime(df_complet['Data'], errors='coerce')

    incidents_interes = ['Entrada no fitxada', 'Sortida no fitxada', 'Entrada no', 'No entrada', 'Sortida no', 'No sortida']
    filtre_incidents = df_complet['Tipus Incident'].str.contains('|'.join(incidents_interes), case=False, na=False)
    fitxatges_erronis = df_complet[filtre_incidents].copy()
    
    comptador_incidents = fitxatges_erronis.groupby('Professor').size()
    professors_filtrats = comptador_incidents[comptador_incidents > incidents_limit]
    
    try:
        with open(nom_arxiu_sortida, 'w', encoding='utf-8') as f:
            f.write("# Informe Detallat d'Incidències de Fitxatge\n\n")
            f.write(f"**Pestanyes analitzades:** {', '.join(fulls_seleccionats)}\n")
            if not df_complet.empty and not df_complet['Data'].isnull().all():
                f.write(f"**Període:** De {df_complet['Data'].min().strftime('%d-%m-%Y')} a {df_complet['Data'].max().strftime('%d-%m-%Y')}\n\n")

            if professors_filtrats.empty:
                f.write(f"No s'han trobat professors amb més de {incidents_limit} incidències de fitxatge.\n")
            else:
                for professor, num_incidents in professors_filtrats.items():
                    f.write(f"## Professor: {professor}\n")
                    f.write(f"**Total d'incidències:** {num_incidents}\n\n")
                    f.write("### Detalls de les incidències:\n")
                    detalls = fitxatges_erronis[fitxatges_erronis['Professor'] == professor]
                    for _, row in detalls.iterrows():
                        f.write(f"- **Data:** {row['Data'].strftime('%d-%m-%Y')}, **Incident:** {row['Tipus Incident']}\n")
                    f.write("\n")
        return True
    except Exception as e:
        messagebox.showerror("Error d'Escriptura", f"S'ha produït un error en generar l'informe detallat de fitxatges: {e}")
        return False


# --- Interfície Gràfica Principal ---

class App:
    def __init__(self):
        self.root = None
        self.ruta_arxiu_excel = None
        self.pestanyes_tractament = []
        self.var_checkboxes = []
        self.nom_base = ""

        # Variables per a les opcions de la GUI
        self.var_minuts_resum = None
        self.var_minuts_detall = None
        self.var_incidents_detall = None
        self.var_minuts_max_retard = None
        self.var_directori_sortida = None # NOU: Nom de la carpeta
        self.var_gen_resum_excel = None
        self.var_gen_detall_md = None
        self.entry_nom_sortida_excel = None
        self.entry_nom_sortida_md_retards = None
        self.entry_nom_sortida_md_fitxatges = None

    def obrir_llicencia(self, event):
        webbrowser.open_new(r"https://creativecommons.org/licenses/by-nc-sa/4.0/")

    def iniciar_aplicacio(self):
        self.root = tk.Tk()
        self.root.title("Analitzador d'Assistència Complet")
        self.root.geometry("850x900")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1) 
        self.root.rowconfigure(1, weight=0) 

        # Inicialitzar variables
        self.var_minuts_resum = tk.IntVar(value=15)
        self.var_minuts_detall = tk.IntVar(value=20)
        self.var_incidents_detall = tk.IntVar(value=10)
        self.var_minuts_max_retard = tk.IntVar(value=25)
        self.var_directori_sortida = tk.StringVar(value="Informes_Generats") # NOU
        self.var_gen_resum_excel = tk.BooleanVar(value=True)
        self.var_gen_detall_md = tk.BooleanVar(value=True)

        # Estils
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabelframe", background="#f0f0f0", foreground="#333333")
        style.configure("TLabel", background="#f0f0f0", foreground="#333333")
        style.configure("TCheckbutton", background="#f0f0f0", foreground="#333333")

        # --- Frame Principal ---
        frame_principal = ttk.Frame(self.root, padding="20")
        frame_principal.grid(row=0, column=0, sticky="nsew")
        frame_principal.columnconfigure(0, weight=1)
        frame_principal.rowconfigure(4, weight=1) 

        # Disposició dels frames
        self._crear_frame_seleccio(frame_principal).grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self._crear_frame_config_analisi(frame_principal).grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self._crear_frame_destinacio(frame_principal).grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self._crear_frame_tipus_informes(frame_principal).grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self._crear_frame_pestanyes(frame_principal).grid(row=4, column=0, sticky="nsew", pady=(0, 10))
        self._crear_frame_botons(frame_principal).grid(row=5, column=0, sticky="ew", pady=(20, 0))

        self._crear_footer(self.root)
        self.root.mainloop()

    def _crear_frame_seleccio(self, parent):
        frame = ttk.LabelFrame(parent, text="1. Fitxer d'Entrada", padding="10")
        ttk.Button(frame, text="Selecciona Fitxer Excel", command=self.seleccionar_fitxer).pack(pady=5)
        self.label_arxiu_seleccionat = ttk.Label(frame, text="Cap fitxer seleccionat", foreground="gray", wraplength=400)
        self.label_arxiu_seleccionat.pack(pady=5)
        return frame

    def _crear_frame_config_analisi(self, parent):
        frame = ttk.LabelFrame(parent, text="2. Configuració d'Anàlisi", padding="10")
        frame.columnconfigure(1, weight=1)
        
        ttk.Label(frame, text="Llindar resum retards (minuts):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=self.var_minuts_resum, width=10).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(frame, text="Llindar detall retards (minuts acumulats):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=self.var_minuts_detall, width=10).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(frame, text="Límit superior per retard individual (detall):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=self.var_minuts_max_retard, width=10).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(frame, text="Llindar detall fitxatges (núm. incidents):").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=self.var_incidents_detall, width=10).grid(row=3, column=1, sticky="w", padx=5)
        return frame

    def _crear_frame_destinacio(self, parent):
        frame = ttk.LabelFrame(parent, text="3. Destinació de Sortida", padding="10")
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Nom de la carpeta:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=self.var_directori_sortida, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        return frame

    def _crear_frame_tipus_informes(self, parent):
        frame = ttk.LabelFrame(parent, text="4. Tipus i Noms dels Informes", padding="10")
        frame.columnconfigure(1, weight=1)
        
        ttk.Checkbutton(frame, text="Generar Resum en Excel", variable=self.var_gen_resum_excel).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nom_sortida_excel = ttk.Entry(frame, width=50)
        self.entry_nom_sortida_excel.grid(row=0, column=1, sticky="ew", padx=5)
        
        ttk.Checkbutton(frame, text="Generar Informes Detallats (Markdown)", variable=self.var_gen_detall_md).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        frame_md_noms = ttk.Frame(frame)
        frame_md_noms.grid(row=1, column=1, sticky="ew")
        frame_md_noms.columnconfigure(1, weight=1)
        ttk.Label(frame_md_noms, text="Retards:").grid(row=0, column=0, sticky="w", padx=(5,2))
        self.entry_nom_sortida_md_retards = ttk.Entry(frame_md_noms, width=22)
        self.entry_nom_sortida_md_retards.grid(row=0, column=1, sticky="ew")
        ttk.Label(frame_md_noms, text="Fitxatges:").grid(row=1, column=0, sticky="w", padx=(5,2))
        self.entry_nom_sortida_md_fitxatges = ttk.Entry(frame_md_noms, width=22)
        self.entry_nom_sortida_md_fitxatges.grid(row=1, column=1, sticky="ew")
        return frame

    def _crear_frame_pestanyes(self, parent):
        frame = ttk.LabelFrame(parent, text="5. Selecció de Pestanyes", padding="10")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        frame_lista_pestanyes = ttk.Frame(frame)
        frame_lista_pestanyes.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(frame_lista_pestanyes, background="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_lista_pestanyes, orient="vertical", command=canvas.yview)
        self.frame_checkboxes = ttk.Frame(canvas)

        self.frame_checkboxes.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.frame_checkboxes, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return frame

    def _crear_frame_botons(self, parent):
        frame = ttk.Frame(parent, style="TFrame")
        self.boto_processar = tk.Button(frame, text="Processar Anàlisi", command=self.processar_analisi,
                                       state=tk.DISABLED, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.boto_processar.pack(side=tk.LEFT, padx=10, ipady=5, ipadx=10)
        ttk.Button(frame, text="Tancar", command=self.root.destroy).pack(side=tk.LEFT, padx=10)
        return frame

    def _crear_footer(self, parent):
        footer_frame = tk.Frame(parent, height=30, bg="#e0e0e0")
        footer_frame.grid(row=1, column=0, sticky="ew")
        
        autor = "Josep M. Sardà Caimel"
        llicencia_text = "CC BY-NC-SA 4.0"
        info_text = f"Autor: {autor}  |  Llicència: {llicencia_text}"

        footer_label = tk.Label(footer_frame, text=info_text, fg="blue", cursor="hand2", bg="#e0e0e0")
        font_subratllada = font.Font(footer_label, footer_label.cget("font"))
        font_subratllada.configure(underline=True)
        footer_label.configure(font=font_subratllada)
        footer_label.bind("<Button-1>", self.obrir_llicencia)
        footer_label.pack(pady=5)

    def seleccionar_fitxer(self):
        fitxer_excel = filedialog.askopenfilename(title="Selecciona un fitxer Excel", filetypes=[("Fitxers Excel", "*.xlsx *.xls")])
        if not fitxer_excel: return

        self.ruta_arxiu_excel = fitxer_excel
        self.nom_base = os.path.splitext(os.path.basename(fitxer_excel))[0]
        self.label_arxiu_seleccionat.config(text=os.path.basename(fitxer_excel))

        try:
            xls = pd.ExcelFile(fitxer_excel)
            self.pestanyes_tractament = [nom for nom in xls.sheet_names if nom.startswith('Tractament')]
            self._actualitzar_llista_pestanyes()
            self._actualitzar_noms_sortida()
            self.boto_processar.config(state=tk.NORMAL if self.pestanyes_tractament else tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"S'ha produït un error en llegir el fitxer: {e}")
            self.boto_processar.config(state=tk.DISABLED)

    def _actualitzar_llista_pestanyes(self):
        for widget in self.frame_checkboxes.winfo_children():
            widget.destroy()
        self.var_checkboxes = []
        
        if not self.pestanyes_tractament:
            ttk.Label(self.frame_checkboxes, text="No s'han trobat pestanyes que comencin amb 'Tractament'.").pack()
            return
            
        for i, nom_pestanya in enumerate(self.pestanyes_tractament):
            var = tk.BooleanVar(value=False)
            chk = ttk.Checkbutton(self.frame_checkboxes, text=nom_pestanya, variable=var)
            chk.grid(row=i // 2, column=i % 2, sticky="w", padx=10, pady=2)
            self.var_checkboxes.append(var)

    def _actualitzar_noms_sortida(self):
        for entry in [self.entry_nom_sortida_excel, self.entry_nom_sortida_md_retards, self.entry_nom_sortida_md_fitxatges]:
            entry.delete(0, tk.END)
        self.entry_nom_sortida_excel.insert(0, f"{self.nom_base}_resum.xlsx")
        self.entry_nom_sortida_md_retards.insert(0, f"{self.nom_base}_detall_retards.md")
        self.entry_nom_sortida_md_fitxatges.insert(0, f"{self.nom_base}_detall_fitxatges.md")
    
    def processar_analisi(self):
        pestanyes_seleccionades = [self.pestanyes_tractament[i] for i, var in enumerate(self.var_checkboxes) if var.get()]
        if not pestanyes_seleccionades:
            messagebox.showwarning("Atenció", "Has de seleccionar almenys una pestanya.")
            return
            
        generar_excel = self.var_gen_resum_excel.get()
        generar_md = self.var_gen_detall_md.get()
        if not generar_excel and not generar_md:
            messagebox.showwarning("Atenció", "Has de seleccionar almenys un tipus d'informe per generar.")
            return

        # NOU: Crear directori de sortida
        directori_sortida = self.var_directori_sortida.get().strip()
        if directori_sortida:
            try:
                os.makedirs(directori_sortida, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Error de Carpeta", f"No s'ha pogut crear la carpeta '{directori_sortida}': {e}")
                return
        
        try:
            dfs = [pd.read_excel(self.ruta_arxiu_excel, sheet_name=s) for s in pestanyes_seleccionades]
            dfs_no_buits = [df for df in dfs if not df.empty]
            
            if not dfs_no_buits:
                messagebox.showwarning("Sense Dades", "Les pestanyes seleccionades estan buides o no contenen dades.")
                return
            
            df_complet = pd.concat(dfs_no_buits, ignore_index=True)
            
            columnes_necessaries = {'Professor', 'Minuts Diferència', 'Tipus Incident', 'Data'}
            if not columnes_necessaries.issubset(df_complet.columns):
                messagebox.showerror("Error", f"Falten columnes. Assegura't que existeixen: {', '.join(columnes_necessaries - set(df_complet.columns))}")
                return

            exit_generat = False
            missatges_exit = []

            # Generar resum Excel
            if generar_excel:
                nom_fitxer_excel = self.entry_nom_sortida_excel.get()
                ruta_sortida_excel = os.path.join(directori_sortida, nom_fitxer_excel) if directori_sortida else nom_fitxer_excel
                
                try:
                    with pd.ExcelWriter(ruta_sortida_excel, engine='openpyxl') as writer:
                        res_retards = generar_resum_retards_excel(df_complet.copy(), self.var_minuts_resum.get(), writer)
                        res_fitxatges = generar_resum_fitxatges_excel(df_complet.copy(), writer)
                        if res_retards and res_fitxatges:
                            exit_generat = True
                            missatges_exit.append(f"Resum Excel '{ruta_sortida_excel}'")
                except Exception as e:
                    messagebox.showerror("Error d'Escriptura", f"S'ha produït un error en crear o desar l'Excel de resum: {e}")

            # Generar informes Markdown
            if generar_md:
                nom_md_retards = self.entry_nom_sortida_md_retards.get()
                ruta_md_retards = os.path.join(directori_sortida, nom_md_retards) if directori_sortida else nom_md_retards
                if generar_informe_detallat_retards(df_complet.copy(), pestanyes_seleccionades, self.var_minuts_detall.get(), self.var_minuts_max_retard.get(), ruta_md_retards):
                    exit_generat = True
                    missatges_exit.append(f"Informe de retards '{ruta_md_retards}'")

                nom_md_fitxatges = self.entry_nom_sortida_md_fitxatges.get()
                ruta_md_fitxatges = os.path.join(directori_sortida, nom_md_fitxatges) if directori_sortida else nom_md_fitxatges
                if generar_informe_detallat_fitxatges(df_complet.copy(), pestanyes_seleccionades, self.var_incidents_detall.get(), ruta_md_fitxatges):
                    exit_generat = True
                    missatges_exit.append(f"Informe de fitxatges '{ruta_md_fitxatges}'")

            if exit_generat:
                messagebox.showinfo("Èxit", "S'han generat correctament els següents informes:\n\n- " + "\n- ".join(missatges_exit))
            else:
                messagebox.showwarning("Sense Resultats", "L'anàlisi s'ha completat, però no s'ha generat cap informe. Revisa la configuració i les dades.")

        except Exception as e:
            messagebox.showerror("Error Greu", f"S'ha produït un error inesperat durant l'anàlisi: {e}")

def iniciar_aplicacio():
    app = App()
    app.iniciar_aplicacio()

if __name__ == "__main__":
    iniciar_aplicacio()
