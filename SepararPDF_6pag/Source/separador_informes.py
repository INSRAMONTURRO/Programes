# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------
# Separador d'Informes d'Alumnes
#
# Aquest programa separa un fitxer PDF d'informes (on cada informe
# ocupa exactament 6 pàgines per alumne) en fitxers PDF individuals,
# anomenant cada fitxer segons les dades de l'alumne importades d'un CSV.
#
# Autor: Josep M.
# Versió: 1.0
# Data: 17 de juny de 2026
# Llicència: Creative Commons BY-NC-SA 4.0
# ----------------------------------------------------------------------

import os
import sys
import re
import csv
import threading
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font

# Intentem importar les llibreries de PDF
try:
    from PyPDF2 import PdfReader, PdfWriter
    PDF_LIBRARY_AVAILABLE = True
except ImportError:
    try:
        from pypdf import PdfReader, PdfWriter
        PDF_LIBRARY_AVAILABLE = True
    except ImportError:
        PDF_LIBRARY_AVAILABLE = False


def resource_path(relative_path):
    """ Retorna la ruta absoluta al recurs, funciona per a desenvolupament i per a PyInstaller """
    try:
        # PyInstaller crea una carpeta temporal i guarda la ruta a _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    
    return os.path.join(base_path, relative_path)


def sanitize_filename(name):
    """Neteja els caràcters no vàlids per a fitxers a Windows/Linux."""
    # Caràcters prohibits: \ / : * ? " < > |
    sanitized = re.sub(r'[\\/*?:"<>|]', "", name)
    # Substitueix múltiples espais per un de sol
    sanitized = re.sub(r'\s+', " ", sanitized)
    return sanitized.strip()


class SeparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Separador d'Informes d'Alumnes")
        self.root.geometry("700x840")
        self.root.configure(bg="#f8fafc")  # Slate-50 background (molt net i modern)
        self.root.minsize(650, 600)

        # Variables d'estat
        self.csv_path = ""
        self.pdf_path = ""
        self.out_dir = ""
        self.students = []
        self.pdf_pages = 0
        self.processing = False

        # Configura els estils font globals de Tkinter
        self.font_title = font.Font(family="Helvetica", size=16, weight="bold")
        self.font_subtitle = font.Font(family="Helvetica", size=10)
        self.font_section = font.Font(family="Helvetica", size=11, weight="bold")
        self.font_body = font.Font(family="Helvetica", size=10)
        self.font_body_bold = font.Font(family="Helvetica", size=10, weight="bold")
        self.font_small = font.Font(family="Helvetica", size=9)

        # Inicialització de la interfície
        self.setup_ui()
        
        # Comprovació inicial de dependències
        if not PDF_LIBRARY_AVAILABLE:
            messagebox.showerror(
                "Dependència Faltant", 
                "No s'ha pogut trobar la biblioteca PyPDF2 o pypdf.\n\n"
                "Instal·la-la executant:\npip install PyPDF2"
            )

    def setup_ui(self):
        # 1. Capçalera (Header)
        header_frame = tk.Frame(self.root, bg="#1e293b", height=80)  # Slate-800
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        lbl_title = tk.Label(
            header_frame, 
            text="📄 Separador d'Informes d'Alumnes", 
            font=self.font_title, 
            fg="#ffffff", 
            bg="#1e293b"
        )
        lbl_title.pack(anchor="w", padx=25, pady=(15, 2))

        lbl_subtitle = tk.Label(
            header_frame, 
            text="Divideix un fitxer PDF (configurable en pàgines per informe) en documents individuals utilitzant les dades d'un CSV.", 
            font=self.font_small, 
            fg="#94a3b8",  # Slate-400
            bg="#1e293b"
        )
        lbl_subtitle.pack(anchor="w", padx=25)

        # Contenidor principal amb scroll/padding
        main_container = tk.Frame(self.root, bg="#f8fafc")
        main_container.pack(fill="both", expand=True, padx=25, pady=20)

        # --- SECCIÓ 1: FITXERS D'ENTRADA (CARD) ---
        lbl_sec1 = tk.Label(
            main_container, 
            text="1. Fitxers de configuració", 
            font=self.font_section, 
            fg="#0f172a", 
            bg="#f8fafc"
        )
        lbl_sec1.pack(anchor="w", pady=(0, 5))

        card1_border, card1 = self.create_card(main_container)
        card1_border.pack(fill="x", pady=(0, 15))

        # CSV input row
        row_csv = tk.Frame(card1, bg="#ffffff")
        row_csv.pack(fill="x", padx=15, pady=10)
        
        lbl_csv_title = tk.Label(row_csv, text="Llistat d'alumnes (CSV):", font=self.font_body_bold, fg="#334155", bg="#ffffff")
        lbl_csv_title.pack(anchor="w")
        
        row_csv_input = tk.Frame(row_csv, bg="#ffffff")
        row_csv_input.pack(fill="x", pady=(4, 0))
        
        self.ent_csv = tk.Entry(row_csv_input, font=self.font_body, fg="#64748b", bg="#f1f5f9", relief="flat", bd=1)
        self.ent_csv.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_csv.config(state="readonly")
        
        btn_csv = self.create_styled_button(
            row_csv_input, 
            text="Selecciona CSV...", 
            bg="#2563eb",  # Blue-600
            fg="#ffffff", 
            hover_bg="#1d4ed8", 
            command=self.select_csv_file
        )
        btn_csv.pack(side="right")

        # PDF input row
        row_pdf = tk.Frame(card1, bg="#ffffff")
        row_pdf.pack(fill="x", padx=15, pady=(0, 15))
        
        lbl_pdf_title = tk.Label(row_pdf, text="Fitxer d'informes agrupats (PDF):", font=self.font_body_bold, fg="#334155", bg="#ffffff")
        lbl_pdf_title.pack(anchor="w")
        
        row_pdf_input = tk.Frame(row_pdf, bg="#ffffff")
        row_pdf_input.pack(fill="x", pady=(4, 0))
        
        self.ent_pdf = tk.Entry(row_pdf_input, font=self.font_body, fg="#64748b", bg="#f1f5f9", relief="flat", bd=1)
        self.ent_pdf.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_pdf.config(state="readonly")
        
        btn_pdf = self.create_styled_button(
            row_pdf_input, 
            text="Selecciona PDF...", 
            bg="#2563eb", 
            fg="#ffffff", 
            hover_bg="#1d4ed8", 
            command=self.select_pdf_file
        )
        btn_pdf.pack(side="right")

        # Files input: Nombre de pàgines per informe
        row_pages = tk.Frame(card1, bg="#ffffff")
        row_pages.pack(fill="x", padx=15, pady=(0, 15))
        
        lbl_pages_title = tk.Label(row_pages, text="Nombre de pàgines per informe:", font=self.font_body_bold, fg="#334155", bg="#ffffff")
        lbl_pages_title.pack(side="left", padx=(0, 10))
        
        self.spin_pages = tk.Spinbox(
            row_pages, 
            from_=1, 
            to=100, 
            width=5, 
            font=self.font_body, 
            relief="flat", 
            bd=1, 
            bg="#f1f5f9", 
            fg="#0f172a",
            justify="center",
            command=self.on_pages_changed
        )
        self.spin_pages.bind("<KeyRelease>", lambda e: self.on_pages_changed())
        self.spin_pages.delete(0, "end")
        self.spin_pages.insert(0, "6")
        self.spin_pages.pack(side="left")

        # --- SECCIÓ 2: RESUM I COINCIDÈNCIES (CARD) ---
        lbl_sec2 = tk.Label(
            main_container, 
            text="2. Resum de les dades detectades", 
            font=self.font_section, 
            fg="#0f172a", 
            bg="#f8fafc"
        )
        lbl_sec2.pack(anchor="w", pady=(0, 5))

        card2_border, card2 = self.create_card(main_container)
        card2_border.pack(fill="x", pady=(0, 15))

        self.lbl_summary_csv = tk.Label(
            card2, 
            text="Llistat d'alumnes: Pendent de carregar el fitxer CSV.", 
            font=self.font_body, 
            fg="#64748b", 
            bg="#ffffff"
        )
        self.lbl_summary_csv.pack(anchor="w", padx=15, pady=(12, 4))

        self.lbl_summary_pdf = tk.Label(
            card2, 
            text="Fitxer PDF: Pendent de carregar el fitxer PDF.", 
            font=self.font_body, 
            fg="#64748b", 
            bg="#ffffff"
        )
        self.lbl_summary_pdf.pack(anchor="w", padx=15, pady=4)

        self.lbl_summary_match = tk.Label(
            card2, 
            text="Estat: Si us plau, selecciona els dos fitxers per continuar.", 
            font=self.font_body_bold, 
            fg="#475569", 
            bg="#ffffff"
        )
        self.lbl_summary_match.pack(anchor="w", padx=15, pady=(4, 12))

        # --- SECCIÓ 3: CARPETA DE SORTIDA (CARD) ---
        lbl_sec3 = tk.Label(
            main_container, 
            text="3. Carpeta de sortida", 
            font=self.font_section, 
            fg="#0f172a", 
            bg="#f8fafc"
        )
        lbl_sec3.pack(anchor="w", pady=(0, 5))

        card3_border, card3 = self.create_card(main_container)
        card3_border.pack(fill="x", pady=(0, 20))

        row_out = tk.Frame(card3, bg="#ffffff")
        row_out.pack(fill="x", padx=15, pady=12)
        
        lbl_out_title = tk.Label(row_out, text="Desa els fitxers separats a:", font=self.font_body_bold, fg="#334155", bg="#ffffff")
        lbl_out_title.pack(anchor="w")
        
        row_out_input = tk.Frame(row_out, bg="#ffffff")
        row_out_input.pack(fill="x", pady=(4, 0))
        
        self.ent_out = tk.Entry(row_out_input, font=self.font_body, fg="#64748b", bg="#f1f5f9", relief="flat", bd=1)
        self.ent_out.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_out.config(state="readonly")
        
        self.btn_out = self.create_styled_button(
            row_out_input, 
            text="Canvia carpeta...", 
            bg="#475569",  # Slate-600
            fg="#ffffff", 
            hover_bg="#334155", 
            command=self.select_output_dir
        )
        self.btn_out.pack(side="right")

        # --- ACCIONS I PROGRESS (CARD PRINCIPAL D'EXECUCIÓ) ---
        exec_frame = tk.Frame(main_container, bg="#f8fafc")
        exec_frame.pack(fill="x", pady=(5, 0))

        # Botó d'inici
        self.btn_run = self.create_styled_button(
            exec_frame,
            text="Comença la Divisió dels Informes",
            bg="#059669",  # Emerald-600
            fg="#ffffff",
            hover_bg="#047857",
            command=self.start_processing,
            font=("Helvetica", 11, "bold")
        )
        self.btn_run.pack(fill="x", ipady=6, pady=(0, 10))
        self.btn_run.config(state="disabled", bg="#cbd5e1")  # Desactivat inicialment

        # Progress frame
        self.progress_frame = tk.Frame(exec_frame, bg="#f8fafc")
        self.progress_frame.pack(fill="x")

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill="x", pady=(2, 5))

        self.lbl_status = tk.Label(
            self.progress_frame, 
            text="Estat: Esperant selecció de fitxers.", 
            font=self.font_body, 
            fg="#64748b", 
            bg="#f8fafc"
        )
        self.lbl_status.pack(anchor="w")

        # Botó extra d'èxit per obrir carpeta (ocult al principi)
        self.btn_open_folder = self.create_styled_button(
            self.progress_frame,
            text="📂 Obrir carpeta de sortida",
            bg="#3b82f6",  # Blue-500
            fg="#ffffff",
            hover_bg="#2563eb",
            command=self.open_output_folder
        )
        # Es mostrarà quan el procés hagi finalitzat amb èxit.

        # --- PEU DE PÀGINA (FOOTER - AUTORIA I LLICÈNCIA) ---
        footer_frame = tk.Frame(self.root, height=45, bg="#f1f5f9", bd=1, relief="solid")
        footer_frame.pack(side="bottom", fill="x")
        footer_frame.pack_propagate(False)

        # Dades d'autoria del fitxer AUTORIA_LLICENCIA
        autor = "Josep M Sardà"
        llicencia_text = "CC BY-NC-SA 4.0"
        url_llicencia = "https://creativecommons.org/licenses/by-nc-sa/4.0/"
        info_text = f"Autor: {autor}  |  Llicència: {llicencia_text}"

        # Etiqueta amb enllaç clicable
        footer_label = tk.Label(
            footer_frame,
            text=info_text,
            fg="#2563eb",
            cursor="hand2",
            bg="#f1f5f9",
            font=self.font_body_bold
        )
        # Subratllar el text per indicar que és un enllaç
        font_subratllada = font.Font(footer_label, footer_label.cget("font"))
        font_subratllada.configure(underline=True)
        footer_label.configure(font=font_subratllada)

        # Esdeveniment clic i situació a l'esquerra
        footer_label.bind("<Button-1>", lambda e: webbrowser.open_new(url_llicencia))
        footer_label.pack(side="left", padx=20, pady=10)

        # Intentem carregar i dibuixar la imatge de la llicència CC
        self.load_license_badge(footer_frame, url_llicencia)

    # --- MÈTODES AUXILIARS DE DISSENY ---

    def create_card(self, parent, bg="#ffffff", border_color="#e2e8f0", border_width=1):
        """Crea un marc estil targeta (card) amb una vora molt fina i neta."""
        border_frame = tk.Frame(parent, bg=border_color)
        card = tk.Frame(border_frame, bg=bg)
        card.pack(fill="both", expand=True, padx=border_width, pady=border_width)
        return border_frame, card

    def create_styled_button(self, parent, text, bg, fg, hover_bg, command, font=None):
        """Crea un botó de Tkinter amb estil pla, cursors moderns i efecte hover."""
        if font is None:
            font = self.font_body_bold
            
        btn = tk.Button(
            parent,
            text=text,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            padx=15,
            pady=5,
            font=font,
            cursor="hand2",
            command=command
        )
        
        # Efectes hover
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg) if btn.cget("state") == "normal" else None)
        btn.bind("<Leave>", lambda e: btn.config(bg=bg) if btn.cget("state") == "normal" else None)
        
        return btn

    def load_license_badge(self, parent, url_llicencia):
        """Intenta carregar la imatge by-nc-sa.png i posar-la a la dreta del peu."""
        image_path = resource_path("by-nc-sa.png")
        
        if os.path.exists(image_path):
            try:
                # Carrega imatge nativa PNG
                cc_img = tk.PhotoImage(file=image_path)
                # Escalar la imatge (subsample divideix la mida: original és 403x141)
                # Amb 4,4 passarà a ser aprox 100x35 píxels, ideal per al peu
                cc_img_scaled = cc_img.subsample(4, 4)
                
                # Cal desar la referència perquè Tkinter no l'elimini del garbage collector
                self.cc_img_ref = cc_img_scaled
                
                img_label = tk.Label(parent, image=cc_img_scaled, bg="#f1f5f9", cursor="hand2")
                img_label.bind("<Button-1>", lambda e: webbrowser.open_new(url_llicencia))
                img_label.pack(side="right", padx=20, pady=5)
            except Exception as e:
                # Si falla, s'ignora i es manté només el text
                print(f"No s'ha pogut carregar la imatge de llicència: {e}")

    # --- LÒGICA D'INTERFÍCIE ---

    def select_csv_file(self):
        """Obre el diàleg per triar el fitxer CSV dels alumnes."""
        initial_dir = os.path.dirname(self.pdf_path) if self.pdf_path else os.getcwd()
        file_path = filedialog.askopenfilename(
            title="Selecciona el fitxer CSV d'alumnes",
            filetypes=[("Arxius CSV", "*.csv"), ("Tots els arxius", "*.*")],
            initialdir=initial_dir
        )
        if not file_path:
            return

        self.csv_path = file_path
        self.update_entry_text(self.ent_csv, file_path)

        # Llegeix els alumnes del CSV
        students, error = self.parse_students_csv(file_path)
        if error:
            messagebox.showerror("Error de format CSV", error)
            self.students = []
        else:
            self.students = students

        self.update_summary_labels()
        self.update_default_output_dir()
        self.check_files_and_validate()

    def select_pdf_file(self):
        """Obre el diàleg per triar el fitxer PDF combinat dels informes."""
        initial_dir = os.path.dirname(self.csv_path) if self.csv_path else os.getcwd()
        file_path = filedialog.askopenfilename(
            title="Selecciona el fitxer PDF d'informes",
            filetypes=[("Documents PDF", "*.pdf"), ("Tots els arxius", "*.*")],
            initialdir=initial_dir
        )
        if not file_path:
            return

        self.pdf_path = file_path
        self.update_entry_text(self.ent_pdf, file_path)

        # Calcula el nombre de pàgines del PDF
        try:
            reader = PdfReader(file_path)
            self.pdf_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("Error de lectura PDF", f"No s'ha pogut analitzar el PDF:\n{str(e)}")
            self.pdf_pages = 0

        self.update_summary_labels()
        self.update_default_output_dir()
        self.check_files_and_validate()

    def select_output_dir(self):
        """Obre el diàleg per triar la carpeta de destinació."""
        initial_dir = self.out_dir if self.out_dir else os.getcwd()
        dir_path = filedialog.askdirectory(
            title="Selecciona la carpeta on desar els informes individuals",
            initialdir=initial_dir
        )
        if dir_path:
            self.out_dir = dir_path
            self.update_entry_text(self.ent_out, dir_path)

    def update_default_output_dir(self):
        """Suggereix automàticament una carpeta de sortida basada en els fitxers d'entrada."""
        # Si ja hi ha un PDF triat, proposem una subcarpeta en el mateix directori
        if self.pdf_path and not self.out_dir:
            pdf_dir = os.path.dirname(self.pdf_path)
            pdf_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            # Neteja de nom per a la carpeta de sortida
            clean_folder_name = f"informes_separats_{sanitize_filename(pdf_name)}"
            default_out = os.path.join(pdf_dir, clean_folder_name)
            self.out_dir = default_out
            self.update_entry_text(self.ent_out, default_out)

    def update_entry_text(self, entry_widget, text):
        """Actualitza de manera segura un Entry configurat com a readonly."""
        entry_widget.config(state="normal")
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, text)
        entry_widget.config(state="readonly")

    def get_pages_per_report(self):
        """Obté el nombre de pàgines per informe del Spinbox de forma segura."""
        try:
            val = int(self.spin_pages.get())
            if val < 1:
                return 1
            return val
        except ValueError:
            return 6  # Valor per defecte si no és un número vàlid

    def on_pages_changed(self):
        """S'executa quan l'usuari canvia el nombre de pàgines per informe."""
        self.update_summary_labels()
        self.check_files_and_validate()

    def update_summary_labels(self):
        """Actualitza els textos resum basant-se en els fitxers i el nombre de pàgines."""
        # 1. Resum CSV
        if self.csv_path:
            if not self.students:
                self.lbl_summary_csv.config(text="Llistat d'alumnes: Error en llegir el fitxer o no conté alumnes.", fg="#ef4444")
            else:
                self.lbl_summary_csv.config(
                    text=f"Llistat d'alumnes: {len(self.students)} alumnes detectats correctament.", 
                    fg="#059669"
                )
        else:
            self.lbl_summary_csv.config(
                text="Llistat d'alumnes: Pendent de carregar el fitxer CSV.", 
                fg="#64748b"
            )

        # 2. Resum PDF
        if self.pdf_path:
            if self.pdf_pages > 0:
                pages_per_report = self.get_pages_per_report()
                reports_count = self.pdf_pages // pages_per_report
                self.lbl_summary_pdf.config(
                    text=f"Fitxer PDF: {self.pdf_pages} pàgines ({reports_count} informes de {pages_per_report} pàgines).", 
                    fg="#059669"
                )
            else:
                self.lbl_summary_pdf.config(
                    text="Fitxer PDF: Error en analitzar el fitxer o està buit.", 
                    fg="#ef4444"
                )
        else:
            self.lbl_summary_pdf.config(
                text="Fitxer PDF: Pendent de carregar el fitxer PDF.", 
                fg="#64748b"
            )

    def check_files_and_validate(self):
        """Comprova si tenim tota la informació i valida les coincidències."""
        if not self.csv_path or not self.pdf_path:
            self.lbl_summary_match.config(
                text="Estat: Si us plau, selecciona els dos fitxers per continuar.", 
                fg="#475569"
            )
            self.btn_run.config(state="disabled", bg="#cbd5e1")
            return

        # Ambdós fitxers estan carregats
        num_students = len(self.students)
        pages_per_report = self.get_pages_per_report()
        total_reports = self.pdf_pages // pages_per_report

        if num_students == 0:
            self.lbl_summary_match.config(text="Estat: El fitxer CSV no conté alumnes vàlids.", fg="#ef4444")
            self.btn_run.config(state="disabled", bg="#cbd5e1")
            return
            
        if self.pdf_pages == 0:
            self.lbl_summary_match.config(text="Estat: El fitxer PDF no es pot llegir o està buit.", fg="#ef4444")
            self.btn_run.config(state="disabled", bg="#cbd5e1")
            return

        # Validem si coincideix exactament
        if num_students == total_reports:
            self.lbl_summary_match.config(
                text=f"Estat d'encaix perfecte: {num_students} alumnes i {total_reports} informes de {pages_per_report} pàgines.",
                fg="#059669"  # Emerald-600
            )
        else:
            # Segons requeriment de l'usuari: "Processar només els alumnes que hi ha al CSV en l'ordre en què apareixen"
            # Mostrem un avís informatiu
            self.lbl_summary_match.config(
                text=f"Avís: CSV té {num_students} alumnes, però el PDF té {total_reports} informes.\n"
                     f"Es processaran els primers {min(num_students, total_reports)} en l'ordre del CSV.",
                fg="#b45309"  # Amber-700
            )

        # Activem el botó d'executar
        self.btn_run.config(state="normal", bg="#059669")

    def parse_students_csv(self, file_path):
        """Processa el fitxer CSV intentant diversos encodificats i delimitadors."""
        students = []
        # Encodificats comuns a provar per a evitar fallades amb accents catalans (Júlia, Díaz, etc.)
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
        
        for encoding in encodings:
            try:
                with open(file_path, mode='r', encoding=encoding) as f:
                    content = f.read(2048)
                    f.seek(0)
                    
                    # Detectem delimitador comú
                    delimiter = ','
                    if ';' in content and content.count(';') > content.count(','):
                        delimiter = ';'
                    
                    reader = csv.reader(f, delimiter=delimiter)
                    headers = next(reader)
                    
                    # Normalitzem les capçaleres per a cercar correspondències
                    headers = [h.strip().upper() for h in headers]
                    
                    # Cerquem índexs basats en el nom del camp
                    nom_idx = -1
                    cognom1_idx = -1
                    ralc_idx = -1
                    
                    for idx, h in enumerate(headers):
                        # Fem comparació exacta o inici de capçalera per a evitar col·lisions (ja que COGNOM conté NOM)
                        if h == 'NOM' or h.startswith('NOM '):
                            nom_idx = idx
                        elif h == 'COGNOM1' or h.startswith('COGNOM1'):
                            cognom1_idx = idx
                        elif h == 'RALC' or h.startswith('RALC'):
                            ralc_idx = idx
                    
                    # Si no troba les capçaleres exactes, apliquem valors per defecte si hi ha prou columnes
                    if nom_idx == -1 and len(headers) > 1:
                        nom_idx = 1
                    if cognom1_idx == -1 and len(headers) > 2:
                        cognom1_idx = 2
                    if ralc_idx == -1:
                        # Fallback flexible cercant la cadena RALC dins la capçalera
                        for idx, h in enumerate(headers):
                            if 'RALC' in h:
                                ralc_idx = idx
                                break
                        if ralc_idx == -1:
                            ralc_idx = len(headers) - 1  # Última columna com a últim recurs
                    
                    # Processa les files
                    for row in reader:
                        # Salta files buides
                        if not row or all(cell.strip() == '' for cell in row):
                            continue
                        
                        # Assegura que la fila té prous elements
                        while len(row) <= max(nom_idx, cognom1_idx, ralc_idx, 0):
                            row.append('')
                            
                        nom = row[nom_idx].strip()
                        cognom1 = row[cognom1_idx].strip()
                        ralc = row[ralc_idx].strip()
                        
                        if nom or cognom1 or ralc:
                            students.append({
                                'nom': nom,
                                'cognom1': cognom1,
                                'ralc': ralc
                            })
                    
                    # Si hem llegit correctament i tenim alumnes, ho donem per bo i sortim del bucle d'encodings
                    if students:
                        return students, None
            except Exception as e:
                # Provant el següent encoding en cas d'error
                continue
                
        return [], "No s'ha pogut obrir o llegir el fitxer CSV. Assegura't que el format és vàlid."

    # --- PROCESSAMENT FILTRAT I BACKGROUND THREADING ---

    def start_processing(self):
        """Prepara l'estat i llança el procés en un fil de fons."""
        if not self.csv_path or not self.pdf_path or not self.out_dir:
            messagebox.showwarning("Dades incompletes", "Falten fitxers o ruta de sortida per assignar.")
            return

        if self.processing:
            return

        self.processing = True
        self.btn_run.config(state="disabled", text="Processant els informes...", bg="#cbd5e1")
        self.btn_out.config(state="disabled", bg="#cbd5e1")
        self.btn_open_folder.pack_forget()  # Ocultem si estava visible d'un run anterior
        self.progress_var.set(0)
        self.lbl_status.config(text="Iniciant la separació del fitxer PDF...", fg="#0f172a")

        pages_per_report = self.get_pages_per_report()
        # Llança el procés en un fil independent per evitar congelar la GUI
        t = threading.Thread(target=self.run_split, args=(pages_per_report,))
        t.daemon = True
        t.start()

    def update_status(self, text, progress_val=None):
        """Actualitza de manera segura els missatges d'estat i el progrés des d'altres fils."""
        def _update():
            self.lbl_status.config(text=text)
            if progress_val is not None:
                self.progress_var.set(progress_val)
        self.root.after(0, _update)

    def run_split(self, pages_per_report):
        """Lògica principal de divisió de fitxers PDF executada al fil de fons."""
        csv_path = self.csv_path
        pdf_path = self.pdf_path
        out_dir = self.out_dir
        students = self.students
        
        try:
            # Creem la carpeta de sortida si no existís
            os.makedirs(out_dir, exist_ok=True)
            
            # Carreguem el lector de PDF
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            
            num_students = len(students)
            total_reports = total_pages // pages_per_report
            
            # Calculem quantes files processarem exactament (el menor dels dos valors)
            to_process = min(num_students, total_reports)
            
            for i in range(to_process):
                # Comprovació si l'usuari tanca o cancel·la
                if not self.processing:
                    self.update_status("Procés aturat.", 0)
                    self.root.after(0, self.reset_ui_buttons)
                    return
                
                student = students[i]
                nom = student['nom']
                cognom1 = student['cognom1']
                ralc = student['ralc']
                
                # Sanitzem els camps per crear un nom de fitxer segur
                safe_nom = sanitize_filename(nom)
                safe_cognom = sanitize_filename(cognom1)
                safe_ralc = sanitize_filename(ralc)
                
                # Format de nom de fitxer: AD_Nom_Cognom1_RALC.pdf
                filename = f"AD_{safe_nom}_{safe_cognom}_{safe_ralc}.pdf"
                dest_file_path = os.path.join(out_dir, filename)
                
                # Pàgines d'inici i final (0-indexed, N pàgines per informe)
                start_page = i * pages_per_report
                end_page = start_page + pages_per_report
                
                # Escriptura del PDF d'aquest alumne
                writer = PdfWriter()
                for p_num in range(start_page, end_page):
                    # Afegim la pàgina (mantenint la compatibilitat amb PyPDF2 i pypdf)
                    writer.add_page(reader.pages[p_num])
                    
                with open(dest_file_path, "wb") as f_out:
                    writer.write(f_out)
                
                # Actualitzem el progrés de forma interactiva
                percent = int(((i + 1) / to_process) * 100)
                status_text = f"Generat ({i+1}/{to_process}): {nom} {cognom1}"
                self.update_status(status_text, percent)
                
            # Finalitzat correctament
            self.update_status(f"Completat! S'han generat {to_process} fitxers correctament a la destinació.", 100)
            self.root.after(0, lambda: self.finish_processing_ui(to_process))
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            self.root.after(0, lambda: messagebox.showerror("Error de processament", f"Hi ha hagut un error separant els PDF:\n{str(e)}"))
            self.root.after(0, self.reset_ui_buttons)

    def finish_processing_ui(self, count):
        """Actualitza la interfície per indicar finalització i mostra botó d'obertura."""
        self.processing = False
        self.btn_run.config(state="normal", text="Comença la Divisió dels Informes", bg="#059669")
        self.btn_out.config(state="normal", bg="#475569")
        
        # Mostrem el botó dinàmic per obrir la carpeta de sortida
        self.btn_open_folder.pack(pady=(10, 0), fill="x")
        
        messagebox.showinfo(
            "Procés Completat", 
            f"El procés ha finalitzat correctament.\n\nS'han generat {count} fitxers PDF individuals a:\n{self.out_dir}"
        )

    def reset_ui_buttons(self):
        """Restaura els botons a l'estat inicial en cas d'error o aturada."""
        self.processing = False
        self.btn_run.config(state="normal", text="Comença la Divisió dels Informes", bg="#059669")
        self.btn_out.config(state="normal", bg="#475569")

    def open_output_folder(self):
        """Obre el directori de destinació amb l'explorador de fitxers del sistema (Linux/Unix/Windows)."""
        if not self.out_dir or not os.path.exists(self.out_dir):
            messagebox.showwarning("Carpeta no trobada", "La carpeta de sortida encara no ha estat creada o no existeix.")
            return
            
        try:
            if sys.platform == 'win32':
                os.startfile(self.out_dir)
            elif sys.platform == 'darwin':
                subprocess.run(["open", self.out_dir])
            else:  # Linux o altres sistemes Unix
                subprocess.run(["xdg-open", self.out_dir])
        except Exception as e:
            messagebox.showerror("Error en obrir carpeta", f"No s'ha pogut obrir l'explorador de fitxers:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SeparadorApp(root)
    root.mainloop()
