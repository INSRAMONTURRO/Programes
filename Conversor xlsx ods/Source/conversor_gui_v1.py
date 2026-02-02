# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------
# Conversor de Fitxers (ODS <> XLSX)
#
# Una aplicació d'escriptori per convertir fitxers entre formats ODS i XLSX
# utilitzant LibreOffice en segon pla.
#
# Autor: Josep Maria
# Versió: 1.0
# Data: 02 de Febrer de 2026
# Llicència: Creative Commons BY-NC-SA 4.0
#
# Canvis:
# - v1.0: Versió inicial.
# ----------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import os
import subprocess
import threading
import webbrowser

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Fitxers (ODS <> XLSX) v1.0")
        self.root.geometry("800x750")
        self.root.configure(bg="#f0f0f0")

        # --- Variables ---
        self.source_folder_path = tk.StringVar()
        self.destination_folder_path = tk.StringVar()
        self.conversion_mode = tk.StringVar(value="ods_a_xlsx")

        # --- Estil ---
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0")
        style.configure("TRadiobutton", background="#f0f0f0")
        style.configure("TLabelframe", background="#f0f0f0", bordercolor="#cccccc")
        style.configure("TLabelframe.Label", background="#f0f0f0")

        # --- Widgets ---
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Selecció de carpeta d'origen
        source_folder_frame = ttk.LabelFrame(main_frame, text="1. Tria la carpeta d'origen", padding="10")
        source_folder_frame.pack(fill=tk.X, pady=10)

        source_folder_entry = ttk.Entry(source_folder_frame, textvariable=self.source_folder_path, state="readonly", width=70)
        source_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        source_browse_button = ttk.Button(source_folder_frame, text="Navega...", command=self.select_source_folder)
        source_browse_button.pack(side=tk.LEFT)

        # Selecció de carpeta de destí
        dest_folder_frame = ttk.LabelFrame(main_frame, text="2. Tria la carpeta de destí", padding="10")
        dest_folder_frame.pack(fill=tk.X, pady=10)

        dest_folder_entry = ttk.Entry(dest_folder_frame, textvariable=self.destination_folder_path, state="readonly", width=70)
        dest_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        dest_browse_button = ttk.Button(dest_folder_frame, text="Navega...", command=self.select_destination_folder)
        dest_browse_button.pack(side=tk.LEFT)

        # Mode de conversió
        mode_frame = ttk.LabelFrame(main_frame, text="3. Tria la conversió", padding="10")
        mode_frame.pack(fill=tk.X, pady=10, anchor=tk.W)

        ttk.Radiobutton(mode_frame, text="De ODS a XLSX", variable=self.conversion_mode, value="ods_a_xlsx").pack(anchor=tk.W, padx=5)
        ttk.Radiobutton(mode_frame, text="De XLSX a ODS", variable=self.conversion_mode, value="xlsx_a_ods").pack(anchor=tk.W, padx=5)

        # Botó d'acció
        self.convert_button = ttk.Button(main_frame, text="4. Converteix!", command=self.start_conversion_thread, style="Accent.TButton")
        style.configure("Accent.TButton", font=("", 10, "bold"))
        self.convert_button.pack(pady=15, fill=tk.X, ipady=5)

        # Log de sortida
        log_frame = ttk.LabelFrame(main_frame, text="Progrés", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(log_frame, height=10, state="disabled", bg="#e8e8e8", relief=tk.SUNKEN, borderwidth=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- Peu de pàgina (Footer) ---
        self.setup_footer()

    def setup_footer(self):
        """Configura el peu de pàgina amb informació d'autoria i llicència."""
        footer_frame = tk.Frame(self.root, height=40, bg="#dcdcdc")
        footer_frame.pack(side="bottom", fill="x")

        autor = "Josep Maria"
        llicencia_text = "CC BY-NC-SA 4.0"
        url_llicencia = "https://creativecommons.org/licenses/by-nc-sa/4.0/"

        info_text = f"Autor: {autor}  |  Llicència: {llicencia_text}"

        footer_label = tk.Label(
            footer_frame,
            text=info_text,
            fg="#00008B",
            cursor="hand2",
            bg="#dcdcdc"
        )

        font_subratllada = font.Font(footer_label, footer_label.cget("font"))
        font_subratllada.configure(underline=True)
        footer_label.configure(font=font_subratllada)

        footer_label.bind("<Button-1>", lambda e: webbrowser.open_new(url_llicencia))
        footer_label.pack(pady=5)

    def log_message(self, message):
        """Insereix missatges al quadre de text de forma segura des de fils."""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def select_source_folder(self):
        """Obre el diàleg per seleccionar la carpeta d'origen."""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.source_folder_path.set(folder_selected)
            self.log_message(f"Carpeta d'origen: {folder_selected}")

    def select_destination_folder(self):
        """Obre el diàleg per seleccionar la carpeta de destí."""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.destination_folder_path.set(folder_selected)
            self.log_message(f"Carpeta de destí: {folder_selected}")

    def start_conversion_thread(self):
        """Inicia la conversió en un fil separat per no bloquejar la GUI."""
        if not self.source_folder_path.get():
            messagebox.showerror("Error", "Si us plau, tria una carpeta d'origen primer.")
            return
        if not self.destination_folder_path.get():
            messagebox.showerror("Error", "Si us plau, tria una carpeta de destí.")
            return

        self.convert_button.config(state="disabled")
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()

    def convert_files(self):
        """Lògica principal de la conversió de fitxers."""
        source_folder = self.source_folder_path.get()
        destination_folder = self.destination_folder_path.get()
        mode = self.conversion_mode.get()

        if mode == "ods_a_xlsx":
            source_ext = ".ods"
            target_format = "xlsx"
            self.log_message("Iniciant conversió de ODS a XLSX...")
        else:
            source_ext = ".xlsx"
            target_format = "ods"
            self.log_message("Iniciant conversió de XLSX a ODS...")

        files_to_convert = [f for f in os.listdir(source_folder) if f.lower().endswith(source_ext)]

        if not files_to_convert:
            self.log_message(f"No s'han trobat fitxers amb l'extensió '{source_ext}' a la carpeta d'origen.")
            self.root.after(0, lambda: self.convert_button.config(state="normal"))
            return

        self.log_message(f"S'han trobat {len(files_to_convert)} fitxers per convertir.")
        
        for filename in files_to_convert:
            file_path = os.path.join(source_folder, filename)
            self.log_message(f"Convertint: {filename}...")
            try:
                command = [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    target_format,
                    "--outdir",
                    destination_folder,
                    file_path
                ]
                process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate()

                if process.returncode != 0:
                    self.log_message(f"ERROR en convertir {filename}:")
                    self.log_message(stderr.decode('utf-8', errors='ignore'))
                else:
                    self.log_message(f"-> {filename} convertit correctament.")

            except FileNotFoundError:
                self.log_message("\nERROR: No es troba l'executable 'libreoffice'.")
                self.log_message("Assegura't que LibreOffice estigui instal·lat i accessible al PATH del sistema.")
                self.root.after(0, lambda: messagebox.showerror("Error d'Execució", "No es troba l'executable 'libreoffice'. Assegura't que estigui instal·lat."))
                self.root.after(0, lambda: self.convert_button.config(state="normal"))
                return
            except Exception as e:
                self.log_message(f"Ha ocorregut un error inesperat: {e}")

        self.log_message("\n--- Conversió completada! ---")
        self.root.after(0, lambda: self.convert_button.config(state="normal"))
        self.root.after(0, lambda: messagebox.showinfo("Finalitzat", "La conversió de tots els fitxers ha finalitzat."))


if __name__ == "__main__":
    try:
        from ttkthemes import ThemedTk
        root = ThemedTk(theme="arc")
    except (ImportError, tk.TclError):
        root = tk.Tk()
        
    app = FileConverterApp(root)
    root.mainloop()
