# -*- coding: utf-8 -*-

# Autor: Josep M. Sardà Caimel
# Copyright: 2025, IES Ramon Turró i Darder
# Llicència: Creative Commons Reconeixement-NoComercial-CompartirIgual 4.0 Internacional (CC BY-NC-SA 4.0)
# https://creativecommons.org/licenses/by-nc-sa/4.0/

"""
Aquest programa està protegit per la llicència CC BY-NC-SA 4.0.
Això significa que sou lliure de:
- Compartir: copiar i redistribuir el material en qualsevol mitjà o format.
- Adaptar: remesclar, transformar i construir sobre el material.

Sota els següents termes:
- Reconeixement: Heu de donar crèdit adequat, proporcionar un enllaç a la llicència i indicar si s'han fet canvis.
- NoComercial: No podeu utilitzar el material amb finalitats comercials.
- CompartirIgual: Si remescleu, transformeu o construïu sobre el material, heu de distribuir les vostres contribucions sota la mateixa llicència que l'original.
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import webbrowser
from tkinter import font

class FileCopierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Copiador i Reanomenador de Fitxers")
        self.root.geometry("775x350")
        self.root.resizable(False, False)
        
        # Variables
        self.origen_path = tk.StringVar()
        self.desti_path = tk.StringVar()
        self.nom_base = tk.StringVar()
        
        # Crear interfície
        self.create_widgets()
    
    def create_widgets(self):
        # Carpeta Origen
        tk.Label(self.root, text="Carpeta Origen:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=10, pady=15, sticky="w"
        )
        tk.Entry(self.root, textvariable=self.origen_path, width=50).grid(
            row=0, column=1, padx=10, pady=15
        )
        tk.Button(self.root, text="Seleccionar", command=self.select_origen).grid(
            row=0, column=2, padx=10, pady=15
        )
        
        # Carpeta Destí
        tk.Label(self.root, text="Carpeta Destí:", font=("Arial", 10, "bold")).grid(
            row=1, column=0, padx=10, pady=15, sticky="w"
        )
        tk.Entry(self.root, textvariable=self.desti_path, width=50).grid(
            row=1, column=1, padx=10, pady=15
        )
        tk.Button(self.root, text="Seleccionar", command=self.select_desti).grid(
            row=1, column=2, padx=10, pady=15
        )
        
        # Nom Base
        tk.Label(self.root, text="Nom Base:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, padx=10, pady=15, sticky="w"
        )
        tk.Entry(self.root, textvariable=self.nom_base, width=50).grid(
            row=2, column=1, padx=10, pady=15
        )
        
        # Botó d'execució
        tk.Button(
            self.root, 
            text="Copiar i Reanomenar", 
            command=self.copy_and_rename,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=10
        ).grid(row=3, column=0, columnspan=3, pady=25)
    
    def select_origen(self):
        self.root.attributes('-topmost', False)
        folder = filedialog.askdirectory(title="Selecciona la carpeta d'origen")
        self.root.attributes('-topmost', True)
        self.root.attributes('-topmost', False)
        if folder:
            self.origen_path.set(folder)
    
    def select_desti(self):
        self.root.attributes('-topmost', False)
        folder = filedialog.askdirectory(title="Selecciona la carpeta de destí")
        self.root.attributes('-topmost', True)
        self.root.attributes('-topmost', False)
        if folder:
            self.desti_path.set(folder)
    
    def copy_and_rename(self):
        origen = self.origen_path.get()
        desti = self.desti_path.get()
        nom = self.nom_base.get()
        
        # Validacions
        if not origen or not desti or not nom:
            messagebox.showerror("Error", "Si us plau, omple tots els camps!")
            return
        
        if not os.path.exists(origen):
            messagebox.showerror("Error", "La carpeta d'origen no existeix!")
            return
        
        if not os.path.exists(desti):
            try:
                os.makedirs(desti)
            except Exception as e:
                messagebox.showerror("Error", f"No s'ha pogut crear la carpeta destí: {e}")
                return
        
        # Obtenir tots els fitxers de l'origen
        try:
            fitxers = [f for f in os.listdir(origen) if os.path.isfile(os.path.join(origen, f))]
            
            if not fitxers:
                messagebox.showwarning("Atenció", "No hi ha fitxers a la carpeta d'origen!")
                return
            
            if len(fitxers) > 999:
                messagebox.showerror("Error", "Massa fitxers! Màxim 999.")
                return
            
            # Copiar i reanomenar fitxers
            for i, fitxer in enumerate(fitxers, start=1):
                origen_file = os.path.join(origen, fitxer)
                extensio = os.path.splitext(fitxer)[1]
                nou_nom = f"{nom}-{i:03d}{extensio}"
                desti_file = os.path.join(desti, nou_nom)
                
                shutil.copy2(origen_file, desti_file)
            
            messagebox.showinfo(
                "Èxit", 
                f"S'han copiat i reanomenat {len(fitxers)} fitxers correctament!"
            )
        
        except Exception as e:
            messagebox.showerror("Error", f"S'ha produït un error: {e}")

def obrir_llicencia(event):
    webbrowser.open_new(r"https://creativecommons.org/licenses/by-nc-sa/4.0/")

# Crear i executar l'aplicació
if __name__ == "__main__":
    root = tk.Tk()
    app = FileCopierApp(root)
    
    # --- Peu de pàgina (Footer) ---
    footer_frame = tk.Frame(root, height=30, background="#f0f0f0")
    footer_frame.grid(row=4, column=0, columnspan=3, sticky="ew")
    
    autor = "Josep M. Sardà Caimel"
    llicencia_text = "CC BY-NC-SA 4.0"
    info_text = f"Autor: {autor}  |  Llicència: {llicencia_text}"
    
    footer_label = tk.Label(footer_frame, text=info_text, fg="blue", cursor="hand2", background="#f0f0f0")
    
    font_subratllada = font.Font(footer_label, footer_label.cget("font"))
    font_subratllada.configure(underline=True)
    footer_label.configure(font=font_subratllada)
    
    footer_label.bind("<Button-1>", obrir_llicencia)
    footer_label.pack(pady=5)

    root.mainloop()