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
from tkinter import filedialog, messagebox, ttk, colorchooser
from tkinter import font
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PIL import Image
import io
import os
import webbrowser

class PDFEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Editor PDF - Afegir Capçalera")
        self.root.geometry("650x700")
        self.pdf_entrada = ""
        self.imatge_path = ""
        self.pdf_sortida = ""
        self.fons_color = "#FFFFFF"
        self.crear_interficie()

    def crear_interficie(self):
        title_font = font.Font(family="Arial", size=16, weight="bold")
        tk.Label(self.root, text="Afegir Capçalera a PDF", font=title_font).pack(pady=20)

        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=20, pady=10, fill='both', expand=True)

        # PDF i imatge
        self.afegir_selector_fitxers(main_frame)

        # Text i posicions
        self.afegir_configuracio_capcalera(main_frame)

        # Rectangle de fons
        self.afegir_opcio_rectangle_fons(main_frame)

        # Botó
        tk.Button(main_frame, text="Processar PDF", command=self.processar_pdf,
                  bg="#FF9800", fg="white", font=("Arial", 12, "bold"),
                  width=20, height=2).pack(pady=10)

        # Progrés i log
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.pack(pady=10)
        self.log_text = tk.Text(main_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill='both', expand=True, pady=10)

    def afegir_selector_fitxers(self, frame):
        pdf_frame = tk.Frame(frame)
        pdf_frame.pack(fill='x', pady=10)
        tk.Label(pdf_frame, text="PDF d'entrada:", font=("Arial", 10, "bold")).pack(anchor='w')
        self.pdf_label = tk.Label(pdf_frame, text="Cap fitxer seleccionat", fg="gray")
        self.pdf_label.pack(anchor='w', pady=5)
        tk.Button(pdf_frame, text="Seleccionar PDF", command=self.seleccionar_pdf,
                  bg="#4CAF50", fg="white", font=("Arial", 10)).pack(anchor='w')

        img_frame = tk.Frame(frame)
        img_frame.pack(fill='x', pady=10)
        tk.Label(img_frame, text="Imatge PNG o JPG:", font=("Arial", 10, "bold")).pack(anchor='w')
        self.img_label = tk.Label(img_frame, text="Cap imatge seleccionada", fg="gray")
        self.img_label.pack(anchor='w', pady=5)
        tk.Button(img_frame, text="Seleccionar Imatge", command=self.seleccionar_imatge,
                  bg="#2196F3", fg="white", font=("Arial", 10)).pack(anchor='w')

    def afegir_configuracio_capcalera(self, frame):
        # Text
        text_frame = tk.Frame(frame)
        text_frame.pack(fill='x', pady=10)
        tk.Label(text_frame, text="Text de la capçalera:", font=("Arial", 10, "bold")).pack(anchor='w')
        self.text_entry = tk.Entry(text_frame, width=50, font=("Arial", 10))
        self.text_entry.insert(0, "Capçalera del document")
        self.text_entry.pack(anchor='w', pady=5)

        coords_frame = tk.Frame(frame)
        coords_frame.pack(fill='x', pady=10)
        tk.Label(coords_frame, text="X:").pack(side='left')
        self.x_entry = tk.Entry(coords_frame, width=6)
        self.x_entry.insert(0, "20")
        self.x_entry.pack(side='left', padx=5)

        tk.Label(coords_frame, text="Y (des de dalt):").pack(side='left', padx=(20, 0))
        self.y_entry = tk.Entry(coords_frame, width=6)
        self.y_entry.insert(0, "20")
        self.y_entry.pack(side='left', padx=5)

        tk.Label(coords_frame, text="Amplada imatge:").pack(side='left', padx=(20, 0))
        self.width_entry = tk.Entry(coords_frame, width=6)
        self.width_entry.insert(0, "100")
        self.width_entry.pack(side='left', padx=5)

        tk.Label(coords_frame, text="Alçada:").pack(side='left')
        self.height_entry = tk.Entry(coords_frame, width=6)
        self.height_entry.insert(0, "50")
        self.height_entry.pack(side='left', padx=5)

        text_frame2 = tk.Frame(frame)
        text_frame2.pack(fill='x', pady=10)
        tk.Label(text_frame2, text="Mida text:").pack(side='left')
        self.text_size_entry = tk.Entry(text_frame2, width=6)
        self.text_size_entry.insert(0, "12")
        self.text_size_entry.pack(side='left', padx=5)

        tk.Label(text_frame2, text="Alineació del text:").pack(side='left', padx=(20, 0))
        self.text_align = tk.StringVar(value="dreta")
        for opt in ["esquerra", "centre", "dreta"]:
            tk.Radiobutton(text_frame2, text=opt.capitalize(), variable=self.text_align, value=opt).pack(side='left')

    def afegir_opcio_rectangle_fons(self, frame):
        rect_frame = tk.Frame(frame)
        rect_frame.pack(fill='x', pady=10)

        self.mostra_fons_var = tk.BooleanVar(value=True)
        tk.Checkbutton(rect_frame, text="Mostrar rectangle de fons", variable=self.mostra_fons_var).pack(side='left')

        tk.Label(rect_frame, text=" Color:").pack(side='left', padx=5)
        self.color_entry = tk.Entry(rect_frame, width=10)
        self.color_entry.insert(0, self.fons_color)
        self.color_entry.pack(side='left')

        tk.Button(rect_frame, text="Tria color", command=self.triar_color).pack(side='left', padx=5)

    def triar_color(self):
        color = colorchooser.askcolor(title="Tria un color")[1]
        if color:
            self.fons_color = color
            self.color_entry.delete(0, tk.END)
            self.color_entry.insert(0, color)

    def log_message(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def seleccionar_pdf(self):
        self.pdf_entrada = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_entrada:
            self.pdf_label.config(text=os.path.basename(self.pdf_entrada), fg="black")
            self.log_message(f"PDF seleccionat: {self.pdf_entrada}")

    def seleccionar_imatge(self):
        self.imatge_path = filedialog.askopenfilename(filetypes=[
            ("Imatges", "*.jpg *.jpeg *.png"), ("Tots els fitxers", "*.*")
        ])
        if self.imatge_path:
            self.img_label.config(text=os.path.basename(self.imatge_path), fg="black")
            self.log_message(f"Imatge seleccionada: {self.imatge_path}")

    def processar_pdf(self):
        try:
            if not self.pdf_entrada or not self.imatge_path:
                messagebox.showerror("Error", "Selecciona un PDF i una imatge.")
                return

            self.pdf_sortida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not self.pdf_sortida:
                return

            text = self.text_entry.get()
            x = int(self.x_entry.get())
            y_top = int(self.y_entry.get())
            width = int(self.width_entry.get())
            height = int(self.height_entry.get())
            text_size = int(self.text_size_entry.get())
            align = self.text_align.get()
            mostra_fons = self.mostra_fons_var.get()
            fons_color = self.color_entry.get()

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            page_width, page_height = A4
            y = page_height - y_top  # origen des de dalt

            with Image.open(self.imatge_path) as img:
                orig_w, orig_h = img.size
                aspect = orig_w / orig_h
                if width / height > aspect:
                    new_h = height
                    new_w = int(height * aspect)
                else:
                    new_w = width
                    new_h = int(width / aspect)

            # Dibuixar rectangle de fons si cal
            if mostra_fons:
                from reportlab.lib.colors import HexColor
                can.setFillColor(HexColor(fons_color))
                alt_rect = max(new_h, text_size + 10)
                y_rect = y - alt_rect
                can.rect(0, y_rect, page_width, alt_rect + 5, fill=1, stroke=0)

            # Imatge
            can.drawImage(self.imatge_path, x, y - new_h, width=new_w, height=new_h)

            # Text
            if align == "esquerra":
                text_x = x + new_w + 10
            elif align == "centre":
                text_x = (page_width - len(text) * text_size * 0.5) / 2
            else:  # dreta
                text_x = page_width - (len(text) * text_size * 0.5) - 20

            text_y = y - (new_h / 2) - (text_size / 2)
            can.setFont("Helvetica", text_size)
            can.setFillColorRGB(0, 0, 0)
            can.drawString(text_x, text_y, text)
            can.save()

            self.progress['value'] = 30
            with open(self.pdf_entrada, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                header = PyPDF2.PdfReader(io.BytesIO(packet.getvalue())).pages[0]

                for i, page in enumerate(reader.pages):
                    page.merge_page(header)
                    writer.add_page(page)
                    self.progress['value'] = 30 + int(60 * (i + 1) / len(reader.pages))
                    self.root.update()

                with open(self.pdf_sortida, 'wb') as f_out:
                    writer.write(f_out)

            self.progress['value'] = 100
            self.log_message("✓ PDF creat correctament.")
            messagebox.showinfo("Èxit", f"PDF guardat a:\n{self.pdf_sortida}")
        except Exception as e:
            self.log_message(f"✗ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.progress['value'] = 0

def obrir_llicencia(event):
    webbrowser.open_new(r"https://creativecommons.org/licenses/by-nc-sa/4.0/")

def main():
    root = tk.Tk()
    app = PDFEditor(root)
    
    # --- Peu de pàgina (Footer) ---
    footer_frame = tk.Frame(root, height=30, background="#f0f0f0")
    footer_frame.pack(side="bottom", fill="x")
    
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

if __name__ == "__main__":
    main()
