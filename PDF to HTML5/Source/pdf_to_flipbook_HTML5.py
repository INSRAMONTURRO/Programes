#!/usr/bin/env python3
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
"""
PDF to HTML5 Flipbook Converter (Cross-Platform Version)
A program to convert PDF files to interactive HTML5 flipbooks using professional templates
Compatible with both Windows and Linux
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image
import tempfile
import os
import shutil
from pathlib import Path
import zipfile
import webbrowser
from tkinter import font
import sys

def resource_path(relative_path):
    """ Retorna la ruta absoluta al recurs, funciona per a desenvolupament i per a PyInstaller """
    try:
        # PyInstaller crea una carpeta temporal i guarda la ruta a _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    
    return os.path.join(base_path, relative_path)

class PDFToFlipbookConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF a Flipbook HTML5 - Versió Professional")
        self.root.geometry("450x600")  # Augmentem encara més la mida de la finestra
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.resolution = tk.IntVar(value=150)
        self.layout_type = tk.StringVar(value="vertical")  # vertical or horizontal
        
        self.setup_ui()
    
    def setup_ui(self):
        # Configure style
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Convertidor PDF a Flipbook HTML5", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # PDF Selection
        ttk.Label(main_frame, text="Fitxer PDF:", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        pdf_frame = ttk.Frame(main_frame)
        pdf_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=55).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(pdf_frame, text="Cercar PDF", command=self.select_pdf, width=12).grid(row=0, column=1, padx=(10, 0))
        pdf_frame.columnconfigure(0, weight=1)
        
        # Resolution
        ttk.Label(main_frame, text="Resolució (DPI):", font=("Arial", 11, "bold")).grid(row=3, column=0, sticky=tk.W, pady=(20, 5))
        resolution_frame = ttk.Frame(main_frame)
        resolution_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        resolution_spinbox = ttk.Spinbox(resolution_frame, from_=72, to=600, textvariable=self.resolution, width=10)
        resolution_spinbox.grid(row=0, column=0, sticky=tk.W)
        ttk.Label(resolution_frame, text="DPI (72-600, recomanat 150)").grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # Layout Type
        ttk.Label(main_frame, text="Tipus de Visualització:", font=("Arial", 11, "bold")).grid(row=5, column=0, sticky=tk.W, pady=(20, 10))
        
        layout_frame = ttk.LabelFrame(main_frame, text="Opcions de Visualització", padding="12")
        layout_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=8)
        
        # Create labels with custom fonts instead of using font in Radiobutton directly
        rb1_frame = ttk.Frame(layout_frame)
        rb1_frame.grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Radiobutton(rb1_frame, variable=self.layout_type, value="vertical").grid(row=0, column=0, sticky=tk.W)
        rb1_text = ttk.Label(rb1_frame, text="Llibre/Revista (2 pàgines - Efecte realista)")
        rb1_text.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        rb1_text.configure(font=("Arial", 10))
        
        rb2_frame = ttk.Frame(layout_frame)
        rb2_frame.grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Radiobutton(rb2_frame, variable=self.layout_type, value="horizontal").grid(row=0, column=0, sticky=tk.W)
        rb2_text = ttk.Label(rb2_frame, text="Fulla Única (1 pàgina - Visualització simple)")
        rb2_text.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        rb2_text.configure(font=("Arial", 10))
        
        # Description for each layout
        desc1 = ttk.Label(layout_frame, text="• Revista: Visualització com un llibre real amb voltes de pàgina", foreground="gray")
        desc1.grid(row=2, column=0, sticky=tk.W, padx=(25, 0), pady=2)
        desc1.configure(font=("Arial", 9))
        
        desc2 = ttk.Label(layout_frame, text="• Fulla Única: Visualització simple d'una pàgina alhora", foreground="gray")
        desc2.grid(row=3, column=0, sticky=tk.W, padx=(25, 0), pady=2)
        desc2.configure(font=("Arial", 9))
        
        # Convert Button - Ara estarà ben visible
        convert_btn = ttk.Button(main_frame, text="Convertir PDF a Flipbook", 
                                command=self.convert_pdf, style="Accent.TButton", width=30)
        convert_btn.grid(row=7, column=0, columnspan=2, pady=25, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=500)
        self.progress.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text=" → Selecciona un fitxer PDF per començar", 
                                     foreground="darkblue", font=("Arial", 10, "bold"))
        self.status_label.grid(row=9, column=0, columnspan=2, pady=10)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(9, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Selecciona un fitxer PDF",
            filetypes=[("Fitxers PDF", "*.pdf"), ("Tots els fitxers", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            self.status_label.config(text=f" ✓ PDF seleccionat: {Path(file_path).name}")
    
    def convert_pdf(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Si us plau, selecciona un fitxer PDF")
            return
        
        if not Path(self.pdf_path.get()).exists():
            messagebox.showerror("Error", "El fitxer PDF no existeix")
            return
        
        try:
            self.status_label.config(text=" → Convertint PDF...")
            self.progress['value'] = 0
            self.root.update()
            
            # Start the conversion process
            self.process_pdf()
            
        except Exception as e:
            messagebox.showerror("Error", f"S'ha produït un error durant la conversió:\n{str(e)}")
            self.status_label.config(text=" → Error durant la conversió")
    
    def process_pdf(self):
        pdf_path = Path(self.pdf_path.get())
        title = pdf_path.stem  # Use the PDF filename without extension as title
        
        # Use a working directory in the same location as the PDF
        working_dir = pdf_path.parent
        with tempfile.TemporaryDirectory(dir=working_dir) as temp_dir:
            temp_path = Path(temp_dir)
            
            # Copy templates based on selected layout
            if self.layout_type.get() == "vertical":
                self.copy_vertical_templates(temp_path)
                book_path = temp_path / "book"
            else:
                self.copy_horizontal_templates(temp_path)
                book_path = temp_path / "book"
            
            # Ensure pages directory exists
            pages_path = book_path / "pages"
            pages_path.mkdir(parents=True, exist_ok=True)
            
            # Extract PDF to images
            self.status_label.config(text=" → Extreient pàgines del PDF...")
            self.root.update()
            
            doc = fitz.open(pdf_path)
            num_pages = len(doc)
            
            # Update progress step
            progress_step = 50 / num_pages if num_pages > 0 else 0
            
            for i, page in enumerate(doc):
                page_num = i + 1
                self.status_label.config(text=f" → Processant pàgina {page_num} de {num_pages}...")
                self.progress['value'] = 20 + (i * progress_step)
                self.root.update()
                
                # Create high resolution image
                pix = page.get_pixmap(dpi=self.resolution.get())
                output_path = pages_path / f"{page_num}.jpg"
                pix.save(str(output_path), "jpeg")
                
                # Create thumbnail
                with Image.open(output_path) as img:
                    # Create thumbnail maintaining aspect ratio
                    img.thumbnail((120, 160))
                    img.save(str(pages_path / f"{page_num}-thumb.jpg"), "jpeg")
            
            doc.close()
            
            # Create HTML structure based on selected layout
            self.status_label.config(text=" → Generant HTML amb plantilles...")
            self.progress['value'] = 70
            self.root.update()
            
            if self.layout_type.get() == "vertical":
                self.create_vertical_flipbook_with_templates(temp_path, book_path, title, num_pages)
            else:
                self.create_horizontal_flipbook_with_templates(temp_path, book_path, title, num_pages)
            
            # Package the result
            self.status_label.config(text=" → Empaquetant resultat...")
            self.progress['value'] = 90
            self.root.update()
            
            output_path = Path(pdf_path.parent) / f"{title}_flipbook.zip"
            self.create_zip(temp_path, output_path)
            
            self.progress['value'] = 100
            self.status_label.config(text=f" ✓ Completat! Fitxer creat: {output_path.name}")
            messagebox.showinfo("Completat", f"El flipbook s'ha creat correctament:\n\n{output_path}\n\nMode: {'Llibre/Revista' if self.layout_type.get() == 'vertical' else 'Fulla Única'}\nResolució: {self.resolution.get()} DPI\nPàgines: {num_pages}")
            self.status_label.config(text=f" ✓ Completat! Fitxer creat: {output_path.name}")

    def copy_vertical_templates(self, temp_path):
        """Copy vertical templates from templates_extracted folder"""
        templates_path = Path(resource_path("templates_extracted"))
        if not templates_path.exists():
            # Try with templates folder as alternative
            templates_path = Path(resource_path("templates"))
            if not templates_path.exists():
                raise FileNotFoundError(f"No s'ha trobat la carpeta de plantilles: {templates_path}")
        
        # Copy all contents from templates to temp_path
        for item in templates_path.iterdir():
            if item.is_dir():
                shutil.copytree(item, temp_path / item.name)
            else:
                shutil.copy2(item, temp_path / item.name)

    def copy_horizontal_templates(self, temp_path):
        """Copy horizontal templates from templates_H_extracted folder"""
        templates_path = Path(resource_path("templates_H_extracted"))
        if not templates_path.exists():
            # Try with templates_H folder as alternative
            templates_path = Path(resource_path("templates_H"))
            if not templates_path.exists():
                raise FileNotFoundError(f"No s'ha trobat la carpeta de plantilles horitzontals: {templates_path}")
        
        # For horizontal templates, we need to handle the nested templates_H structure
        templates_h_path = templates_path / "templates_H"
        if templates_h_path.exists():
            # Copy all contents from templates_H subfolder
            for item in templates_h_path.iterdir():
                if item.is_dir():
                    shutil.copytree(item, temp_path / item.name)
                else:
                    shutil.copy2(item, temp_path / item.name)
        else:
            # If no templates_H subfolder, copy everything directly
            for item in templates_path.iterdir():
                if item.is_dir():
                    shutil.copytree(item, temp_path / item.name)
                else:
                    shutil.copy2(item, temp_path / item.name)

    def create_vertical_flipbook_with_templates(self, temp_path, book_path, title, num_pages):
        """Create the vertical flipbook using the original templates from templates_extracted"""
        # Read template files and customize them
        top_html_path = book_path / "plantilla_index_top.html"
        bottom_html_path = book_path / "plantilla_index_bottom.html"
        
        if not top_html_path.exists() or not bottom_html_path.exists():
            raise FileNotFoundError(f"No s'han trobat les plantilles necessàries: {top_html_path} o {bottom_html_path}")
        
        top_html = top_html_path.read_text(encoding="utf-8")
        bottom_html = bottom_html_path.read_text(encoding="utf-8")
        
        # Customize the templates with our specific values
        top_html = top_html.replace("_CADENA_PER_CANVIAR_", title)
        # Fix: Replace the pages count in the JavaScript part - proper formatting
        bottom_html = bottom_html.replace("_N_PAGES_", str(num_pages))
        
        # Write the customized index.html to the book directory (not to root)
        book_index_path = book_path / "index.html"
        book_index_path.write_text(top_html + bottom_html, encoding="utf-8")
        
        # Keep the original root index.html as is (it should have been preserved from template extraction)
        # If there's no root index.html, we need to create one that points to the book version
        root_index_path = temp_path / "index.html"
        if not root_index_path.exists():
            # Create a simple root index that redirects to the book version
            redirect_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="refresh" content="0; url=book/index.html">
    <title>{title}</title>
</head>
<body>
    <p>Si no es redirigeix automàticament, <a href="book/index.html">fes clic aquí</a>.</p>
</body>
</html>"""
            root_index_path.write_text(redirect_html, encoding="utf-8")
        
        # Remove the template files
        top_html_path.unlink()
        bottom_html_path.unlink()

    def create_horizontal_flipbook_with_templates(self, temp_path, book_path, title, num_pages):
        """Update the horizontal flipbook book/index.html with title and page count"""
        # Look for the main index.html file in the book directory
        book_index_path = book_path / "index.html"
        
        if book_index_path.exists():
            # If there's already an index.html in book/, customize it
            html_content = book_index_path.read_text(encoding="utf-8")
            
            # Customize the template with our specific values
            html_content = html_content.replace("_CADENA_PER_CANVIAR_", title)
            html_content = html_content.replace("_N_PAGES_", str(num_pages))
            
            # Write the customized index.html back
            book_index_path.write_text(html_content, encoding="utf-8")
        else:
            # If there's no index.html in book/, look for templates or create a basic one
            # First, check if there are template files similar to vertical mode
            top_template_path = book_path / "plantilla_index_top.html"
            bottom_template_path = book_path / "plantilla_index_bottom.html"
            
            if top_template_path.exists() and bottom_template_path.exists():
                # Use template system similar to vertical mode
                top_html = top_template_path.read_text(encoding="utf-8")
                bottom_html = bottom_template_path.read_text(encoding="utf-8")
                
                # Customize the templates with our specific values
                top_html = top_html.replace("_CADENA_PER_CANVIAR_", title)
                bottom_html = bottom_html.replace("_N_PAGES_", str(num_pages))
                
                # Write the customized index.html to the book directory
                book_index_path.write_text(top_html + bottom_html, encoding="utf-8")
                
                # Remove the template files
                top_template_path.unlink()
                bottom_template_path.unlink()
            else:
                # Create a basic horizontal flipbook with the same structure as vertical
                self.create_basic_horizontal_flipbook(temp_path, book_path, title, num_pages)

    def create_basic_horizontal_flipbook(self, temp_path, book_path, title, num_pages):
        """Create a basic horizontal flipbook with the same directory structure as vertical"""
        # Since both modes should have the same structure, we just need to create
        # a book/index.html that shows one page at a time instead of two
        
        # Read any existing index.html in the book directory to use as base
        original_index_path = book_path / "index.html"
        
        # If there's no existing index.html, create a basic one based on the horizontal approach
        if not original_index_path.exists():
            # Create a basic horizontal flipbook structure
            # We'll use the same CSS and JS paths as the vertical mode but with different functionality
            
            html_content = f"""<!DOCTYPE html>
<html lang="ca">
<head>
    <meta charset="UTF-8">
    <title>_CADENA_PER_CANVIAR_</title>
    <link rel="stylesheet" type="text/css" href="css/magazine.css">
    <script type="text/javascript" src="../extras/jquery.min.js"></script>
    <style>
        body {{
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: space-between;
            height: 100vh;
            margin: 0;
            background: linear-gradient(to bottom, #8b743d 0%, #c8af6b 100%);
            overflow: hidden;
            color: #333;
        }}
        header {{
            width: 100%;
            text-align: center;
            padding: 15px 0;
            background-color: rgba(0, 0, 0, 0.3);
            color: white;
            flex-shrink: 0;
            font-size: 1.4em;
            font-weight: bold;
        }}
        main {{
            flex-grow: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            overflow: auto;
            padding: 10px;
        }}
        #viewer {{
            text-align: center;
            max-width: 95%;
            max-height: 85vh;
            background: white;
            border-radius: 8px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.3);
            padding: 20px;
        }}
        #page-image {{
            max-width: 100%;
            max-height: 100%;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            border: 2px solid #ddd;
            object-fit: contain;
            background: white;
        }}
        footer {{
            width: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px 0;
            background-color: rgba(0, 0, 0, 0.3);
            color: white;
            flex-shrink: 0;
        }}
        footer button {{
            padding: 12px 25px;
            font-size: 1.1em;
            cursor: pointer;
            border: none;
            background: linear-gradient(to bottom, #4CAF50, #45a049);
            color: white;
            border-radius: 6px;
            box-shadow: 0 3px 6px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
            margin: 0 15px;
            font-weight: bold;
        }}
        footer button:hover {{
            background: linear-gradient(to bottom, #45a049, #3d8b40);
            transform: translateY(-3px);
            box-shadow: 0 5px 10px rgba(0,0,0,0.3);
        }}
        footer button:active {{
            transform: translateY(0);
        }}
        footer button:disabled {{
            background: #cccccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }}
        #page-indicator {{
            margin: 0 20px;
            font-size: 1.3em;
            font-weight: bold;
            color: white;
            background: rgba(0, 0, 0, 0.4);
            padding: 8px 20px;
            border-radius: 25px;
            min-width: 120px;
            text-align: center;
        }}
    </style>
</head>
<body>
    <header>
        <h1>_CADENA_PER_CANVIAR_</h1>
    </header>
    <main>
        <div id="viewer">
            <img id="page-image" src="pages/1.jpg" alt="Pàgina 1">
        </div>
    </main>
    <footer>
        <button id="prev-btn" title="Pàgina anterior">‹ Anterior</button>
        <span id="page-indicator">1 / _N_PAGES_</span>
        <button id="next-btn" title="Pàgina següent">Següent ›</button>
    </footer>
    <script>
        // JavaScript per a la navegació d'una sola pàgina
        document.addEventListener('DOMContentLoaded', () => {{
            let currentPage = 1;
            const totalPages = {num_pages};
            const pageImage = document.getElementById('page-image');
            const pageIndicator = document.getElementById('page-indicator');
            const prevBtn = document.getElementById('prev-btn');
            const nextBtn = document.getElementById('next-btn');

            function showPage(pageNumber) {{
                if (pageNumber < 1 || pageNumber > totalPages) {{
                    return;
                }}
                currentPage = pageNumber;
                const imagePath = `pages/${{currentPage}}.jpg`;
                
                // Update image source with timestamp to prevent caching issues
                const timestamp = new Date().getTime();
                pageImage.src = imagePath + '?' + timestamp;
                
                pageImage.alt = `Pàgina {{currentPage}}`;
                pageIndicator.textContent = `{{currentPage}} / {{totalPages}}`;
                
                prevBtn.disabled = (currentPage === 1);
                nextBtn.disabled = (currentPage === totalPages);
                
                // Update page URL for bookmarking
                if (history.pushState) {{
                    history.pushState(null, null, `#/page/{{currentPage}}`);
                }}
            }}

            prevBtn.addEventListener('click', () => {{
                if (currentPage > 1) {{
                    showPage(currentPage - 1);
                }}
            }});

            nextBtn.addEventListener('click', () => {{
                if (currentPage < totalPages) {{
                    showPage(currentPage + 1);
                }}
            }});

            // Navegació amb tecles de fletxa
            document.addEventListener('keydown', (e) => {{
                if (e.key === 'ArrowLeft' && currentPage > 1) {{
                    showPage(currentPage - 1);
                }} else if (e.key === 'ArrowRight' && currentPage < totalPages) {{
                    showPage(currentPage + 1);
                }}
            }});

            // Handle hash changes for direct page navigation
            window.addEventListener('hashchange', function() {{
                const hash = window.location.hash.substring(2); // Remove '#/'
                if (hash.startsWith('page/')) {{
                    const pageNum = parseInt(hash.split('/')[1]);
                    if (!isNaN(pageNum)) {{
                        showPage(pageNum);
                    }}
                }}
            }});

            // Initialize with first page
            showPage(1);
        }});
    </script>
</body>
</html>"""

            # Write the customized index.html to book directory
            book_index_path.write_text(html_content, encoding="utf-8")
        else:
            # If there is an existing index.html, just replace the placeholders
            html_content = original_index_path.read_text(encoding="utf-8")
            html_content = html_content.replace("_CADENA_PER_CANVIAR_", title)
            html_content = html_content.replace("_N_PAGES_", str(num_pages))
            original_index_path.write_text(html_content, encoding="utf-8")

    def create_zip(self, root_path, output_path):
        """Create a zip file with the flipbook content with correct structure"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(root_path):
                for file in files:
                    file_path = Path(root) / file
                    # Calculate the relative path from the temporary directory
                    arc_path = file_path.relative_to(root_path)
                    zipf.write(file_path, arc_path)

def obrir_llicencia(event):
    webbrowser.open_new(r"https://creativecommons.org/licenses/by-nc-sa/4.0/")

def main():
    root = tk.Tk()
    app = PDFToFlipbookConverter(root)

    # --- Peu de pàgina (Footer) ---
    footer_frame = tk.Frame(root, height=30, background="#f0f0f0")
    footer_frame.grid(row=1, column=0, sticky="ew")
    
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