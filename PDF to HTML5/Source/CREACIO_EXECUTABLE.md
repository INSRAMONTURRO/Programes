# Creació d'executable PDF to HTML5 Flipbook Converter

## Opció 1: Amb PyInstaller (recomanat)

### Instal·lació de PyInstaller

```
pip install pyinstaller
```

### Creació de l'executable

#### A Linux:

```
pyinstaller --onefile --windowed --add-data "templates_extracted:templates_extracted" --add-data "templates_H_extracted:templates_H_extracted" --hidden-import=tkinter pdf_to_flipbook_cross_platform.py
```

#### A Windows:

```
pyinstaller --onefile --windowed --add-data "templates_extracted;templates_extracted" --add-data "templates_H_extracted;templates_H_extracted" --hidden-import=tkinter pdf_to_flipbook_cross_platform.py
```

NOTA: Si utilitzes les carpetes amb nom simple (`templates` i `templates_H`), ajusta els noms als comandos anteriors.

### Resultat

- L'executable es crearà a la carpeta `dist/`
- El nom del fitxer serà `pdf_to_flipbook_corrected_final.exe` (a Windows) o `pdf_to_flipbook_corrected_final` (a Linux)
- El fitxer inclou totes les dependències necessàries

## Opció 2: Amb cx_Freeze

### Instal·lació de cx_Freeze

```
pip install cx_Freeze
```

### Creació d'un fitxer setup.py

Crea un fitxer `setup.py` amb el següent contingut:

```python
from cx_Freeze import setup, Executable
import sys

# Opcions per a Windows
if sys.platform == "win32":
    include_files = ["templates_extracted", "templates_H_extracted"]
    base = "Win32GUI" if sys.platform == "win32" else None  # Amaga la consola a Windows
else:
    # Opcions per a Linux
    include_files = ["templates_extracted", "templates_H_extracted"]
    base = None

build_exe_options = {
    "packages": ["tkinter", "fitz", "PIL"],
    "include_files": include_files,
    "excludes": ["tkinter.test"]
}

setup(
    name="PDF to Flipbook Converter",
    version="1.0",
    description="Convertidor de PDF a llibres HTML5 interactius",
    options={"build_exe": build_exe_options},
    executables=[Executable("pdf_to_flipbook_cross_platform.py", base=base, target_name="pdf_flipbook_converter")]
)
```

### Executar la compilació

```
python setup.py build
```

Això crearà una carpeta `build/` amb l'executable i les llibreries necessàries.

## Llibreries necessàries

Assegura't que tens instal·lades les següents llibreries abans de compilar:

```
pip install PyMuPDF Pillow
```

## Fitxers necessaris

A més del fitxer Python, necessitaràs:

- `pdf_to_flipbook_cross_platform.py` - El fitxer principal del programa (versió cross-platform)
- `templates/` o `templates_extracted/` - Carpeta amb plantilla per al mode vertical (descomprimida)
- `templates_H/` o `templates_H_extracted/` - Carpeta amb plantilla per al mode horitzontal (descomprimida)

### Consideracions per a Windows

Aquesta versió del programa ja és totalment compatible amb Windows perquè utilitza carpetes descomprimides en lloc de fitxers `.tar.gz`.

Per generar l'executable, has d'incloure les carpetes de plantilles descomprimides dins de l'executable per assegurar la compatibilitat. Aquesta és la millor aproximació per a Windows.

Aquestes llibreries es poden empaquetar dins de l'executable o distribuir-les separadament.

## Execució de l'executable

Un cop creat, l'executable pot executar-se directament sense necessitat de tenir Python instal·lat al sistema:

### A Linux:

```
./pdf_to_flipbook_corrected_final
```

### A Windows:

Feu doble clic al fitxer `.exe` o executeu-lo des de la línia d'ordres:

```
pdf_to_flipbook_corrected_final.exe
```

## Recomanacions

- Prova l'executable en un sistema net per assegurar-te que totes les dependències estan incloses
- Si es produeixen errors de mòduls no trobats, afegiu `--hidden-import` per als mòduls problemàtics
- Per a Windows, l'opció `--windowed` evita que es mostri una finestra de consola