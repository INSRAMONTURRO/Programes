# Guia per Afegir un Nou Programa al Repositori

Aquest document detalla l'estructura i els passos a seguir cada cop que es vulgui afegir un nou programa al conjunt d'utilitats de l'**INS Ramon Turró**, i explica com funciona el procés de compilació automàtica (CI/CD) a GitHub.

---

## ❓ És necessari recompilar tots els programes cada cop?

**Sí, amb la configuració actual és imprescindible.** 

### Explicació:

Els enllaços de descàrrega del [README.md](README.md) principal apunten sempre a la darrera versió publicada a GitHub (`releases/latest/download/nom-programa`). 
Si el workflow de GitHub Actions només compilés el programa que s'ha modificat:

1. La release etiquetada com a "latest" (la més recent) **només tindria l'executable del programa modificat**.
2. Tots els altres enllaços de descàrrega del README es trencarien o donarien un error `404 Not Found` perquè anirien a buscar els seus executables a la darrera release, on ja no estarien.

Per tant, per mantenir tots els enllaços actius i en funcionament, el workflow compila i puja la totalitat dels programes del repositori a cada nova release (etiqueta `v*`).

---

## 🛠️ Estructura obligatòria d'un nou programa

Perquè GitHub Actions detecti i compili automàticament el nou programa, has de respectar estrictament la següent estructura de fitxers i carpetes:

```text
El teu repositori/
└── El_Nom_Del_Teu_Programa/           <-- Nom del directori (amb espais o guions)
    ├── Manual.md                      <-- Manual en Markdown (obligatori a l'arrel de la carpeta)
    └── Source/                        <-- Carpeta Source o source (obligatòria)
        ├── programa_principal.py      <-- El codi de Python principal
        ├── requeriments.txt           <-- Fitxer amb les llibreries pip necessàries
        ├── fitxers_addicionals.txt    <-- (Opcional) Llistat de recursos a incloure al .exe
        └── logo.png                   <-- (Opcional) Qualsevol imatge o recurs estàtic
```

### 1. La carpeta `Source/` (o `source/`)

El workflow busca específicament aquesta carpeta dins de cada programa. Si no hi és, l'ignora.

* El programa principal `.py` ha d'estar a dins de `Source/`.
* Si hi ha múltiples fitxers de codi, assegura't que el fitxer principal s'ordeni alfabèticament primer, o que sigui l'únic `.py` a l'arrel de `Source/` (el script agafa el primer fitxer `.py` que troba a la carpeta).

### 2. Gestió de Dependències (`requeriments.txt`)

Si el teu programa utilitza llibreries que no vénen per defecte a Python (ex: `pandas`, `openpyxl`, `pypdf`, `requests`, etc.):

* Crea un fitxer anomenat `requeriments.txt` (o `requirements.txt`) dins de `Source/`.
* Afegeix les llibreries una per línia (ex: `pypdf`). El servidor de GitHub les instal·larà abans de compilar.

### 3. Fitxers estàtics i llicències (`fitxers_addicionals.txt`)

Si el teu codi fa servir imatges, logos o plantilles, has de:

* Crear `fitxers_addicionals.txt` a dins de `Source/` i llistar el camí dels fitxers/carpetes a afegir (un per línia).

* Modificar el codi Python per a usar una funció de resolució de ruta dinàmica perquè PyInstaller localitzi correctament els recursos un cop empaquetat:
  
  ```python
  import sys
  import os
  
  def resource_path(relative_path):
      try:
          base_path = sys._MEIPASS
      except Exception:
          base_path = os.path.abspath(os.path.dirname(__file__))
      return os.path.join(base_path, relative_path)
  
  # Exemple de càrrega d'imatge:
  image_path = resource_path("logo.png")
  ```

### 4. El Manual d'Usuari (`Manual.md` o `manual.md`)

* Crea un fitxer de text en format Markdown a dins de la carpeta del programa.
* El workflow el copiarà automàticament a la carpeta de releases com a `nom-programa_manual.txt` perquè els usuaris el puguin descarregar.

---

## 📝 Passos per afegir i publicar el programa

### Pas 1: Crear la carpeta i els fitxers

Crea el teu projecte seguint l'estructura anterior (carpeta del projecte, carpeta `Source/` a dins, codi, manual i requeriments).

### Pas 2: Actualitzar el README general

Afegeix una nova fila a la taula del fitxer [README.md](README.md) de l'arrel del repositori.

#### Regla de normalització de noms de fitxer:

El workflow transforma el nom de la carpeta del programa per a generar els noms dels executables. La regla és:

1. Tot en minúscules.
2. Els espais (` `) es converteixen en guions (`-`).
3. Els guions baixos (`_`) es converteixen en guions (`-`).

*Exemple:* 

* Carpeta: `SepararPDF_6pag` $\rightarrow$ Executable: `separarpdf-6pag.exe` (Windows), `separarpdf-6pag` (Linux), Manual: `separarpdf-6pag_manual.txt`.
* Carpeta: `PDF to HTML5` $\rightarrow$ Executable: `pdf-to-html5.exe` (Windows), `pdf-to-html5` (Linux), Manual: `pdf-to-html5_manual.txt`.

#### Exemple de fila a afegir al README.md:

```markdown
| **El meu Programa** | Descripció curta del programa | [⬇️ Descarregar](https://github.com/INSRAMONTURRO/Programes/releases/latest/download/el-meu-programa.exe) | [⬇️ Descarregar](https://github.com/INSRAMONTURRO/Programes/releases/latest/download/el-meu-programa) | [📖 MANUAL](https://github.com/INSRAMONTURRO/Programes/releases/latest/download/el-meu-programa_manual.txt) |
```

### Pas 3: Pujar els canvis i crear la versió (Tag)

Obre un terminal al directori del projecte i executa:

1. **Afegeix els nous fitxers a git:**
   
   ```bash
   git add .
   ```

2. **Fes un commit descriptiu:**
   
   ```bash
   git commit -m "feat: Afegit nou programa [Nom del Programa]"
   ```

3. **Puja el commit a GitHub:**
   
   ```bash
   git push
   ```

4. **Crea una nova etiqueta incrementant la versió** (pots comprovar la darrera executant `git tag`):
   
   ```bash
   # Si la darrera era v1.87, creem la v1.88
   git tag -a v1.88 -m "Release v1.88: S'afegeix el programa [Nom del programa]"
   ```

5. **Puja l'etiqueta per activar el compilador:**
   
   ```bash
   git push origin v1.88
   ```

Un cop pujada l'etiqueta, podràs veure el progrés a la secció **Actions** de GitHub, i en uns minuts la descàrrega estarà disponible de forma totalment autònoma.

---

## 🤖 Directrius per a Assistents d'IA (com Antigravity)

Si un usuari demana a un assistent d'IA que afegeixi, mogui o actualitzi un programa d'aquest repositori, l'assistent **ha de dur a terme automàticament els següents passos**:

1. **Reorganitzar o crear l'estructura de carpetes:**
   - Assegurar-se que els fitxers de codi i requisits estan dins d'una carpeta `Source/` (o `source/`).
   - Crear o actualitzar el `Manual.md` corresponent a l'arrel de la carpeta del programa.
   - Si s'inclouen fitxers estàtics (logos, llicències...), configurar `fitxers_addicionals.txt` i comprovar que el codi empra `resource_path`.
2. **Actualitzar de manera autònoma el `README.md` general:**
   - Obtenir el nom de la carpeta del programa.
   - Normalitzar el nom per als enllaços de descàrrega de releases (aplicant minúscules i convertint espais/guions baixos en guions mitjans).
   - Localitzar la taula de programes del `README.md` de l'arrel del repositori i afegir-hi una nova fila amb els enllaços de descàrrega normalitzats i el manual de forma totalment automàtica, sense necessitat d'esperar que l'usuari ho demani.
3. **Executar la pujada i l'etiquetatge:**
   - Fer `git add .`, fer el commit, demanar confirmació per a fer `git push` a `main`.
   - Consultar les etiquetes existents (`git tag`), calcular la versió següent (incrementant la versió més recent, com per exemple de `v1.87` a `v1.88`) i proposar fer `git tag` i `git push origin vX.XX` per activar les GitHub Actions de compilació.
