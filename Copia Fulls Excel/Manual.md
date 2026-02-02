# Processador d'Excel d'Actes amb Imatges

## Descripció General

Aquest programa és una eina d'escriptori amb interfície gràfica en català, dissenyada per automatitzar el processament de fitxers Excel d'actes acadèmiques. La seva funció principal és agafar un fitxer Excel d'origen, duplicar una de les seves fulles de càlcul diverses vegades, i desar el resultat com un nou fitxer a una carpeta de destinació.

Aquesta versió millorada permet, a més, afegir un logotip o imatge a cada fulla generada amb dimensions precises i personalitzar el nom dels fitxers de sortida.

## Funcionalitats Principals

* **Interfície gràfica amigable**: Interfície d'usuari en català amb selectors de carpetes i d'imatge.
* **Processament de fitxers en lot**: Processa tots els fitxers Excel (`.xlsx`) d'una carpeta d'origen.
* **Còpia de fulles**: Copia la fulla "1aAVA" (o la primera fulla si "1aAVA" no existeix) a quatre noves fulles: "2aAVA", "3aAVA", "Final" i "Extraor", conservant el format.
* **Addició d'imatges**: Afegeix una imatge seleccionada per l'usuari a totes les fulles del document Excel.
* **Configuració d'impressió**: Aplica una configuració d'impressió uniforme a totes les fulles.
* **Reanomenament intel·ligent**: Els fitxers de sortida es reanomenen segons un prefix seleccionable i els noms de grup definits en un fitxer `Noms.csv`.
* **Barra de progrés i notificacions**: Mostra l'estat del procés i informa sobre l'evolució i possibles errors.

## Requisits

### Software

* **Python 3.x**: El programa està escrit en Python. Si no el teniu instal·lat, el podeu descarregar des de [python.org](https://www.python.org/).
* **Pip**: El gestor de paquets de Python, que normalment ve inclòs amb la instal·lació de Python.

### Biblioteques de Python

El programa depèn de les següents biblioteques, llistades al fitxer `requirements.txt`:

* `openpyxl`
* `pandas`

### Fitxers Necessaris

* **`Noms.csv`**: Un fitxer en format CSV que conté una llista de noms de grups o classes (p. ex., "1 ESO A"). S'utilitza per determinar el nom dels fitxers de sortida.
* **Imatge del logotip**: Un fitxer d'imatge (formats suportats: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff`, `.webp`) que es vulgui afegir a les fulles de càlcul.

## Instal·lació

Per instal·lar les dependències necessàries, obriu un terminal o línia de comandes, navegueu fins a la carpeta on es troba el programa i executeu la següent comanda:

```bash
pip install -r requirements.txt
```

## Ús del Programa

1. **Executa el programa**:
   
   ```bash
   python Copia_Format_fulls_v5.py
   ```
2. **Selecciona la Carpeta Origen** que conté els fitxers Excel a processar.
3. **Selecciona la Carpeta Destí** on es desaran els fitxers processats.
4. **Selecciona la Imatge a afegir** (opcional).
5. **Tria el Prefix del nom** desitjat per als fitxers de sortida (p. ex., "1aAVA", "2aAVA").
6. **Fes clic al botó "Processa Fitxers"** per iniciar el procés.

## Fitxers d'Entrada

### `Noms.csv`

Aquest fitxer ha de contenir un nom de grup o classe per línia. El programa buscarà aquests noms als noms dels fitxers Excel d'origen per generar els noms dels fitxers de sortida.

*Exemple de `Noms.csv`*:

```csv
1 ESO A
1 ESO B
1 BTX A
```

### Fitxers Excel d'Origen

Els fitxers Excel d'origen han de ser en format `.xlsx`. El programa no processa fitxers temporals (que comencen amb `~$`). Buscarà una fulla anomenada "1aAVA" per duplicar-la; si no la troba, utilitzarà la primera fulla del llibre.

## Fitxers de Sortida

Els fitxers de sortida es desaran a la carpeta de destinació. El nom es generarà segons el patró `Acta_[PREFIX]_[NOM_GRUP].xlsx`.

*Exemple de noms de fitxer*:

- Si el prefix és "1aAVA" i el grup "1 BTX A": `Acta_1aAVA_1_BTX_A.xlsx`
- Si el prefix és "Final" i el grup "2 ESO B": `Acta_Final_2_ESO_B.xlsx`

Cada fitxer de sortida contindrà la fulla original duplicada amb els noms "2aAVA", "3aAVA", "Final", i "Extraor". Totes les fulles tindran la imatge seleccionada afegida.

## Configuració d'Impressió

Totes les fulles tenen aplicades les següents configuracions d'impressió:

* **Orientació**: Vertical ("portrait")
* **Mida del paper**: A4
* **Ajust de pàgina**: Ajustat a 1 pàgina d'amplada per 2 d'alçada.
* **Marges**: 0.39375 polzades (aprox. 1 cm) a tots els costats.
* **Opcions d'impressió**: Sense línies de graella ni capçaleres de fila/columna.

## Mides de la Imatge

La imatge afegida a cada fulla tindrà les següents dimensions per assegurar una presentació uniforme:

* **Amplada**: 1,84 cm
* **Alçada**: 1,19 cm

## Formats Suportats

* **Fitxers Excel**: `.xlsx`
* **Imatges**: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff`, `.webp`

## Notes Importants

* El programa manté totes les dades i el format de les fulles originals.
* Les imatges s'afegeixen a la posició predeterminada (propera a la cel·la A1).
* En cas d'errors durant el processament d'un fitxer, es mostrarà un missatge d'advertència i el programa continuarà amb el següent.

## Autoria i Llicència

* **Autor:** Josep Maria Sardà
* **Llicència:** Creative Commons BY-NC-SA 4.0
