# Manual d'Ús: PDF_to_HTML5

Aquest programa converteix un document PDF en un "flipbook" interactiu en format HTML5. El resultat és un paquet de fitxers web que simulen l'experiència de passar les pàgines d'un llibre o revista digital.

## Funcionalitat

El programa ofereix una interfície gràfica per seleccionar un fitxer PDF i personalitzar la conversió. Les seves característiques principals són:

* Converteix les pàgines d'un PDF en imatges d'alta qualitat.
* Permet escollir la **resolució (DPI)** de les imatges per equilibrar qualitat i pes.
* Ofereix dos **modes de visualització** per al flipbook final.
* Empaqueta tot el projecte web (HTML, CSS, JavaScript i imatges) en un únic fitxer **ZIP** per a una fàcil distribució.

## Com Utilitzar

### Interfície Gràfica

1. **Fitxer PDF**: Fes clic a `Cercar PDF` per seleccionar el document que vols convertir.
2. **Resolució (DPI)**: Introdueix la resolució per a la conversió de les pàgines. Un valor més alt (`300`) ofereix més qualitat però genera fitxers més pesats. Un valor més baix (`150`) és més ràpid i lleuger. El valor recomanat és `150`.
3. **Tipus de Visualització**: Tria com vols que es vegi el teu flipbook.
   * **Llibre/Revista (2 pàgines)**: Aquesta opció simula un llibre obert, mostrant dues pàgines a la vegada. Inclou un efecte realista de passada de pàgina. Ideal per a revistes, llibres i catàlegs.
   * **Fulla Única (1 pàgina)**: Aquesta opció mostra una sola pàgina a la vegada. La navegació es fa amb botons de "Següent" i "Anterior". És una visualització més simple, adequada per a documents o presentacions.
4. **Convertir**: Fes clic al botó `Convertir PDF a Flipbook` per iniciar el procés.

Un cop finalitzat, el programa crearà un fitxer anomenat `[nom_del_pdf]_flipbook.zip` a la mateixa carpeta on es troba el PDF original.

### El Fitxer de Sortida

El resultat és un fitxer ZIP. Per veure el flipbook, has de:

1. Descomprimir el fitxer ZIP.

2. Obrir la carpeta generada.

3. Fer doble clic a l'arxiu `index.html`. Això obrirà el flipbook al teu navegador web.

### ### Execució del Programa

Per executar el programa, no necessites instal·lar Python ni cap llibreria. Pots fer servir directament els executables proporcionats.

#### A Linux:

1. Obre un terminal.

2. Navega fins a la carpeta `PDF_to_HTML5/Linux`.

3. Dóna permisos d'execució a l'arxiu si és necessari amb la comanda:
   
   ```bash
   chmod +x pdf_to_flipbook_HTML5
   ```

4. Executa el programa amb:
   
   ```bash
   ./pdf_to_flipbook_HTML5
   ```

#### A Windows:

1. Navega fins a la carpeta `PDF_to_HTML5\windows`.
2. Aquesta versió de Windows no està disponible.

Com Pujar el Flipbook a Moodle

Perquè els teus alumnes puguin veure el flipbook directament a Moodle, has de pujar-lo com un paquet de contingut.

1. **Activa el mode d'edició** al teu curs de Moodle.

2. Fes clic a **Afegeix una activitat o un recurs** a la secció on vulguis afegir el flipbook.

3. Selecciona el recurs **Fitxer** (`File`).

4. Posa-li un **Nom** (per exemple, "Revista Digital de la Classe").

5. A la secció **Selecciona fitxers**, arrossega i deixa anar el fitxer `[nom_del_pdf]_flipbook.zip` que has creat.

6. Un cop pujat, fes clic sobre el fitxer ZIP i selecciona l'opció **Descomprimeix** (`Unzip`). Moodle crearà una carpeta amb tot el contingut del flipbook.

7. Busca el fitxer `index.html` principal a la llista de fitxers descomprimits (hauria d'estar al nivell arrel), fes-hi clic i selecciona **Defineix com a fitxer principal** (`Set as main file`).

8. A la secció **Aparença** (`Appearance`), canvia el paràmetre **Mostra** (`Display`) a **En una finestra emergent** (`In pop-up`) o **Incrusta** (`Embed`). Això evitarà que es descarregui automàticament.

9. Fes clic a **Desa i torna al curs**.

Ara, quan un estudiant faci clic a l'enllaç, el flipbook s'obrirà i es podrà navegar directament des de Moodle.
