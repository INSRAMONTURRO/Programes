# Manual d'Ús: ReanomenaFitxers

Aquest programa permet copiar fitxers d'una carpeta a una altra i reanomenar-los de forma massiva amb un nom base i un número seqüencial.

## Funcionalitat

El programa ofereix una interfície gràfica senzilla per seleccionar una carpeta d'origen, una carpeta de destí i un nom base per als fitxers.

Un cop executat, el programa:

1. Copia tots els fitxers de la carpeta d'origen a la de destí.
2. Reanomena els fitxers copiats amb el format `nom-base-XXX.extensió`, on `XXX` és un número de tres dígits que va de `001` a `999`.

## Com Utilitzar

### Interfície Gràfica

La finestra principal del programa té els següents camps:

* **Carpeta Origen**: Aquí has de seleccionar la carpeta que conté els fitxers que vols copiar. Pots escriure la ruta directament o fer clic al botó **Seleccionar** per buscar-la.
* **Carpeta Destí**: Aquí has de seleccionar la carpeta on es guardaran els fitxers copiats i reanomenats. Si la carpeta no existeix, el programa la crearà.
* **Nom Base**: Aquest és el nom que s'utilitzarà per reanomenar els fitxers.

Un cop hagis omplert tots els camps, fes clic al botó **Copiar i Reanomenar** per iniciar el procés.

### Execució del Programa

Per executar el programa, no necessites instal·lar Python ni cap llibreria. Pots fer servir directament els executables proporcionats.

#### A Linux:

1. Obre un terminal.
2. Navega fins a la carpeta `ReanomenaFitxers/Linux`.
3. Dóna permisos d'execució a l'arxiu si és necessari amb la comanda:
   
   ```bash
   chmod +x reanomena
   ```
4. Executa el programa amb:
   
   ```bash
   ./reanomena
   ```

#### A Windows:

1. Navega fins a la carpeta `ReanomenaFitxers\windows`.
2. Fes doble clic sobre l'arxiu `reanomena.exe`.
