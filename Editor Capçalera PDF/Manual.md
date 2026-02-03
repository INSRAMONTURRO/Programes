# Manual d'Ús: Editor_PDF_cap

Aquest programa permet afegir una capçalera personalitzada a totes les pàgines d'un document PDF. La capçalera pot incloure una imatge (logotip) i una línia de text.

## Funcionalitat

El programa ofereix una interfície gràfica per configurar i afegir una capçalera a un fitxer PDF. Les principals funcionalitats són:

* Seleccionar un fitxer PDF d'entrada.
* Seleccionar una imatge (JPG o PNG) per a la capçalera.
* Personalitzar el text de la capçalera.
* Ajustar la posició i la mida de la imatge.
* Ajustar la mida i l'alineació del text.
* Afegir un rectangle de fons opcional a la capçalera i escollir-ne el color.
* Guardar el resultat com un nou fitxer PDF.

## Com Utilitzar

### Interfície Gràfica

La finestra principal del programa es divideix en diverses seccions:

1. **Selecció de Fitxers**:
   
   * **PDF d'entrada**: Fes clic a `Seleccionar PDF` per triar el document que vols modificar.
   * **Imatge PNG o JPG**: Fes clic a `Seleccionar Imatge` per triar la imatge que vols afegir a la capçalera.

2. **Configuració de la Capçalera**:
   
   * **Text de la capçalera**: Escriu el text que vols que aparegui al costat de la imatge.
   * **Coordenades i Mides**:
     * `X`: Posició horitzontal de la imatge des del marge esquerre.
     * `Y (des de dalt)`: Posició vertical de la imatge des del marge superior.
     * `Amplada imatge`: Amplada màxima que ocuparà la imatge.
     * `Alçada`: Alçada màxima que ocuparà la imatge.
   * **Configuració del Text**:
     * `Mida text`: Mida de la font del text.
     * `Alineació del text`: Alineació del text respecte a la pàgina (Esquerra, Centre, Dreta).

3. **Rectangle de Fons**:
   
   * **Mostrar rectangle de fons**: Marca aquesta casella si vols que la capçalera tingui un color de fons.
   * **Color**: Introdueix un color en format hexadecimal (p. ex., `#FFFFFF` per al blanc) o fes clic a `Tria color` per seleccionar-lo visualment.

4. **Processar**:
   
   * Fes clic al botó **Processar PDF** per iniciar la creació del nou document. El programa et demanarà on vols guardar el fitxer PDF resultant.

5. **Progrés i Registre**:
   
   * Una barra de progrés mostrarà l'avanç del procés.
   * Una caixa de text a la part inferior registrarà les accions realitzades i els possibles errors.

### Execució del Programa

Per executar el programa, no necessites instal·lar Python ni cap llibreria. Pots fer servir directament els executables proporcionats.

#### A Linux:

1. Obre un terminal.
2. Navega fins a la carpeta `Editor_PDF_cap/Linux`.
3. Dóna permisos d'execució a l'arxiu si és necessari amb la comanda:
   
   ```bash
   chmod +x EditaPdf2
   ```
4. Executa el programa amb:
   
   ```bash
   ./EditaPdf2
   ```

#### A Windows:

1. Navega fins a la carpeta `Editor_PDF_cap\windows`.
2. Fes doble clic sobre l'arxiu `EditaPdf2.exe`.
