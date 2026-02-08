# Manual d'Usuari: Analitzador d'Assistència Complet

## 1. Introducció

L'**Analitzador d'Assistència Complet** és una aplicació d'escriptori dissenyada per facilitar l'anàlisi de fitxers d'assistència del professorat en format Excel. El programa unifica diverses funcionalitats per processar les dades i generar informes clars i útils.

Les seves **característiques principals** són:

- **Anàlisi de Retards**: Identifica i suma els minuts de retard.
- **Anàlisi de Fitxatges**: Detecta i compta les vegades que no s'ha fitxat a l'entrada o a la sortida.
- **Generació d'Informes Flexibles**: Crea dos tipus d'informes:
  1. Un **resum en format Excel** amb el total de retards i faltes de fitxatge per professor.
  2. **Informes detallats en format Markdown** (.md) amb la llista específica d'incidències per als professors que superen certs llindars.
- **Interfície Gràfica Intuïtiva**: Totes les opcions són configurables a través d'una finestra fàcil d'utilitzar.
- **Organització de Fitxers**: Desa tots els informes generats en una carpeta específica per mantenir l'ordre.

## 2. Requisits per al Funcionament

Perquè el programa funcioni correctament, cal complir els següents requisits:

### A. Programari

1. **Python**: Cal tenir instal·lat Python (versió 3.6 o superior). Pots descarregar-lo des de [python.org](https://www.python.org/).

2. **Biblioteques de Python**: El programa depèn de dues biblioteques externes: `pandas` i `openpyxl`. Per instal·lar-les, obre un terminal o línia de comandes, navega fins a la carpeta on es troba el programa i executa la següent comanda:
   
   ```bash
   pip install -r requirements.txt
   ```
   
   Això instal·larà automàticament les versions necessàries.

### B. Format del Fitxer Excel d'Entrada

El fitxer Excel que s'utilitzi com a entrada ha de tenir una estructura concreta:

- Ha de contenir una o més pestanyes (fulls de càlcul) on les dades a analitzar comencin pel nom **"Tractament"** (p. ex., "Tractament_Setmana1", "Tractament_2026_02").
- Aquestes pestanyes han de contenir, com a mínim, les següents columnes:
  - `Professor`: Nom del professor/a.
  - `Data`: La data de la incidència.
  - `Minuts Diferència`: Un valor numèric que representa els minuts de retard.
  - `Tipus Incident`: Un text que descriu la incidència (p. ex., "Entrada no fitxada").

## 3. Com Executar el Programa

Per iniciar l'aplicació, segueix aquests passos:

1. Obre un terminal o símbol del sistema (cmd).

2. Navega fins a la carpeta on has desat el fitxer `AnalitzadorAssistenciaComplet.py`.

3. Executa la següent comanda:
   
   ```bash
   python AnalitzadorAssistenciaComplet.py
   ```

4. S'obrirà la finestra principal de l'aplicació.

## 4. Ús de la Interfície Gràfica

La interfície està dividida en seccions numerades per guiar l'usuari a través del procés.

![Imatge de la interfície d'usuari](https://i.imgur.com/link-a-la-imatge.png) *(Nota: Aquesta és una representació. La imatge real no està inclosa en aquest Markdown)*

### Pas 1: Seleccionar el Fitxer d'Entrada

- Fes clic al botó **"Selecciona Fitxer Excel"**.
- S'obrirà una finestra per navegar pels teus arxius. Tria el fitxer Excel que vols analitzar.
- Un cop seleccionat, el nom del fitxer apareixerà a l'etiqueta.

### Pas 2: Configurar l'Anàlisi

En aquesta secció pots ajustar els llindars que el programa utilitzarà per als seus càlculs:

- **Llindar resum retards (minuts)**: Per a l'informe Excel, només es comptaran com a retard les incidències amb menys minuts que aquest valor.
- **Llindar detall retards (minuts acumulats)**: A l'informe detallat de retards, només apareixeran els professors la suma total de minuts de retard dels quals superi aquest llindar.
- **Límit superior per retard individual (detall)**: Defineix què es considera un retard. Per exemple, si el límit és 25, una incidència de 40 minuts no es comptarà com a retard acumulat (es podria considerar una absència parcial).
- **Llindar detall fitxatges (núm. incidents)**: A l'informe detallat de fitxatges, només apareixeran els professors amb un nombre total d'incidents superior a aquest valor.

### Pas 3: Indicar la Destinació de Sortida

- **Nom de la carpeta**: Escriu el nom de la carpeta on vols que es guardin els informes generats. Per defecte, és "Informes_Generats". Si la carpeta no existeix, el programa la crearà.

### Pas 4: Triar els Tipus i Noms dels Informes

Pots decidir quins informes vols generar:

- **Generar Resum en Excel**: Marca aquesta casella per crear el fitxer `.xlsx`. Pots personalitzar el nom del fitxer al camp de text adjacent.
- **Generar Informes Detallats (Markdown)**: Marca aquesta casella per crear els fitxers `.md`. Pots personalitzar els noms dels fitxers de retards i de fitxatges.

### Pas 5: Seleccionar les Pestanyes

- Un cop hagis seleccionat un fitxer Excel, en aquesta àrea apareixerà una llista de totes les pestanyes disponibles que comencen amb "Tractament".
- **Marca les caselles** de les pestanyes que vols incloure a l'anàlisi. Pots seleccionar-ne tantes com vulguis.

### Pas 6: Processar i Tancar

- **Processar Anàlisi**: Un cop hagis configurat tot, fes clic en aquest botó per iniciar el procés. El programa llegirà les dades, farà els càlculs i generarà els fitxers. Al final, mostrarà un missatge indicant si l'operació ha tingut èxit i quins fitxers s'han creat.
- **Tancar**: Fes clic per tancar l'aplicació.

### Peu de Pàgina (Footer)

A la part inferior de la finestra, trobaràs la informació de l'autor i la llicència del programa. Si fas clic sobre el text de la llicència, s'obrirà la pàgina web oficial de la llicència en el teu navegador.

## 5. Fitxers de Sortida

Els informes generats es desaran a la carpeta que hagis especificat.

- **Fitxer Excel (`.xlsx`)**: Conté dues pestanyes:
  - `Resum_Retards`: Una taula amb el recompte de retards i la suma total de minuts per professor.
  - `Resum_Faltes_Fitxatge`: Una taula amb el recompte de faltes de fitxatge a l'entrada i a la sortida per professor.
- **Fitxers Markdown (`.md`)**:
  - `..._detall_retards.md`: Un informe de text detallat amb cada un dels retards dels professors que han superat el llindar de minuts acumulats.
  - `..._detall_fitxatges.md`: Un informe de text detallat amb cada una de les incidències de fitxatge dels professors que han superat el llindar d'incidents.

## 6. Solució de Problemes Comuns

- **"No s'han trobat pestanyes..."**: Assegura't que el nom de les pestanyes que vols analitzar al teu fitxer Excel comença amb la paraula `Tractament`.
- **"Falten columnes..."**: Revisa el teu fitxer Excel. Ha de contenir les columnes `Professor`, `Data`, `Minuts Diferència` i `Tipus Incident`.
- **"Sense Dades"**: Pot ser que les pestanyes que has seleccionat estiguin buides.
- **Errors de permisos**: Assegura't que tens permisos per crear carpetes i escriure fitxers a la ubicació on estàs executant el programa.
