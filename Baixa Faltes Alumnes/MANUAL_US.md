# Manual d'Ús: Gestor de Faltes (BaixaFaltes)

Aquest document explica com utilitzar l'aplicació "Gestor de Faltes" per automatitzar la descàrrega, processament i anàlisi de les faltes d'assistència de l'alumnat.

---

## 1. Requisits Previs

Abans d'executar l'aplicació, assegura't de tenir el següent:

1. **Google Chrome:** El navegador ha d'estar instal·lat a l'ordinador.
2. **Fitxer `dades.xlsx`:** Aquest és el fitxer mestre i ha de contenir dues fulles:
   * **`Alumnes`**: Amb les dades de tot l'alumnat (nom, curs, contactes, etc.).
   * **`Tutors`**: Amb la llista de tutors i coordinadors. Ha de contenir una columna on s'identifiqui el rol (p. ex., "Coordinador ESO") i la columna del costat amb el correu electrònic.
3. **Fitxer `Expedients2526.ods` (Opcional):** Si aquest fitxer existeix, el programa filtrarà les faltes dels alumnes amb expedients oberts, considerant només les faltes posteriors a la data de l'última sanció.

Aquests fitxers han d'estar dins de la mateixa **carpeta de treball** que seleccionaràs a l'aplicació.

---

## 2. Configuració de l'Aplicació

En obrir l'aplicació, veuràs una finestra amb diverses seccions que cal configurar:

![Interfície del Gestor de Faltes](https://i.imgur.com/URL_DE_LA_IMATGE.png)  <!-- Afegeix aquí una captura de pantalla si és possible -->

1. **Configuració Enviament:**
   
   * **Email:** Introdueix l'adreça de correu de Gmail des de la qual s'enviaran els resums (p. ex., `cap.estudis@iesmalgrat.cat`).
   * **Pwd App:** Introdueix la **contrasenya d'aplicació** de 16 caràcters generada per a aquest compte de Gmail. **No és la teva contrasenya habitual.**

2. **Configuració Chrome:**
   
   * **Versió de Chrome:** Assegura't que la versió de Chrome que apareix aquí coincideix amb la que tens instal·lada. Normalment, no cal canviar-ho.

3. **Rang de Dates:**
   
   * **Inici / Fi:** Selecciona el període de dates per al qual vols descarregar el report d'assistència.

4. **Carpeta de Treball:**
   
   * Fes clic a **`📂 Triar...`** i selecciona la carpeta on tens els fitxers `dades.xlsx` i `Expedients2526.ods`. Aquesta carpeta també serà on es desaran els informes generats.

---

## 3. Passos d'Execució

El procés es divideix en dos o tres passos simples, guiats pels botons numerats.

### Pas 1: OBRIR CHROME

* Fes clic al botó **`1. OBRIR CHROME`**.
* S'obrirà una finestra del navegador Chrome.
* **Important:** Inicia sessió manualment a la plataforma iEduca amb el teu usuari i contrasenya. Deixa la finestra oberta un cop hagis iniciat sessió.

### Pas 2: BAIXAR AUTOMÀTIC

* Un cop has iniciat sessió a iEduca, fes clic al botó **`2. BAIXAR AUTOMÀTIC`**.
* El programa realitzarà automàticament totes les tasques següents:
  1. Descarregarà el fitxer Excel amb les faltes del període seleccionat.
  2. Processarà les dades, creuant-les amb els fitxers `dades.xlsx` i `Expedients2526.ods`.
  3. Generarà el fitxer **`RESUM_GLOBAL.xlsx`** amb dues pestanyes: "Resum General" i "Casos Greus".
  4. Crearà una carpeta anomenada **`informes/`** amb informes individuals en format Markdown per a cada curs de la llista de "Casos Greus".
  5. Enviarà un correu electrònic als coordinadors amb el fitxer `RESUM_GLOBAL.xlsx` adjunt.
* Pots seguir el progrés de totes aquestes accions a la consola de text de l'aplicació.

### Alternativa: ANALITZAR LOCAL

* Si has descarregat manualment un fitxer de faltes, pots analitzar-lo sense necessitat d'obrir Chrome.
* Fes clic a **`3. ANALITZAR LOCAL`**, selecciona el fitxer Excel que has descarregat i el programa el processarà de la mateixa manera.

---

## 4. Fitxers Generats

Un cop finalitzat el procés, trobaràs els següents fitxers a la teva carpeta de treball:

* **`RESUM_GLOBAL.xlsx`:**
  * **`Resum General`**: Llista de tots els alumnes amb incidències de nivell 3 o 4, ordenats per prioritat. Les files estan acolorides per identificar ràpidament els casos més urgents.
  * **`Casos Greus`**: Un subconjunt del resum general, mostrant només els alumnes amb un nombre més elevat d'incidències greus.
* **Carpeta `informes/`:**
  * Conté fitxers `.md` per a cada curs amb alumnes considerats greus. Aquests fitxers detallen les faltes de Nivell 3 i 4, incloent data, professor i observacions, llestos per a la seva consulta o impressió.

---

## 5. Solució de Problemes Comuns

* **El correu no s'envia:**
  * Verifica que l'email i la **contrasenya d'aplicació** siguin correctes.
  * Assegura't que la fulla "Tutors" del `dades.xlsx` té una columna amb la paraula "COORDINADOR" i que la columna del costat conté els correus.
* **Error en obrir Chrome:**
  * Comprova que la versió de Chrome a l'aplicació és la correcta.
  * Assegura't que Google Chrome està instal·lat.
* **Error "No s'ha trobat el fitxer..."**:
  * Verifica que el fitxer `dades.xlsx` es troba a la "Carpeta de Treball" que has seleccionat.
