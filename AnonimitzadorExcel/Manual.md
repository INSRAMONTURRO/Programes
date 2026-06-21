# Manual d'Ús: Gestor d'Anonimització iEduca (AnonimitzadorExcel.py)

Aquest programa és una aplicació interactiva amb entorn gràfic (GUI) en Python dissenyada per anonimitzar i desanonimitzar columnes específiques (com ara professors, alumnes, etc.) en un fitxer de dades Excel, garantint la protecció de dades personals de manera coherent.

---

## 🚀 Què fa el programa?

L'aplicació permet seleccionar un fitxer Excel, definir les columnes a anonimitzar, configurar-hi prefixos independents per a cadascuna d'elles i aplicar els canvis a **tots els fulls** del fitxer:

1. **Anonimització (Noms Reals ➔ Pseudònims)**:
   
   * Reemplaça els noms reals de les columnes seleccionades per pseudònims com `Alumne 001`, `Professor 001`, etc.
   * Els pseudònims es generen de manera incremental i es desen a una base de dades local (`mapeig_anonim.json`). Això garanteix que una mateixa persona tingui **sempre el mateix pseudònim** a tots els fulls i en futures execucucions.
   * Permet especificar una "fila límit" mitjançant una paraula clau (ex: `mitjanes`). Si es troba aquesta paraula clau, s'esborraran totes les files posteriors per netejar el document.

2. **Desanonimització (Pseudònims ➔ Noms Reals)**:
   
   * Realitza l'operació inversa. Utilitza la base de dades de mapeigs (`mapeig_anonim.json`) per restaurar els noms reals a partir dels pseudònims en cas de necessitar consultar la informació original.

---

## 📦 Requisits del sistema

Per poder executar l'aplicació, es necessari tenir instal·lat Python 3 i les següents llibreries:

```bash
pip install pandas openpyxl
```

*Nota: La llibreria de gràfics `tkinter` ve instal·lada per defecte amb la majoria de distribucions de Python. En sistemes Linux (com Ubuntu/Debian), si no la tens, pots instal·lar-la amb:*

```bash
sudo apt install python3-tk
```

---

## 🛠️ Com s'utilitza?

1. **Inicia l'aplicació**:
   Executa el script des del teu terminal dins de la carpeta del projecte:
   
   ```bash
   python3 AnonimitzadorExcel.py
   ```

2. **Selecciona l'operació**:
   
   * Tria **Anonimitzar** si vols amagar els noms.
   * Tria **Desanonimitzar** si vols restaurar els noms reals a partir de la base de dades de mapeigs.

3. **Tria el fitxer Excel**:
   Fes clic a **Tria fitxer...** i selecciona l'arxiu Excel que vols tractar. El programa desarà la sortida a la mateixa carpeta amb el prefix `ANON_` o `REAL_` (per exemple, `ANON_Visites-WC.xlsx`).

4. **Configura les columnes i prefixos**:
   
   * **Columnes a anonimitzar**: Escriu les columnes que vols processar separades per comes (per defecte: `B, D`).
     * *Exemple: Columna B (Professors) i Columna D (Alumnes).*
   * **Prefixos de cada columna**: Escriu els prefixos dels pseudònims en el mateix ordre separats per comes (per defecte: `Professor, Alumne`).
     * La columna B usarà el prefix `Professor` (`Professor 001`, `Professor 002`...).
     * La columna D usarà el prefix `Alumne` (`Alumne 001`, `Alumne 002`...).
     * *Si ometies un prefix, el programa utilitzarà `Alumne` per defecte.*
   * **Fila límit (Paraula clau)**: Si vols esborrar files de resum/mitjanes al final dels fulls, escriu la paraula clau (ex: `mitjanes`). Si no en vols cap, deixa el camp buit.

5. **Execució**:
   
   * Fes clic a **EXECUTAR PROCESSAMENT A TOTS ELS FULLS**.
   * Pots seguir el progrés en el quadre de text blau de la part inferior.
   * **Netejar Mapeig (Opcional)**: El botó vermell **NETEJAR MAPEIG DE NOMS** serveix per esborrar l'historial de pseudònims (`mapeig_anonim.json`) si vols començar de nou la numeració des de zero.

---

## 👥 Autoria i Llicència

* **Autor**: Josepm
* **Llicència**: Creative Commons BY-NC-SA 4.0 (Reconeixement-NoComercial-CompartirIgual)
* L'aplicació conté un enllaç interactiu al peu de pàgina de la finestra que permet obrir les condicions d'ús de la llicència al teu navegador.
