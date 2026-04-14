# Manual: Quadre de Disponibilitat (QuadreDisponibilitat_v03.py)

### **1. Què necessita el programa?**

Igual que l'anterior, utilitza el fitxer d'exportació de dades del centre.

* **Origen**: Exportació de **GP-Untis**.
* **Format**: Fitxer de text (.txt) o CSV amb la graella d'horaris completa del professorat i grups.

### **2. Què fa el programa?**

Aquest programa crea una **matriu de disponibilitat setmanal**. És una eina visual per decidir quin dia és millor fer una sortida o per veure quins professors coincideixen amb uns grups determinats durant tota la setmana.

* Mostra en una taula els professors que tenen classe amb els grups seleccionats.
* Permet veure la càrrega lectiva de diversos dies i hores simultàniament.

### **3. Com funciona?**

1. **Carrega l'horari**: Prem el botó **"Carregar Horaris.TXT"** i tria el fitxer de GP-Untis.
2. **Defineix la graella**:
   * **Dies**: Selecciona un o diversos dies (pots marcar-ne tota la setmana).
   * **Hores**: Tria les franges horàries que t'interessen.
   * **Grups**: Selecciona els grups (ex: tots els de 4t d'ESO).
3. **Visualitza el quadre**: Prem **"GENERAR QUADRE v03"**. Apareixerà una taula on:
   * Les **columnes** són els dies de la setmana.
   * Les **files** són les hores.
   * A cada cel·la hi haurà la llista de professors que tenen classe amb aquells grups en aquell moment.
4. **Exporta la informació**: Prem **"GUARDAR RESULTATS"** per desar la taula en:
   * **Markdown (.md)**: Perfecte per enganxar-lo en un document de text o acta de reunió mantenint el format de taula.
   * **CSV (.csv)**: Per obrir-lo amb Excel o Google Sheets i fer gestions posteriors.

---

**Autoria**: Josep M. | **Llicència**: CC BY-NC-SA 4.0
