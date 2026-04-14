# Manual: Rànquing de Professors (RankingProfes_v03.py)

### **1. Què necessita el programa?**

El programa requereix el fitxer d'exportació d'horaris del centre.

* **Origen**: Exportació de **GP-Untis**.
* **Format**: Fitxer de text (.txt) o CSV on les columnes estiguin separades per comes.
* **Dades clau**: El fitxer ha de contenir informació sobre el Grup, el Professor, el Dia de la setmana (numèric 1-5) i l'Hora (numèrica 1-7).

### **2. Què fa el programa?**

L'objectiu d'aquest script és ajudar a triar els acompanyants més "barats" (en termes d'afectació horària) per a una sortida escolar. 

* **Analitza** quins professors tenen classe amb els grups que marxen (aquestes hores no es perden, ja que el professor "seguiria" amb el seu grup).
* **Detecta** hores lliures, guàrdies o reunions.
* **Genera un rànquing** d'idoneïtat on els primers de la llista són els que menys classes amb altres grups perden.

### **3. Com funciona?**

1. **Carrega l'horari**: Prem el botó **"1. Carregar Horaris.TXT"** i selecciona l'exportació de GP-Untis.
2. **Configura la sortida**:
   * Tria el **Dia** de la setmana.
   * Selecciona les **Hores** (1a, 2a, etc.). Pots fer doble clic per seleccionar-les totes de cop.
   * Marca el **Grup o Grups** que fan la sortida de la llista que apareixerà a la dreta.
3. **Genera l'anàlisi**: Prem **"GENERAR LLISTAT D'IDONEÏTAT"**.
4. **Consulta els resultats**: 
   * A la part superior veuràs el **Top 10** de professors recomanats.
   * A la part inferior veuràs el detall de què fa cada professor en cada hora de la sortida.

---

**Autoria**: Josep M. | **Llicència**: CC BY-NC-SA 4.0
