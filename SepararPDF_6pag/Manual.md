# Manual d'Usuari: Separador d'Informes d'Alumnes

Aquest document explica com configurar i utilitzar el programa Separador d'Informes d'Alumnes per dividir fitxers PDF combinats en informes individuals personalitzats.

## 1. Visió General

El Separador d'Informes d'Alumnes és una eina que permet agafar un fitxer PDF gran on hi ha agrupats els informes de molts alumnes (per exemple, butlletins de notes, informes d'orientació, etc.) i extreure'ls en fitxers PDF individuals de forma automàtica. 

Cada fitxer de sortida s'anomena de manera personalitzada amb el nom, cognom i codi RALC de l'alumne, fent servir les dades importades d'un llistat CSV.

---

## 2. Requisits dels Fitxers d'Entrada

Perquè el programa funcioni correctament, has de preparar dos fitxers:

### a. Llistat d'Alumnes (CSV)
Un fitxer de text amb format CSV que conté les dades de l'alumnat.
* **Separador:** Pot estar separat per comes (`,`) o punt i comes (`;`).
* **Codificació:** Suporta codificacions habituals (`UTF-8`, `Latin-1`, `Windows-1252`).
* **Columnes obligatòries (amb capçalera exactament així o començant per):**
  * `NOM`: Nom de l'alumne.
  * `COGNOM1`: Primer cognom de l'alumne.
  * `RALC`: Codi d'identificació de l'alumne (el número sencer de 10-11 xifres).

*Exemple de contingut del CSV:*
```csv
CURS 2025-2026,NOM,COGNOM1,COGNOM2,RALC
2 ESO C,Xavier,Mates,Pla,1234567890
2 ESO C,Sofia,Tordera,Malgrat,1987654321
```

### b. Fitxer PDF d'Informes Agrupats
Un únic fitxer PDF que conté tots els informes un rere l'altre, en el mateix ordre que els alumnes apareixen al llistat CSV.
* **Mida de cada informe:** Configurable directament des de la pantalla de l'aplicació (per defecte 6 pàgines per alumne).

---

## 3. Instruccions d'Ús pas a pas

1. **Executa el programa** fent doble clic a l'executable o obrint-lo des del terminal.
2. **Selecciona el llistat d'alumnes:** Fes clic al botó **"Selecciona CSV..."** i tria el teu fitxer CSV preparat.
3. **Selecciona el document PDF:** Fes clic al botó **"Selecciona PDF..."** i tria el fitxer PDF combinat que vols dividir.
4. **Configura les pàgines:** Al camp **"Nombre de pàgines per informe"**, utilitza les fletxes o escriu el nombre exacte de pàgines que ocupa l'informe de cada alumne.
5. **Comprova les dades detectades:** Revisa la targeta de resum a la pantalla per comprovar si el nombre d'alumnes i el d'informes detectats coincideix correctament.
6. **Carpeta de sortida:** El programa proposarà automàticament desar els fitxers individuals en una nova carpeta anomenada `informes_separats_[NomPDF]` al mateix directori de l'original. Si vols canviar-ho, prem **"Canvia carpeta..."**.
7. **Executa la divisió:** Clica el botó verd **"Comença la Divisió dels Informes"**. Veuràs el progrés en temps real.
8. **Finalització:** Quan el procés hagi acabat, es mostrarà un missatge de confirmació i apareixerà un botó blau anomenat **"Obrir carpeta de sortida"** per accedir immediatament als PDF generats.

---

## 4. Nomenclatura del Fitxer de Sortida

Cada fitxer individual generat es guardarà amb el format:
`AD_[Nom]_[Cognom]_[RALC].pdf`

*Exemple:* `AD_Xavier_Mates_1234567890.pdf`

---

## 5. Llicència i Autoria

* **Autor:** Josep M Sardà
* **Llicència:** Creative Commons Reconeixement-NoComercial-CompartirIgual 4.0 Internacional (CC BY-NC-SA 4.0).
