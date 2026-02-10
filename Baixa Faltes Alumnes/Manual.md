# Manual d'Usuari: Gestor de Faltes (V51)

Aquest document explica com instal·lar, configurar i utilitzar el programa Gestor de Faltes.

## 1. Visió General

El Gestor de Faltes és una eina dissenyada per automatitzar la descàrrega, processament i anàlisi de les faltes d'assistència de l'alumnat des de la plataforma ieduca. L'aplicació genera informes individuals i un resum global per facilitar el seguiment per part de la prefectura d'estudis i els coordinadors.

## 2. Instal·lació

Per poder utilitzar aquest programa, necessites tenir **Python** instal·lat al teu ordinador. Si no el tens, pots descarregar-lo des de [python.org](https://www.python.org/).

Un cop tinguis Python, segueix aquests passos:

### a. Descarrega els Fitxers del Projecte

Descarrega tots els fitxers d'aquest repositori i desa'ls en una carpeta al teu ordinador.

### b. Instal·la les Dependències

Obre una terminal o símbol del sistema, navega fins a la carpeta on has desat els fitxers i executa la següent comanda per instal·lar totes les llibreries necessàries:

```bash
pip install -r requirements.txt
```

Això instal·larà `pandas`, `openpyxl`, `undetected-chromedriver` i altres llibreries necessàries.

## 3. Configuració

Abans d'executar el programa per primer cop, has de configurar les teves credencials de correu electrònic per permetre l'enviament d'informes.

### a. Crea el fitxer `.env`

A la carpeta del projecte, trobaràs un fitxer anomenat `.env.example`. Fes una còpia d'aquest fitxer i anomena-la `.env`.

### b. Edita el fitxer `.env`

Obre el nou fitxer `.env` amb un editor de text. Veuràs el següent:

```
EMAIL_ORIGEN="el_teu_correu@gmail.com"
EMAIL_PASSWORD="la_teva_contrasenya_d_aplicacio"
```

- **EMAIL_ORIGEN**: Substitueix `"el_teu_correu@gmail.com"` pel teu compte de correu de Gmail des del qual vols enviar els informes.
- **EMAIL_PASSWORD**: Aquí has d'introduir una **Contrasenya d'Aplicació**, no la teva contrasenya habitual de Gmail. Segueix els passos de la següent secció per obtenir-ne una.

### c. Com Obtenir una Contrasenya d'Aplicació de Google

Per motius de seguretat, el programa no utilitza la teva contrasenya principal de Gmail. En el seu lloc, necessita una "Contrasenya d'Aplicació" que dones a l'eina per accedir només a la funcionalitat d'enviar correus.

**Requisit previ:** Has de tenir activada la **Verificació en dos passos** al teu compte de Google.

1. **Ves al teu Compte de Google**: Accedeix a [myaccount.google.com](https://myaccount.google.com/).

2. **Secció "Seguretat"**: Al menú de l'esquerra, fes clic a **Seguretat**.

3. **Inici de sessió a Google**: Busca la secció "Com inicies la sessió a Google" i fes clic a **Contrasenyes d'aplicacions** (o "App Passwords"). Si no veus aquesta opció, és probable que no tinguis la verificació en dos passos activada.

4. **Genera la Contrasenya**:
   
   - A "Selecciona l'aplicació", tria **Altra (*nom personalitzat*)**.
   - Escriu un nom descriptiu, com ara `GestorFaltesPython`.
   - Fes clic al botó **Generar**.

5. **Copia la Contrasenya**: Google et mostrarà una contrasenya de **16 caràcters** sobre un fons groc.
   
   ![Exemple de contrasenya d'aplicació](https://i.imgur.com/Eyw4F9H.png)
   
   **Copia aquesta contrasenya de 16 lletres (sense els espais)** i enganxa-la al teu fitxer `.env` a la variable `EMAIL_PASSWORD`.
   
   ```
   EMAIL_PASSWORD="xxxx"  <-- Enganxa-la aquí
   ```

Un cop desat el fitxer `.env`, el programa ja està llest per ser utilitzat.

## 4. Execució del Programa

Per iniciar l'aplicació, obre una terminal, navega a la carpeta del projecte i executa:

```bash
python BaixaFaltes51.py
```

S'obrirà la interfície gràfica del programa i ja podràs començar a treballar.
