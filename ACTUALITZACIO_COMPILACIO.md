# Instruccions per Actualitzar la Compilació Automàtica (GitHub Actions)

Aquest document resumeix els passos necessaris per pujar els canvis al teu repositori de GitHub i activar el flux de treball de GitHub Actions, que compila els programes i genera les "releases" (versions) amb totes les dependències incloses.

## Passos a Seguir

### 1. Preparar els Canvis (Stage)

Aquesta comanda afegeix tots els fitxers modificats, creats o eliminats a la "zona de preparació" de Git, deixant-los a punt per ser inclosos en el següent commit.

```bash
git add .
```

### 2. Desar els Canvis Localment (Commit)

Crea un "commit" amb tots els canvis preparats. És important utilitzar un missatge clar i descriptiu que resumeixi què s'ha fet.

```bash
git commit -m "CI: Integra la instal·lació de dependències a la compilació i estandarditza requeriments"
```

### 3. Crear una Etiqueta de Versió (Tag)

El flux de treball de GitHub Actions (`.github/workflows/build.yml`) s'activa específicament quan es puja una etiqueta (tag) que comença per `v`. Has de crear una nova etiqueta per a aquesta versió.

**Abans d'executar la comanda:**
*   **Tria el número de versió:** Decideix quin serà el següent número de versió (per exemple, `v1.0.0`, `v1.0.1`, `v2.0.0`, etc.). Si no hi ha versions anteriors, `v1.0.0` és un bon punt de partida.
*   **Afegeix un missatge:** El missatge de l'etiqueta (`-m`) ha de descriure breument el contingut d'aquesta versió.

```bash
git tag -a vX.Y.Z -m "Release vX.Y.Z: Actualització de la compilació amb gestió de dependències"
```
*(Substitueix `vX.Y.Z` pel teu número de versió real, per exemple `v1.0.0`.)*

### 4. Pujar els Canvis i l'Etiqueta a GitHub (Push)

Finalment, aquesta comanda envia tant els teus commits com l'etiqueta que has creat al repositori remot de GitHub. Això desencadenarà l'execució del flux de treball de GitHub Actions.

```bash
git push --tags
```
*(Aquesta comanda pujarà la branca actual al mateix temps que totes les etiquetes.)*

---

Un cop executats aquests passos, podràs veure el progrés de la compilació a la secció "Actions" del teu repositori a GitHub, i les noves versions estaran disponibles a la secció "Releases" un cop finalitzi el procés.
