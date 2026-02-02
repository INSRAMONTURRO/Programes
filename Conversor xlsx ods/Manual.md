# Manual d'Usuari: Conversor de Fitxers (ODS <> XLSX)

Aquesta és una aplicació d'escriptori senzilla que permet convertir fitxers de full de càlcul entre els formats `.ods` (Open Document Spreadsheet) i `.xlsx` (Microsoft Excel).

## Funcionalitats

- Converteix fitxers en massa des d'una carpeta d'origen a una de destí.
- Suporta dues direccions de conversió:
  - De `.ods` a `.xlsx`
  - De `.xlsx` a `.ods`
- Proporciona un registre en temps real del progrés de la conversió.
- Interfície gràfica intuïtiva i fàcil d'utilitzar.

## Com Funciona

L'aplicació utilitza la línia de comandes de **LibreOffice** per realitzar les conversions. Per tant, és un requisit indispensable tenir LibreOffice instal·lat al sistema.

El procés que segueix l'aplicació és el següent:

1. **Selecció de carpetes**: L'usuari tria una carpeta que conté els fitxers a convertir (carpeta d'origen) i una altra on es desaran els fitxers convertits (carpeta de destí).
2. **Selecció del mode de conversió**: L'usuari indica si vol convertir d'ODS a XLSX o viceversa.
3. **Inici de la conversió**: En prémer el botó "Converteix!", l'aplicació cerca tots els fitxers amb l'extensió corresponent a la carpeta d'origen.
4. **Execució en segon pla**: Per cada fitxer trobat, l'aplicació executa una comanda de LibreOffice en segon pla (`--headless`) que realitza la conversió i desa el resultat a la carpeta de destí.
5. **Registre del progrés**: Cada pas de la conversió (fitxer trobat, inici de conversió, finalització o error) es mostra al quadre de text de progrés.
6. **Finalització**: Un cop s'han processat tots els fitxers, es mostra un missatge informatiu.

## Autor i Llicència

- **Autor**: Josep Maria
- **Llicència**: Creative Commons BY-NC-SA 4.0.
  Pots consultar els detalls de la llicència fent clic a l'enllaç que apareix al peu de pàgina de l'aplicació.
