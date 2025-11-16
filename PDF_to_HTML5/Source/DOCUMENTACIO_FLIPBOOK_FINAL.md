# PDF to HTML5 Flipbook Converter - Documentació Tècnica

## Descripció
Aquest programa converteix fitxers PDF en llibres HTML5 interactius utilitzant plantilles professionals. Permet triar entre dos modes de visualització: vertical (efecte de llibre/revista amb dues pàgines) i horitzontal (fulla única amb una pàgina).

## Característiques principals
- Interfície gràfica fàcil d'utilitzar
- Conversió de PDF a imatges JPEG de pàgines
- Generació de miniatures
- Dos modes de visualització:
  - Mode vertical: Efecte de llibre amb voltes de pàgina realistes
  - Mode horitzontal: Visualització simple d'una pàgina alhora
- Configuració de resolució (DPI ajustable entre 72-600)
- Empaquetat en fitxer ZIP

## Estructura de sortida (comuna per a ambdós modes)
El fitxer ZIP generat té la següent estructura:

```
flipbook.zip
├── index.html (fitxer original sense modificacions)
├── book/
│   ├── index.html (fitxer personalitzat amb títol i pàgines)
│   ├── pages/ 
│   │   ├── 1.jpg, 2.jpg, 3.jpg... (imatges de pàgines)
│   │   └── 1-thumb.jpg, 2-thumb.jpg... (miniatures)
│   ├── js/
│   │   └── magazine.js
│   ├── css/
│   │   └── magazine.css
│   └── pics/ (imatges de fons i botons)
├── extras/ (llibreries jQuery, etc.)
└── lib/ (llibreries turn.js, zoom.js, etc.)
```

## Requisits del sistema
- Python 3.6 o superior
- Llibreries necessàries:
  - PyMuPDF (fitz)
  - Pillow (PIL)
  - tkinter

## Funcionament
1. Selecció del fitxer PDF
2. Triar el mode de visualització (vertical o horitzontal)
3. Ajustar la resolució (opcional)
4. Executar la conversió
5. El programa genera un fitxer ZIP amb la estructura adequada

## Modes de visualització
### Mode vertical (Llibre/Revista)
- Utilitza la llibreria turn.js per l'efecte de voltes de pàgina
- Mostra dues pàgines alhora
- Experiència semblant a llegir un llibre real

### Mode horitzontal (Fulla Única)
- Mostra una sola pàgina alhora
- Navegació amb botons d'anterior/següent
- Interfície més simple i directa

## Personalització
El títol del llibre es substitueix automàticament a la variable `_CADENA_PER_CANVIAR_`
El número de pàgines es substitueix a la variable `_N_PAGES_`

## Plantilles
- `templates.tar.gz`: Plantilla per al mode vertical
- `templates_H.tar.gz`: Plantilla per al mode horitzontal