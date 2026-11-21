# 🧭 Welcome Kit — Induction Project

Proyecto final de inglés: **Welcome Kit for a new foreign colleague** (4 diapositivas, guion oral de
4–5 minutos y parte escrita de 250–350 palabras), presentado con un diseño propio en **PowerPoint** y
en **web**, más su PDF.

> **Davis Fernández** · técnico independiente de IT · David, Chiriquí · octubre 2026

## 🌐 Ver la presentación (en vivo)

**https://davisday13.github.io/welcome-kit-induction/**

| Tecla | Acción |
|---|---|
| `←` `→` / clic / swipe | Navegar entre diapositivas |
| `1`–`4` | Ir directo a una diapositiva |
| `N` | Abrir/cerrar el **guion oral** (speaker notes) |
| `F` | Pantalla completa |
| `P` | Exportar a PDF (imprimir → *Guardar como PDF*) |

## 🎨 Diseño

Estilo **menta / Canva**: fondo menta claro `#E9FBF5`, títulos en azul marino `#0E1B4D`, acentos en
turquesa `#12C2A0`, amarillo `#FFC53D`, coral `#FF6B8A` y azul `#3D6BFF`. Tarjetas blancas con
esquinas redondeadas y sombra suave, chips-píldora para los rótulos, ilustración plana del técnico
en un círculo amarillo y una franja de cuatro colores en la portada. El mismo lenguaje visual se usa
en el PPTX, la web y el PDF.

| Diapositiva | Recursos gráficos |
|---|---|
| 1 · Portada | Círculo amarillo con ilustración, chip *New team member*, título grande navy y franja de 4 colores |
| 2 · About me | Chips de personalidad, 3 tarjetas de rutina, tarjeta navy *This week* e ilustración a la derecha |
| 3 · How to get there | Tarjeta navy del cliente, pasos numerados en turquesa y mapa real en tarjeta blanca |
| 4 · Office & safety | Filas de comida con precios, chips *Yesterday* / *Tomorrow* y 3 tarjetas de seguridad |

Todas las decisiones de color mantienen buen contraste: navy sobre blanco/menta y blanco sobre navy.

## 🎬 Transiciones y animaciones

| Momento | Qué pasa |
|---|---|
| **Cambiar de diapositiva** | Barrido turquesa con borde amarillo que cruza el escenario (340 ms); la diapositiva anterior sale hacia el lado por el que viajas y la nueva entra |
| **Al llegar a cada diapositiva** | Entrada escalonada de bloques: *chip* → título → tarjetas una a una → bloque inferior (60–80 ms de retraso, curva `cubic-bezier`) |
| **Portada** | Ilustración con zoom, título y datos apareciendo en orden, chip inferior y franja de color deslizándose |
| **Mapa (diapositiva 3)** | **Captura real de Google Maps** con la ruta en azul de La Riviera a PikTa, el pin de salida y el de llegada, y la etiqueta *10 min · 5 km* |
| **Detalles** | Puntos de navegación redondos que crecen, cajón de notas que sube, pista que se oculta sola |
| **Accesibilidad** | Todo se desactiva con `prefers-reduced-motion`; el PDF se exporta con las animaciones apagadas |

Atajos: `←` `→` navegar · `N` notas · `F` pantalla completa · `P` imprimir/PDF.

## 📁 Archivos del repositorio

| Archivo | Qué es |
|---|---|
| [`Induction_Project_Davis_Fernandez_FINAL.pptx`](Induction_Project_Davis_Fernandez_FINAL.pptx) | **Entregable 1** — 4 diapositivas (16:9) con notas del orador. **311 palabras visibles**. |
| [`Welcome_Kit_Escrito_Davis_Fernandez.docx`](Welcome_Kit_Escrito_Davis_Fernandez.docx) | **Entregable 2** — parte escrita. **347 palabras** (rango 250–350). |
| [`Welcome_Kit_Davis_Fernandez.pdf`](Welcome_Kit_Davis_Fernandez.pdf) | **Entregable 3** — la misma presentación web exportada a PDF (4 páginas, lista para Classroom). |
| [`avatar-davis.png`](avatar-davis.png) / [`avatar-davis.svg`](avatar-davis.svg) | Ilustración plana del técnico (PNG para el PPTX y la web; SVG es la fuente editable). |
| [`mapa-pikta.png`](mapa-pikta.png) | Captura de Google Maps con la ruta de La Riviera a PikTa (la usa la web y el PPTX). |
| [`Guia_Induction_Project.docx`](Guia_Induction_Project.docx) | Guía de calificación: checklist de la rúbrica, datos pendientes y cronometraje. |
| [`index.html`](index.html) | **Versión web** de la misma presentación (autocontenida, sin librerías externas). |
| [`build_deck.py`](build_deck.py) / [`build_word.py`](build_word.py) / [`build_guia.py`](build_guia.py) | Scripts que generan los tres documentos. |
| [`screenshots/`](screenshots/) | Capturas de las 4 diapositivas de la versión web. |

## 📋 Cada punto de la rúbrica, dónde está

| Requisito de la rúbrica | Dónde se cumple |
|---|---|
| Exactamente **4 diapositivas** (una por sección) | Portada + 2 · About me → 3 · How to get → 4 · Around the office & Safety |
| Secciones: *About me / How to get to workplace / Around the office / Safety first* | Portada + diapositivas 2, 3 y 4 (Around the office y Safety first comparten la 4ª, como en la muestra del profesor) |
| **5+ preposiciones de lugar/dirección** (*on, near, behind, in, to, between, across…*) | Diapositivas 2, 3 y 4 + parte escrita |
| **Vocabulario de directions** (*turn left/right, go straight, go east/north, one block ahead, next to*) | Diapositiva 3 + parte escrita |
| **Present simple + present continuous** | *"I start/finish…"* + *"This week I am setting up a network…"* (dip. 2) |
| **Una frase en pasado y una en futuro** | *"Yesterday I finished work and went to dinner…"* / *"Tomorrow I will have a long day…"* (dip. 4) |
| **Mapa con la ruta marcada** | Diapositiva 3: captura de Google Maps, desde la parada *Romero La Riviera* → Carretera 48 → Av. 8a Oeste / Ave. La Libertad → giro a la izquierda en C. Arístides Romero → **Restaurante PikTa** (junto al Hospital Chiriquí, a 50 m de Julio Gato), unos 10 min / 5 km |
| **Direcciones desde un lugar conocido** | *"From the Romero La Riviera bus stop, go north on Carretera 48…"* (dip. 3 y parte escrita) |
| **Quantifiers** (*some, any, a lot of, many, much*) | *"There aren't any… but there are a lot of… Some are cheap and many people eat… I don't have much time."* (dip. 4) |
| **Secuencia** (*First, Then, After that, Finally*) | Diapositiva 3, los 4 pasos numerados |
| **Descripción de primeros auxilios** | Diapositiva 4: kit *in my car, in a red box* → *if someone is bleeding, apply pressure and call 911* |
| Extensión 250–350 palabras | 347 en Word / 311 visibles en PPT |
| Oral 4–5 minutos | 561 palabras en las notas ≈ 4,3–4,9 min |

## 🎤 Cronometraje del oral (4–5 min)

1. **Diapositiva 1** — saludo y objetivo (~30 s)
2. **Diapositiva 2** — perfil, horario y rutina (~1:10)
3. **Diapositiva 3** — base en La Riviera, ruta a PikTa con *First / Then / After that / Finally* (~1:15)
4. **Diapositiva 4** — comida, ayer/mañana y seguridad (~1:20)
5. Cierre: *"Thank you. Any questions?"*

## 🛠️ Reconstruir los archivos

```bash
pip install -r requirements.txt
python build_deck.py     # → Induction_Project_Davis_Fernandez_FINAL.pptx (+ contador de palabras)
python build_word.py     # → Welcome_Kit_Escrito_Davis_Fernandez.docx (+ contador de palabras)
python build_guia.py     # → Guia_Induction_Project.docx
```

`build_deck.py` guarda en `_build.pptx` si el archivo está abierto en PowerPoint; ciérralo y vuelve a
ejecutar. El PDF de la presentación **web** se genera con `Ctrl+P` sobre `index.html` → *Guardar como
PDF*, papel **13.333 × 7.5 in**, márgenes **ninguno**, marcar **gráficos de fondo** (en la web la tecla
`P` hace lo mismo). También se puede exportar el `.pptx` desde PowerPoint → *Guardar como → PDF*.

La versión web no tiene dependencias externas: es un solo archivo `index.html` más las dos imágenes
locales (`mapa-pikta.png` y `avatar-davis.png`). Para probarla en local:

```bash
python -m http.server 8080   # y abre http://localhost:8080
```

## ✅ Checklist antes de entregar

- [ ] `Induction_Project_Davis_Fernandez_FINAL.pptx` (4 diapositivas, revisar notas con `Alt+F5`)
- [ ] `Welcome_Kit_Escrito_Davis_Fernandez.docx` (347 palabras)
- [ ] `Welcome_Kit_Davis_Fernandez.pdf` (subir a Classroom si piden PDF)
- [ ] Enlace de GitHub Pages en el comentario/entrega de Classroom
- [ ] Ensayo cronometrado: 4–5 minutos

## ⚠️ Datos que aún se pueden afinar

1. Dirección exacta de la base en La Riviera (número de casa / referencia).
2. Confirmar el giro exacto de la ruta (hoy: Calle 5 → Calle 4 → Calle Arístides Romero).
3. Dónde se estaciona en PikTa (ahora: *behind the building*).
4. Extensión interna o seguro médico (opcional, junto al 911).

---

Repositorio público por decisión del autor. Licencia: uso académico.
