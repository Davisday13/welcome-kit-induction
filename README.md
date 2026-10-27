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

## 📁 Archivos del repositorio

| Archivo | Qué es |
|---|---|
| [`Induction_Project_Davis_Fernandez_FINAL.pptx`](Induction_Project_Davis_Fernandez_FINAL.pptx) | **Entregable 1** — 4 diapositivas (16:9) con notas del orador. **340 palabras visibles**. |
| [`Welcome_Kit_Escrito_Davis_Fernandez.docx`](Welcome_Kit_Escrito_Davis_Fernandez.docx) | **Entregable 2** — parte escrita. **338 palabras** (rango 250–350). |
| [`Welcome_Kit_Davis_Fernandez.pdf`](Welcome_Kit_Davis_Fernandez.pdf) | **Entregable 3** — la presentación exportada a PDF (4 páginas, lista para Classroom). |
| [`Guia_Induction_Project.docx`](Guia_Induction_Project.docx) | Guía de calificación: checklist de la rúbrica, datos pendientes y cronometraje. |
| [`index.html`](index.html) | **Versión web** de la misma presentación (autocontenida, sin librerías externas). |
| [`build_deck.py`](build_deck.py) / [`build_word.py`](build_word.py) / [`build_guia.py`](build_guia.py) | Scripts que generan los tres documentos. |
| [`screenshots/`](screenshots/) | Capturas de las 4 diapositivas de la versión web. |

## 📋 Cada punto de la rúbrica, dónde está

| Requisito de la rúbrica | Dónde se cumple |
|---|---|
| Exactamente **4 diapositivas** (una por sección) | 1 · About me → 2 · How to get → 3/4 · Around the office & Safety |
| Secciones: *About me / How to get to workplace / Around the office / Safety first* | Diapositivas 1, 2, 3 y 4 (3 y 4 comparten la 4ª, como en la muestra del profesor) |
| **5+ preposiciones de lugar/dirección** (*on, near, behind, in, to, between, across…*) | Diapositivas 2, 3 y 4 + parte escrita |
| **Vocabulario de directions** (*turn left/right, go straight, go east/north, one block ahead, next to*) | Diapositiva 3 + parte escrita |
| **Present simple + present continuous** | *"I start/finish…"* + *"This week I am setting up a network…"* (dip. 2) |
| **Una frase en pasado y una en futuro** | *"Yesterday I finished work and went to dinner…"* / *"Tomorrow I will have a long day…"* (dip. 4) |
| **Mapa con la ruta marcada** | Diapositiva 3: La Riviera (C. 5) → C. 4 → C. Arístides Romero → **Restaurante PikTa**, línea naranja, brújula y Ozaga Market |
| **Quantifiers** (*some, any, a lot of, many, much*) | *"There aren't any… but there are a lot of… Some are cheap and many people eat… I don't have much time."* (dip. 4) |
| **Secuencia** (*First, Then, After that, Finally*) | Diapositiva 3, los 4 pasos numerados |
| **Descripción de primeros auxilios** | Diapositiva 4: kit *in my car, in a red box* → *if someone is bleeding, apply pressure and call 911* |
| Extensión 250–350 palabras | 338 en Word / 340 visibles en PPT |
| Oral 4–5 minutos | 548 palabras en las notas ≈ 4,2–4,9 min |

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
ejecutar. El PDF se exporta desde PowerPoint: **Guardar como → PDF** (o desde `index.html` con la tecla `P`).

La versión web no tiene dependencias: es un solo archivo `index.html`. Para probarla en local:

```bash
python -m http.server 8080   # y abre http://localhost:8080
```

## ✅ Checklist antes de entregar

- [ ] `Induction_Project_Davis_Fernandez_FINAL.pptx` (4 diapositivas, revisar notas con `Alt+F5`)
- [ ] `Welcome_Kit_Escrito_Davis_Fernandez.docx` (338 palabras)
- [ ] `Welcome_Kit_Davis_Fernandez.pdf` (subir a Classroom si piden PDF)
- [ ] Enlace de GitHub Pages en el comentario/entrega de Classroom
- [ ] Ensayo cronometrado: 4–5 minutos

## ⚠️ Datos que aún se pueden afinar

1. Dirección exacta de la base en La Riviera (número de casa / referencia).
2. Confirmar el giro exacto de la ruta (hoy: Calle 5 → Calle 4 → Calle Arístides Romero).
3. Dónde se estaciona en PikTa (ahora: *behind the building*).
4. Nombres reales de los dos lugares de comida (hoy genéricos: *the bakery on the corner* y *the market*).
5. Extensión interna o seguro médico (opcional, junto al 911).

---

Repositorio público por decisión del autor. Licencia: uso académico.
