const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'CRIM Project';
pres.title = 'CRIM Intervals - Herramientas de Análisis para la Polífonia Renacentista';

// Color palette
const C = {
  burgundy: "6D2E46",
  cream: "ECE2D0",
  dustyRose: "A26769",
  darkBurgundy: "4A1E30",
  white: "FFFFFF",
  lightCream: "F5EFE6",
  medBurgundy: "8B3A54",
  darkText: "2C1018",
  roseLight: "C8979A",
};

// Helper: left accent bar for content slides
function addAccentBar(slide, yStart, height) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: yStart, w: 0.12, h: height,
    fill: { color: C.burgundy }, line: { color: C.burgundy, width: 0 }
  });
}

function addDarkBg(slide) {
  slide.background = { color: C.darkBurgundy };
}

function addLightBg(slide) {
  slide.background = { color: C.cream };
}

// Helper: decorative corner ornament for dark slides
function addCornerOrnament(slide) {
  slide.addShape(pres.shapes.OVAL, {
    x: 7.8, y: -1.2, w: 3.5, h: 3.5,
    fill: { color: C.burgundy, transparency: 60 },
    line: { color: C.burgundy, width: 0 }
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 8.5, y: -0.5, w: 2.2, h: 2.2,
    fill: { color: C.dustyRose, transparency: 70 },
    line: { color: C.dustyRose, width: 0 }
  });
}

// Helper: badge/tag shape
function addBadge(slide, text, x, y) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: 3.4, h: 0.38,
    fill: { color: C.burgundy }, line: { color: C.burgundy, width: 0 }
  });
  slide.addText(text, {
    x: x, y: y, w: 3.4, h: 0.38,
    fontSize: 12, fontFace: "Calibri", color: C.cream,
    bold: true, align: "center", valign: "middle", margin: 0
  });
}

// Helper: top header band
function addHeaderBand(slide, title, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.0,
    fill: { color: color || C.burgundy }, line: { color: color || C.burgundy, width: 0 }
  });
  slide.addText(title, {
    x: 0.3, y: 0.0, w: 9.4, h: 1.0,
    fontSize: 30, fontFace: "Georgia", color: C.cream,
    bold: true, align: "left", valign: "middle", margin: [0, 0, 0, 0.2]
  });
}

// ============================================================
// SLIDE 1 - COVER
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.darkBurgundy };

  // Large decorative overlapping circles (musical motif)
  slide.addShape(pres.shapes.OVAL, {
    x: 6.5, y: 0.2, w: 4.2, h: 4.2,
    fill: { color: C.burgundy, transparency: 30 },
    line: { color: C.medBurgundy, width: 2 }
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 7.2, y: 1.0, w: 2.8, h: 2.8,
    fill: { color: C.dustyRose, transparency: 50 },
    line: { color: C.roseLight, width: 1 }
  });

  // Decorative staff lines (musical motif)
  for (let i = 0; i < 5; i++) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.2, y: 1.5 + i * 0.28, w: 3.5, h: 0.04,
      fill: { color: C.cream, transparency: 60 },
      line: { color: C.cream, width: 0 }
    });
  }

  // Note head
  slide.addShape(pres.shapes.OVAL, {
    x: 7.6, y: 2.1, w: 0.55, h: 0.42,
    fill: { color: C.cream },
    line: { color: C.cream, width: 0 }
  });
  // Note stem
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 8.12, y: 1.05, w: 0.05, h: 1.1,
    fill: { color: C.cream },
    line: { color: C.cream, width: 0 }
  });

  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  slide.addText("CRIM Intervals", {
    x: 0.5, y: 0.7, w: 6.0, h: 1.1,
    fontSize: 48, fontFace: "Georgia", color: C.cream,
    bold: true, align: "left", valign: "middle"
  });

  slide.addText("Herramientas de Análisis para la Polifonía Renacentista", {
    x: 0.5, y: 1.85, w: 6.0, h: 0.8,
    fontSize: 19, fontFace: "Georgia", color: C.roseLight,
    italic: true, align: "left", valign: "middle"
  });

  slide.addText("Una introducción para estudiantes de musicología", {
    x: 0.5, y: 2.72, w: 6.0, h: 0.55,
    fontSize: 15, fontFace: "Calibri", color: C.cream,
    align: "left", valign: "middle"
  });

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.42, w: 5.5, h: 0.04,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  slide.addText("Citations: The Renaissance Imitation Mass  ·  crimproject.org", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: C.roseLight,
    align: "left", valign: "middle"
  });
}

// ============================================================
// SLIDE 2 - ¿Qué es el Proyecto CRIM?
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "¿Qué es el Proyecto CRIM?");

  // Left column header
  slide.addText("Sobre el Proyecto", {
    x: 0.3, y: 1.15, w: 4.8, h: 0.42,
    fontSize: 16, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left", valign: "middle"
  });

  slide.addText([
    { text: "CRIM = Citations: The Renaissance Imitation Mass", options: { bullet: true, breakLine: true, bold: true } },
    { text: "Estudia cómo compositores del s. XVI transformaron obras en ciclos de Misa", options: { bullet: true, breakLine: true } },
    { text: "Dirigido por el Prof. Richard Freedman (Haverford College)", options: { bullet: true } },
  ], {
    x: 0.3, y: 1.62, w: 4.7, h: 2.0,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Right column — stat callouts
  const stats = [
    { num: "50+", label: "investigadores internacionales" },
    { num: "~3.000", label: "relaciones documentadas" },
    { num: "2", label: "instituciones principales" },
  ];

  stats.forEach((s, i) => {
    let bx = 5.4, by = 1.18 + i * 1.15;
    let boxColor = i % 2 === 0 ? C.burgundy : C.dustyRose;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: 4.2, h: 0.95,
      fill: { color: boxColor },
      line: { color: boxColor, width: 0 }
    });
    slide.addText(s.num, {
      x: bx + 0.15, y: by, w: 1.5, h: 0.95,
      fontSize: 28, fontFace: "Georgia", color: C.cream,
      bold: true, align: "left", valign: "middle", margin: 0
    });
    slide.addText(s.label, {
      x: bx + 1.7, y: by, w: 2.35, h: 0.95,
      fontSize: 13, fontFace: "Calibri", color: C.cream,
      align: "left", valign: "middle", margin: 0
    });
  });

  slide.addText("Haverford College  +  Centre d'Études Supérieures de la Renaissance (Tours, Francia)", {
    x: 0.3, y: 4.62, w: 9.4, h: 0.5,
    fontSize: 11, fontFace: "Calibri", color: C.dustyRose,
    italic: true, align: "left"
  });

  // Decorative shape
  slide.addShape(pres.shapes.OVAL, {
    x: 8.5, y: 3.85, w: 1.7, h: 1.7,
    fill: { color: C.burgundy, transparency: 85 },
    line: { color: C.dustyRose, width: 1 }
  });
}

// ============================================================
// SLIDE 3 - La Herramienta Streamlit
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Acceso Sin Código: La Aplicación Streamlit");

  const rows = [
    { color: C.burgundy, icon: "*", title: "Sin programación", desc: "Interfaz web intuitiva, sin necesidad de conocimientos técnicos" },
    { color: C.dustyRose, icon: "=", title: "Formatos abiertos", desc: "Compatible con MEI, MusicXML y MIDI" },
    { color: C.burgundy, icon: "v", title: "Resultados descargables", desc: "Tablas, gráficos y CSV para publicaciones" },
    { color: C.dustyRose, icon: "@", title: "Corpus integrado", desc: "Acceso directo al corpus CRIM + tus propias partituras" },
  ];

  rows.forEach((r, i) => {
    let ry = 1.18 + i * 0.88;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.3, y: ry + 0.06, w: 0.6, h: 0.6,
      fill: { color: r.color }, line: { color: r.color, width: 0 }
    });
    slide.addText(["1", "2", "3", "4"][i], {
      x: 0.3, y: ry + 0.06, w: 0.6, h: 0.6,
      fontSize: 16, fontFace: "Georgia", color: C.cream,
      align: "center", valign: "middle", bold: true, margin: 0
    });
    slide.addText(r.title, {
      x: 1.1, y: ry + 0.04, w: 2.8, h: 0.32,
      fontSize: 15, fontFace: "Georgia", color: C.burgundy,
      bold: true, align: "left", valign: "middle"
    });
    slide.addText(r.desc, {
      x: 1.1, y: ry + 0.36, w: 8.6, h: 0.32,
      fontSize: 13, fontFace: "Calibri", color: C.darkText,
      align: "left", valign: "middle"
    });
  });

  // Bottom tech stack
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.88, w: 9.4, h: 0.45,
    fill: { color: C.darkBurgundy }, line: { color: C.darkBurgundy, width: 0 }
  });
  slide.addText("Built on:  Python  ·  CRIM Intervals Library  ·  music21  ·  Streamlit", {
    x: 0.3, y: 4.88, w: 9.4, h: 0.45,
    fontSize: 12, fontFace: "Calibri", color: C.cream,
    align: "center", valign: "middle", bold: true, margin: 0
  });
}

// ============================================================
// SLIDE 4 - El Repertorio (dark section marker)
// ============================================================
{
  let slide = pres.addSlide();
  addDarkBg(slide);
  addCornerOrnament(slide);

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.25, h: 5.625,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  slide.addText("El Repertorio", {
    x: 0.5, y: 0.4, w: 7.5, h: 0.7,
    fontSize: 40, fontFace: "Georgia", color: C.cream,
    bold: true, align: "left"
  });
  slide.addText("La Misa Imitación Renacentista", {
    x: 0.5, y: 1.1, w: 7.5, h: 0.45,
    fontSize: 18, fontFace: "Georgia", color: C.roseLight,
    italic: true, align: "left"
  });

  const bullets4 = [
    "Género clave del siglo XVI: la Misa parodia o de imitación",
    "Técnica compositiva: transformar motetes, chansons y otras piezas en ciclos de Misa de 5 movimientos",
    "Compositores: Josquin · Palestrina · Lassus · Morales · Victoria",
    "Técnicas contrapuntísticas: fugas, dúos imitativos, entradas periódicas",
    "El análisis computacional permite estudiar este repertorio a escala masiva",
  ];

  bullets4.forEach((b, i) => {
    let by = 1.72 + i * 0.63;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: by + 0.13, w: 0.22, h: 0.22,
      fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
    });
    slide.addText(b, {
      x: 0.9, y: by, w: 8.8, h: 0.5,
      fontSize: 14, fontFace: "Calibri", color: C.cream,
      align: "left", valign: "middle"
    });
  });
}

// ============================================================
// SLIDE 5 - Herramienta 1: Intervalos Melódicos
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 1: Análisis de Intervalos Melódicos");

  addBadge(slide, "piece.melodic()", 0.3, 1.1);

  // Left column
  slide.addText("¿Qué hace?", {
    x: 0.3, y: 1.62, w: 4.5, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Identifica y cuantifica todos los intervalos melódicos de cada voz", options: { bullet: true, breakLine: true } },
    { text: "Opciones: diatónico o cromático", options: { bullet: true, breakLine: true } },
    { text: "Con o sin calidad interválica (mayor, menor, perfecto)", options: { bullet: true } },
  ], {
    x: 0.3, y: 2.08, w: 4.5, h: 1.8,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.1, w: 0.06, h: 3.8,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  // Right column
  slide.addText("Utilidad para la polifonía renacentista", {
    x: 5.3, y: 1.62, w: 4.4, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Detectar patrones motívicos recurrentes", options: { bullet: true, breakLine: true } },
    { text: "Comparar los perfiles melódicos entre voces", options: { bullet: true, breakLine: true } },
    { text: "Identificar material prestado (citación melódica)", options: { bullet: true, breakLine: true } },
    { text: "Estudiar el estilo de cada compositor", options: { bullet: true } },
  ], {
    x: 5.3, y: 2.08, w: 4.4, h: 2.2,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Decorative element
  slide.addShape(pres.shapes.OVAL, {
    x: 8.6, y: 4.1, w: 1.2, h: 1.2,
    fill: { color: C.burgundy, transparency: 80 },
    line: { color: C.dustyRose, width: 1.5 }
  });
}

// ============================================================
// SLIDE 6 - Herramienta 2: Intervalos Armónicos
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 2: Análisis de Intervalos Armónicos", C.dustyRose);

  addBadge(slide, "piece.harmonic()", 0.3, 1.1);

  // Left column
  slide.addText("¿Qué hace?", {
    x: 0.3, y: 1.62, w: 4.5, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Examina los intervalos verticales entre cada par de voces", options: { bullet: true, breakLine: true } },
    { text: "Mide la distancia interválica en cada punto de la textura", options: { bullet: true } },
  ], {
    x: 0.3, y: 2.08, w: 4.5, h: 1.5,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Mini bar chart visualization
  const intervalData = [
    { label: "8va", w: 3.2, color: C.burgundy },
    { label: "5ta", w: 2.6, color: C.dustyRose },
    { label: "3ra", w: 2.0, color: C.medBurgundy },
    { label: "6ta", w: 2.3, color: C.roseLight },
  ];
  intervalData.forEach((d, i) => {
    let by = 3.55 + i * 0.35;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: by, w: d.w, h: 0.27,
      fill: { color: d.color }, line: { color: d.color, width: 0 }
    });
    slide.addText(d.label, {
      x: 0.3 + d.w + 0.1, y: by, w: 0.7, h: 0.27,
      fontSize: 11, fontFace: "Calibri", color: C.darkText,
      align: "left", valign: "middle"
    });
  });

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.1, w: 0.06, h: 3.8,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  // Right column
  slide.addText("Utilidad para la polifonía renacentista", {
    x: 5.3, y: 1.62, w: 4.4, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Detectar consonancias y disonancias características del estilo", options: { bullet: true, breakLine: true } },
    { text: "Analizar el tratamiento del contrapunto estricto", options: { bullet: true, breakLine: true } },
    { text: "Identificar progresiones armónicas características", options: { bullet: true, breakLine: true } },
    { text: "Estudiar la densidad contrapuntística de la textura", options: { bullet: true } },
  ], {
    x: 5.3, y: 2.08, w: 4.4, h: 2.2,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });
}

// ============================================================
// SLIDE 7 - Herramienta 3: N-gramas
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 3: N-gramas y Módulos Contrapuntísticos");

  // Large callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 9.4, h: 0.65,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });
  slide.addText("El corazón analítico de CRIM Intervals", {
    x: 0.3, y: 1.1, w: 9.4, h: 0.65,
    fontSize: 18, fontFace: "Georgia", color: C.cream,
    bold: true, italic: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Un módulo contrapuntístico combina:", {
    x: 0.3, y: 1.88, w: 5.0, h: 0.38,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    bold: true, align: "left"
  });

  // Two combination boxes
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 2.32, w: 3.8, h: 0.72,
    fill: { color: C.burgundy }, line: { color: C.burgundy, width: 0 }
  });
  slide.addText("Intervalos verticales entre voces", {
    x: 0.3, y: 2.32, w: 3.8, h: 0.72,
    fontSize: 13, fontFace: "Calibri", color: C.cream,
    align: "center", valign: "middle", bold: true, margin: 0
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.3, y: 2.32, w: 3.8, h: 0.72,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });
  slide.addText("Movimiento melódico de la voz inferior", {
    x: 4.3, y: 2.32, w: 3.8, h: 0.72,
    fontSize: 13, fontFace: "Calibri", color: C.cream,
    align: "center", valign: "middle", bold: true, margin: 0
  });

  // Plus sign between boxes
  slide.addText("+", {
    x: 4.05, y: 2.32, w: 0.3, h: 0.72,
    fontSize: 22, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "center", valign: "middle", margin: 0
  });

  // Code example box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 3.18, w: 5.2, h: 0.65,
    fill: { color: C.darkBurgundy }, line: { color: C.darkBurgundy, width: 0 }
  });
  slide.addText('"7_Held · 6_-2 · 8"', {
    x: 0.3, y: 3.18, w: 5.2, h: 0.65,
    fontSize: 18, fontFace: "Calibri", color: C.cream,
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Intervalos verticales 7→6→8 + nota ligada + descenso de 2a = fórmula cadencial típica", {
    x: 0.3, y: 3.95, w: 9.4, h: 0.4,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    italic: true, align: "left"
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.5, w: 9.4, h: 0.75,
    fill: { color: C.lightCream }, line: { color: C.dustyRose, width: 1 }
  });
  slide.addText("Utilidad: Detectar citas y transformaciones  ·  Comparar técnicas entre compositores", {
    x: 0.3, y: 4.5, w: 9.4, h: 0.75,
    fontSize: 13, fontFace: "Calibri", color: C.burgundy,
    bold: true, align: "center", valign: "middle", margin: 0
  });
}

// ============================================================
// SLIDE 8 - Herramienta 4: Cadencias
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 4: Detección de Cadencias", C.dustyRose);

  addBadge(slide, "piece.cadences()", 0.3, 1.1);

  // Three columns
  const cols8 = [
    {
      title: "Tipos detectados",
      items: ["Cadencia auténtica", "Cadencia frigia", "Cadencia evitada", "Cadencia de engaño"],
      x: 0.25, color: C.burgundy
    },
    {
      title: "Información extraída",
      items: ["Voz que realiza la cláusula", "Tono de la cadencia", "Posición métrica y formal"],
      x: 3.6, color: C.dustyRose
    },
    {
      title: "Utilidad musicológica",
      items: ["Trazar el perfil tonal (modal)", "Comparar modelo vs. Misa", "Analizar la articulación formal"],
      x: 6.7, color: C.medBurgundy
    },
  ];

  cols8.forEach((col) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: col.x, y: 1.65, w: 3.0, h: 0.45,
      fill: { color: col.color }, line: { color: col.color, width: 0 }
    });
    slide.addText(col.title, {
      x: col.x, y: 1.65, w: 3.0, h: 0.45,
      fontSize: 13, fontFace: "Georgia", color: C.cream,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    col.items.forEach((item, j) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: col.x, y: 2.18 + j * 0.75, w: 3.0, h: 0.65,
        fill: { color: C.lightCream }, line: { color: col.color, width: 1 }
      });
      slide.addText(item, {
        x: col.x + 0.1, y: 2.18 + j * 0.75, w: 2.8, h: 0.65,
        fontSize: 13, fontFace: "Calibri", color: C.darkText,
        align: "center", valign: "middle", margin: 0
      });
    });
  });
}

// ============================================================
// SLIDE 9 - Herramienta 5: Tipos de Presentación
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 5: Tipos de Presentación Imitativa");

  addBadge(slide, "piece.presentationTypes()", 0.3, 1.1);

  const types9 = [
    { name: "FUGA", desc: "Canon o quasi-canon entre voces" },
    { name: "DUO IMITATIVO", desc: "Dos voces en imitación estricta" },
    { name: "DUO NO IMITATIVO", desc: "Dos voces en contrapunto libre" },
    { name: "ENTRADAS PERIÓDICAS", desc: "Todas las voces con el mismo motivo" },
    { name: "HOMOFONÍA RÍTMICA", desc: "piece.homorhythm() — Movimiento sincrónico de voces" },
  ];

  types9.forEach((t, i) => {
    let ty = 1.62 + i * 0.63;
    let sqColor = i % 2 === 0 ? C.burgundy : C.dustyRose;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: ty + 0.08, w: 0.44, h: 0.42,
      fill: { color: sqColor }, line: { color: sqColor, width: 0 }
    });
    slide.addText(t.name, {
      x: 0.9, y: ty, w: 2.8, h: 0.55,
      fontSize: 14, fontFace: "Georgia", color: C.burgundy,
      bold: true, align: "left", valign: "middle"
    });
    slide.addText("— " + t.desc, {
      x: 3.7, y: ty, w: 5.9, h: 0.55,
      fontSize: 13, fontFace: "Calibri", color: C.darkText,
      align: "left", valign: "middle"
    });
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.9, w: 9.4, h: 0.44,
    fill: { color: C.darkBurgundy }, line: { color: C.darkBurgundy, width: 0 }
  });
  slide.addText("Utilidad: Mapear la arquitectura contrapuntística  ·  Comparar uso del material en la Misa", {
    x: 0.3, y: 4.9, w: 9.4, h: 0.44,
    fontSize: 12, fontFace: "Calibri", color: C.cream,
    align: "center", valign: "middle", bold: true, margin: 0
  });
}

// ============================================================
// SLIDE 10 - Herramienta 6: Heatmaps
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 6: Mapas de Calor (Heatmaps)", C.dustyRose);

  // Decorative heatmap grid (left portion of slide)
  const heatRows = 7;
  const heatCols = 12;
  const heatX = 0.25;
  const heatY = 1.1;
  const cellW = 0.43;
  const cellH = 0.33;

  const heatColors = [
    C.lightCream, C.roseLight, C.dustyRose, C.medBurgundy, C.burgundy, C.darkBurgundy
  ];

  for (let r = 0; r < heatRows; r++) {
    for (let c = 0; c < heatCols; c++) {
      let val = Math.sin(c * 0.5 + r * 0.7) * 0.5 + Math.cos(c * 0.3 - r * 0.4) * 0.4;
      let intensity = Math.max(0, Math.min(5, Math.floor((val + 1) * 2.5)));
      slide.addShape(pres.shapes.RECTANGLE, {
        x: heatX + c * cellW,
        y: heatY + r * cellH,
        w: cellW - 0.03,
        h: cellH - 0.03,
        fill: { color: heatColors[intensity] },
        line: { color: C.cream, width: 0.5 }
      });
    }
  }

  // Description right side
  slide.addText("¿Qué muestran?", {
    x: 5.7, y: 1.2, w: 4.0, h: 0.42,
    fontSize: 15, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Eje horizontal: posición en la partitura (compases)", options: { bullet: true, breakLine: true } },
    { text: "Eje vertical: tipo de patrón o voz", options: { bullet: true, breakLine: true } },
    { text: "Color: frecuencia o densidad del patrón", options: { bullet: true } },
  ], {
    x: 5.7, y: 1.67, w: 4.0, h: 1.4,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  slide.addText("Utilidad", {
    x: 5.7, y: 3.15, w: 4.0, h: 0.42,
    fontSize: 15, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Ver dónde se concentran las citas y transformaciones", options: { bullet: true, breakLine: true } },
    { text: "Comparar arquitectura formal de modelo y Misa", options: { bullet: true, breakLine: true } },
    { text: "Identificar secciones de imitación densa vs. textura libre", options: { bullet: true } },
  ], {
    x: 5.7, y: 3.62, w: 4.0, h: 1.7,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });
}

// ============================================================
// SLIDE 11 - Herramienta 7: Radar Plots
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 7: Radar Plots y Gráficos de Cadencias");

  // Left description
  slide.addText("Radar Plot (diagrama radial)", {
    x: 0.3, y: 1.15, w: 4.5, h: 0.42,
    fontSize: 15, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Muestra el perfil tonal de la obra", options: { bullet: true, breakLine: true } },
    { text: "Cada \"radio\" = un grado o tono cadencial", options: { bullet: true, breakLine: true } },
    { text: "Compara el perfil modal de varias obras", options: { bullet: true } },
  ], {
    x: 0.3, y: 1.62, w: 4.5, h: 1.4,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Decorative radar shape — concentric rings
  const cx = 2.5, cy = 4.0;
  for (let ring = 3; ring >= 1; ring--) {
    let rr = ring * 0.38;
    slide.addShape(pres.shapes.OVAL, {
      x: cx - rr, y: cy - rr,
      w: rr * 2, h: rr * 2,
      fill: { color: ring === 1 ? C.burgundy : ring === 2 ? C.dustyRose : C.roseLight, transparency: 65 + ring * 3 },
      line: { color: C.burgundy, width: 1 }
    });
  }
  // Axis lines from center
  for (let a = 0; a < 8; a++) {
    let angle = (a / 8) * Math.PI * 2 - Math.PI / 2;
    let r = 1.1;
    let ex = Math.cos(angle) * r;
    let ey = Math.sin(angle) * r;
    slide.addShape(pres.shapes.LINE, {
      x: cx, y: cy,
      w: ex, h: ey,
      line: { color: C.dustyRose, width: 1.2 }
    });
  }

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.1, w: 0.06, h: 4.1,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  // Right column
  slide.addText("Progress Charts (Cadencias)", {
    x: 5.3, y: 1.15, w: 4.4, h: 0.42,
    fontSize: 15, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Cadencias en secuencia temporal", options: { bullet: true, breakLine: true } },
    { text: "Visualiza la planificación tonal", options: { bullet: true, breakLine: true } },
    { text: "Analiza la articulación formal de los movimientos", options: { bullet: true } },
  ], {
    x: 5.3, y: 1.62, w: 4.4, h: 1.4,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Mini bar chart
  const barData11 = [0.3, 0.7, 0.5, 0.9, 0.4, 0.6, 0.8, 0.35, 0.55, 0.75];
  const barColors11 = [C.burgundy, C.dustyRose, C.burgundy, C.medBurgundy, C.dustyRose, C.burgundy, C.medBurgundy, C.dustyRose, C.burgundy, C.dustyRose];
  barData11.forEach((h11, i) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3 + i * 0.41, y: 4.5 - h11 * 1.2, w: 0.34, h: h11 * 1.2,
      fill: { color: barColors11[i] }, line: { color: barColors11[i], width: 0 }
    });
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 4.5, w: 4.1, h: 0.04,
    fill: { color: C.darkText }, line: { color: C.darkText, width: 0 }
  });
  slide.addText("Secuencia temporal de cadencias", {
    x: 5.3, y: 4.6, w: 4.1, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: C.dustyRose,
    italic: true, align: "center"
  });
}

// ============================================================
// SLIDE 12 - Herramienta 8: Network Diagrams
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Herramienta 8: Diagramas de Red", C.dustyRose);

  // Decorative network diagram (left)
  const nodes12 = [
    { x: 0.4, y: 2.8, r: 0.3, color: C.burgundy, label: "Motete" },
    { x: 2.0, y: 1.6, r: 0.25, color: C.dustyRose, label: "Kyrie" },
    { x: 2.0, y: 4.0, r: 0.25, color: C.dustyRose, label: "Gloria" },
    { x: 3.6, y: 1.1, r: 0.2, color: C.medBurgundy, label: "Credo" },
    { x: 3.6, y: 2.8, r: 0.22, color: C.medBurgundy, label: "Sanctus" },
    { x: 3.6, y: 4.5, r: 0.2, color: C.dustyRose, label: "Agnus" },
  ];

  // Edges
  [[0,1],[0,2],[0,4],[1,3],[1,4],[2,4],[2,5],[3,4],[4,5]].forEach(([a, b]) => {
    let na = nodes12[a], nb = nodes12[b];
    slide.addShape(pres.shapes.LINE, {
      x: na.x + na.r, y: na.y + na.r,
      w: nb.x - na.x, h: nb.y - na.y,
      line: { color: C.dustyRose, width: 1.2 }
    });
  });

  // Node circles
  nodes12.forEach(n => {
    slide.addShape(pres.shapes.OVAL, {
      x: n.x, y: n.y, w: n.r * 2, h: n.r * 2,
      fill: { color: n.color }, line: { color: C.cream, width: 1.5 }
    });
    slide.addText(n.label, {
      x: n.x - 0.15, y: n.y + n.r * 2 + 0.04, w: 0.9, h: 0.28,
      fontSize: 10, fontFace: "Calibri", color: C.darkText,
      align: "center", valign: "top"
    });
  });

  // Content right side
  slide.addText("Nodos = obras o motivos musicales\nAristas = similitudes, citas o relaciones", {
    x: 5.0, y: 1.2, w: 4.7, h: 0.8,
    fontSize: 14, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  slide.addText("Utilidad", {
    x: 5.0, y: 2.1, w: 4.7, h: 0.42,
    fontSize: 15, fontFace: "Georgia", color: C.burgundy,
    bold: true, align: "left"
  });
  slide.addText([
    { text: "Visualizar la red de influencias entre modelo y misas", options: { bullet: true, breakLine: true } },
    { text: "Detectar agrupaciones de obras por patrones compartidos", options: { bullet: true, breakLine: true } },
    { text: "Mostrar la genealogía de un motivo a través del corpus", options: { bullet: true, breakLine: true } },
    { text: "Investigar la transmisión musical y el estilo compositivo", options: { bullet: true } },
  ], {
    x: 5.0, y: 2.58, w: 4.7, h: 2.4,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });
}

// ============================================================
// SLIDE 13 - Herramienta 9: Corpus Tools (dark)
// ============================================================
{
  let slide = pres.addSlide();
  addDarkBg(slide);
  addCornerOrnament(slide);

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.25, h: 5.625,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  slide.addText("Herramienta 9: Análisis de Corpus", {
    x: 0.5, y: 0.4, w: 8.5, h: 0.7,
    fontSize: 36, fontFace: "Georgia", color: C.cream,
    bold: true, align: "left"
  });
  slide.addText("Del análisis individual al estudio a gran escala", {
    x: 0.5, y: 1.1, w: 8.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.roseLight,
    italic: true, align: "left"
  });

  const bullets13 = [
    "Aplica todos los métodos anteriores sobre múltiples obras simultáneamente",
    "Resultados agregados: distribución de intervalos, perfiles tonales, frecuencia de cadencias",
    "Comparación estadística entre compositores, géneros y períodos",
  ];
  bullets13.forEach((b, i) => {
    let by = 1.72 + i * 0.65;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: by + 0.12, w: 0.22, h: 0.22,
      fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
    });
    slide.addText(b, {
      x: 0.9, y: by, w: 8.8, h: 0.52,
      fontSize: 14, fontFace: "Calibri", color: C.cream,
      align: "left", valign: "middle"
    });
  });

  // Two stat callouts
  [
    { title: '"Close Reading"', sub: "Análisis detallado de una obra" },
    { title: '"Distant Reading"', sub: "Patrones en cientos de obras a la vez" },
  ].forEach((c, i) => {
    let bx = 0.5 + i * 4.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: 3.92, w: 4.3, h: 1.1,
      fill: { color: i === 0 ? C.burgundy : C.medBurgundy },
      line: { color: C.dustyRose, width: 1 }
    });
    slide.addText(c.title, {
      x: bx + 0.1, y: 3.97, w: 4.1, h: 0.5,
      fontSize: 17, fontFace: "Georgia", color: C.cream,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(c.sub, {
      x: bx + 0.1, y: 4.47, w: 4.1, h: 0.45,
      fontSize: 12, fontFace: "Calibri", color: C.roseLight,
      align: "center", valign: "middle", margin: 0
    });
  });
}

// ============================================================
// SLIDE 14 - Formatos de Archivo
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Formatos de Archivo Compatibles");

  const cards14 = [
    {
      color: C.burgundy,
      title: "MEI",
      lines: ["Music Encoding Initiative", "Estándar académico internacional", "Usado por el corpus CRIM"],
      x: 0.25, isDark: true
    },
    {
      color: C.dustyRose,
      title: "MusicXML",
      lines: ["Formato universal de notación", "Compatible con Sibelius, Finale,", "MuseScore, Dorico"],
      x: 3.6, isDark: true
    },
    {
      color: C.lightCream,
      title: "MIDI",
      lines: ["Formato básico de datos MIDI", "Para análisis exploratorio inicial", ""],
      x: 6.88, isDark: false, border: true
    },
  ];

  cards14.forEach((card) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: card.x, y: 1.2, w: 3.1, h: 3.05,
      fill: { color: card.color },
      line: card.border ? { color: C.dustyRose, width: 2 } : { color: card.color, width: 0 },
      shadow: { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 }
    });

    // Top icon circle
    let iconBg = card.isDark ? C.darkBurgundy : C.dustyRose;
    slide.addShape(pres.shapes.OVAL, {
      x: card.x + 1.05, y: 1.35, w: 1.0, h: 1.0,
      fill: { color: iconBg }, line: { color: iconBg, width: 0 }
    });
    slide.addText(["M", "X", "~"][cards14.indexOf(card)], {
      x: card.x + 1.05, y: 1.35, w: 1.0, h: 1.0,
      fontSize: 22, fontFace: "Georgia", color: C.cream,
      align: "center", valign: "middle", bold: true, margin: 0
    });

    let tc = card.isDark ? C.cream : C.darkText;
    let ttc = card.isDark ? C.cream : C.burgundy;

    slide.addText(card.title, {
      x: card.x + 0.1, y: 2.45, w: 2.9, h: 0.5,
      fontSize: 22, fontFace: "Georgia", color: ttc,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    card.lines.forEach((line, li) => {
      if (line) {
        slide.addText(line, {
          x: card.x + 0.1, y: 3.0 + li * 0.32, w: 2.9, h: 0.3,
          fontSize: 12, fontFace: "Calibri", color: tc,
          align: "center", valign: "middle", margin: 0
        });
      }
    });
  });

  // Bottom note
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.25, y: 4.55, w: 9.5, h: 0.72,
    fill: { color: C.lightCream }, line: { color: C.dustyRose, width: 1 }
  });
  slide.addText("El corpus CRIM ofrece ediciones de alta calidad en formato MEI  ·  También puedes importar tus propias partituras", {
    x: 0.25, y: 4.55, w: 9.5, h: 0.72,
    fontSize: 12, fontFace: "Calibri", color: C.burgundy,
    italic: true, align: "center", valign: "middle", margin: 0
  });
}

// ============================================================
// SLIDE 15 - Flujo de Trabajo
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Flujo de Trabajo para Estudiantes");

  const steps15 = [
    { num: "1", title: "Acceder", desc: "intervals-streamlit.crimproject.org" },
    { num: "2", title: "Cargar", desc: "Una obra del corpus CRIM o tu propio archivo MEI/MusicXML" },
    { num: "3", title: "Seleccionar", desc: "La herramienta de análisis deseada" },
    { num: "4", title: "Configurar", desc: "Parámetros: tipo de intervalo, voces, n-grama, etc." },
    { num: "5", title: "Visualizar", desc: "Resultados en tablas y gráficos interactivos" },
    { num: "6", title: "Descargar", desc: "Datos en CSV para publicaciones académicas" },
  ];

  steps15.forEach((s, i) => {
    let col = i % 3, row = Math.floor(i / 3);
    let sx = 0.3 + col * 3.22, sy = 1.15 + row * 2.1;
    let bgColor = col % 2 === 0 ? C.burgundy : C.dustyRose;

    slide.addShape(pres.shapes.OVAL, {
      x: sx, y: sy, w: 0.75, h: 0.75,
      fill: { color: bgColor }, line: { color: bgColor, width: 0 }
    });
    slide.addText(s.num, {
      x: sx, y: sy, w: 0.75, h: 0.75,
      fontSize: 22, fontFace: "Georgia", color: C.cream,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.title, {
      x: sx, y: sy + 0.83, w: 3.0, h: 0.4,
      fontSize: 15, fontFace: "Georgia", color: C.burgundy,
      bold: true, align: "left"
    });
    slide.addText(s.desc, {
      x: sx, y: sy + 1.25, w: 3.0, h: 0.65,
      fontSize: 12, fontFace: "Calibri", color: C.darkText,
      align: "left", valign: "top"
    });
  });
}

// ============================================================
// SLIDE 16 - Aplicaciones
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Aplicaciones Docentes e Investigadoras", C.dustyRose);

  // Left column header
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 4.4, h: 0.45,
    fill: { color: C.burgundy }, line: { color: C.burgundy, width: 0 }
  });
  slide.addText("EN EL AULA", {
    x: 0.3, y: 1.1, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: "Georgia", color: C.cream,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  slide.addText([
    { text: "Análisis comparativo de motetes y sus Misas parodia", options: { bullet: true, breakLine: true } },
    { text: "Prácticas de análisis contrapuntístico asistido por ordenador", options: { bullet: true, breakLine: true } },
    { text: "Proyectos de investigación sobre estilo y cita musical", options: { bullet: true, breakLine: true } },
    { text: "Actividades de escucha activa con datos visuales", options: { bullet: true } },
  ], {
    x: 0.3, y: 1.62, w: 4.4, h: 2.6,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Decorative left
  slide.addShape(pres.shapes.OVAL, {
    x: 0.5, y: 4.1, w: 1.6, h: 1.4,
    fill: { color: C.burgundy, transparency: 80 },
    line: { color: C.dustyRose, width: 1.5 }
  });

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: 1.1, w: 0.06, h: 4.1,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  // Right column
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.28, y: 1.1, w: 4.4, h: 0.45,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });
  slide.addText("EN LA INVESTIGACIÓN", {
    x: 5.28, y: 1.1, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: "Georgia", color: C.cream,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  slide.addText([
    { text: "Identificación sistemática de citas y transformaciones", options: { bullet: true, breakLine: true } },
    { text: "Estudios de recepción y transmisión musical", options: { bullet: true, breakLine: true } },
    { text: "Análisis estadístico del estilo de un compositor", options: { bullet: true, breakLine: true } },
    { text: "Publicaciones con visualizaciones reproducibles", options: { bullet: true } },
  ], {
    x: 5.28, y: 1.62, w: 4.4, h: 2.6,
    fontSize: 13, fontFace: "Calibri", color: C.darkText,
    align: "left", valign: "top"
  });

  // Decorative right
  slide.addShape(pres.shapes.OVAL, {
    x: 8.1, y: 4.1, w: 1.6, h: 1.4,
    fill: { color: C.dustyRose, transparency: 80 },
    line: { color: C.burgundy, width: 1.5 }
  });
}

// ============================================================
// SLIDE 17 - Recursos
// ============================================================
{
  let slide = pres.addSlide();
  addLightBg(slide);
  addAccentBar(slide, 0, 5.625);
  addHeaderBand(slide, "Recursos y Comunidad");

  const resources17 = [
    { color: C.burgundy, title: "crimproject.org", desc: "Web principal del proyecto" },
    { color: C.dustyRose, title: "intervals-streamlit.crimproject.org", desc: "Aplicación web sin código" },
    { color: C.burgundy, title: "github.com/HCDigitalScholarship/intervals", desc: "Código fuente y tutoriales" },
    { color: C.dustyRose, title: "pip install crim_intervals", desc: "Biblioteca Python" },
    { color: C.burgundy, title: "+50 investigadores", desc: "Comunidad internacional activa" },
  ];

  resources17.forEach((r, i) => {
    let ry = 1.15 + i * 0.7;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.3, y: ry + 0.04, w: 0.52, h: 0.52,
      fill: { color: r.color }, line: { color: r.color, width: 0 }
    });
    slide.addText(String(i + 1), {
      x: 0.3, y: ry + 0.04, w: 0.52, h: 0.52,
      fontSize: 14, fontFace: "Georgia", color: C.cream,
      align: "center", valign: "middle", bold: true, margin: 0
    });
    slide.addText([
      { text: r.title, options: { bold: true } },
      { text: "  —  " + r.desc, options: {} }
    ], {
      x: 1.0, y: ry, w: 8.7, h: 0.6,
      fontSize: 14, fontFace: "Calibri", color: C.darkText,
      align: "left", valign: "middle"
    });
  });

  // Bottom callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.75, w: 9.4, h: 0.55,
    fill: { color: C.darkBurgundy }, line: { color: C.darkBurgundy, width: 0 }
  });
  slide.addText("Los tutoriales en Jupyter Notebooks están disponibles gratuitamente en GitHub", {
    x: 0.3, y: 4.75, w: 9.4, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.cream,
    italic: true, bold: true, align: "center", valign: "middle", margin: 0
  });
}

// ============================================================
// SLIDE 18 - CONCLUSION (dark)
// ============================================================
{
  let slide = pres.addSlide();
  addDarkBg(slide);
  addCornerOrnament(slide);

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.25, h: 5.625,
    fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
  });

  slide.addText("CRIM Intervals:", {
    x: 0.5, y: 0.28, w: 7.5, h: 0.65,
    fontSize: 38, fontFace: "Georgia", color: C.cream,
    bold: true, align: "left"
  });
  slide.addText("El Futuro del Análisis Musicológico", {
    x: 0.5, y: 0.93, w: 7.5, h: 0.55,
    fontSize: 22, fontFace: "Georgia", color: C.roseLight,
    italic: true, align: "left"
  });

  const bullets18 = [
    "Transforma el análisis de la polifonía renacentista",
    "Combina rigor analítico con accesibilidad (sin programación)",
    "Del análisis detallado al estudio masivo de corpus",
    "Herramienta ideal para estudiantes e investigadores",
  ];

  bullets18.forEach((b, i) => {
    let by = 1.65 + i * 0.6;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: by + 0.12, w: 0.22, h: 0.22,
      fill: { color: C.dustyRose }, line: { color: C.dustyRose, width: 0 }
    });
    slide.addText(b, {
      x: 0.9, y: by, w: 8.8, h: 0.5,
      fontSize: 15, fontFace: "Calibri", color: C.cream,
      align: "left", valign: "middle"
    });
  });

  // Large quote box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.02, w: 9.0, h: 0.88,
    fill: { color: C.burgundy }, line: { color: C.dustyRose, width: 1.5 }
  });
  slide.addText('"Las Humanidades Digitales al servicio del repertorio histórico"', {
    x: 0.5, y: 4.02, w: 9.0, h: 0.88,
    fontSize: 17, fontFace: "Georgia", color: C.cream,
    italic: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("intervals-streamlit.crimproject.org  ·  crimproject.org", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: C.roseLight,
    align: "center", valign: "middle"
  });
}

// Write file
pres.writeFile({ fileName: "C:/Users/MC.5055521/Desktop/Claude code/CRIM_Intervals_Musicologia.pptx" })
  .then(() => { console.log("Presentation created successfully!"); })
  .catch(err => { console.error("Error:", err); process.exit(1); });
