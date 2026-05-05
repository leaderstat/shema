const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { ArrowHeads, StandardConnectionPoints, VisioDocument } = require("ts-visio");

const docsDir = path.join(__dirname, "..", "docs");
fs.mkdirSync(docsDir, { recursive: true });

const technical = [
  ["Директор технического департамента"],
  ["Главный конструктор", "Главный технолог", "Специалист комплексных решений"],
  ["Руководитель насосного подразделения", "Технолог", "КИП"],
  ["Ведущий расчетчик", "Материаловед"],
  ["Ведущий конструктор"],
  ["Штат конструкторов"],
];

const production = [
  ["Руководитель производственного отдела"],
  ["Производственный отдел"],
  ["Полевой сервис"],
  ["СПБ сервисный центр"],
  ["Мурманский сервисный центр"],
  ["Сахалинский сервисный центр"],
];

function svgText(text, x, y, size = 16, weight = 600) {
  return `<text x="${x}" y="${y}" text-anchor="middle" font-family="Arial, sans-serif" font-size="${size}" font-weight="${weight}" fill="#1f2937">${text}</text>`;
}

function boxSvg(text, x, y, w, h, fill = "#f8fafc", stroke = "#2563eb") {
  const lines = text.split(" ");
  const chunks = [];
  let current = "";
  for (const word of lines) {
    if ((current + " " + word).trim().length > 22) {
      chunks.push(current);
      current = word;
    } else {
      current = `${current} ${word}`.trim();
    }
  }
  if (current) chunks.push(current);

  const textLines = chunks
    .map((line, index) => `<tspan x="${x + w / 2}" dy="${index === 0 ? 0 : 18}">${line}</tspan>`)
    .join("");

  return [
    `<desc>${text}</desc>`,
    `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="8" fill="${fill}" stroke="${stroke}" stroke-width="2"/>`,
    `<text x="${x + w / 2}" y="${y + h / 2 - (chunks.length - 1) * 9 + 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" font-weight="600" fill="#111827">${textLines}</text>`,
  ].join("\n");
}

function lineSvg(x1, y1, x2, y2) {
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#64748b" stroke-width="2" marker-end="url(#arrow)"/>`;
}

function buildSvg() {
  const items = [];
  items.push(`<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">`);
  items.push(`<defs><marker id="arrow" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#64748b"/></marker></defs>`);
  items.push(`<rect width="1600" height="900" fill="#ffffff"/>`);
  items.push(svgText("Блок-схема 1. Структура технического департамента", 430, 55, 24, 700));
  items.push(svgText("Блок-схема 2. Производственный отдел", 1210, 55, 24, 700));

  const techBoxes = [
    ["Директор технического департамента", 275, 90, 310, 62],
    ["Главный конструктор", 80, 210, 220, 58],
    ["Главный технолог", 350, 210, 220, 58],
    ["Специалист комплексных решений", 620, 210, 250, 58],
    ["Руководитель насосного подразделения", 80, 315, 220, 62],
    ["Технолог", 350, 315, 220, 58],
    ["КИП", 635, 315, 220, 58],
    ["Ведущий расчетчик", 80, 420, 220, 58],
    ["Материаловед", 350, 420, 220, 58],
    ["Ведущий конструктор", 80, 525, 220, 58],
    ["Штат конструкторов", 80, 630, 220, 58],
  ];
  for (const [label, x, y, w, h] of techBoxes) items.push(boxSvg(label, x, y, w, h));
  items.push(lineSvg(430, 152, 190, 210));
  items.push(lineSvg(430, 152, 460, 210));
  items.push(lineSvg(430, 152, 745, 210));
  items.push(lineSvg(190, 268, 190, 315));
  items.push(lineSvg(460, 268, 460, 315));
  items.push(lineSvg(745, 268, 745, 315));
  items.push(lineSvg(190, 377, 190, 420));
  items.push(lineSvg(460, 373, 460, 420));
  items.push(lineSvg(190, 478, 190, 525));
  items.push(lineSvg(190, 583, 190, 630));

  const prodBoxes = [
    ["Руководитель производственного отдела", 1085, 90, 300, 62],
    ["Производственный отдел", 1105, 210, 260, 58],
    ["Полевой сервис", 1105, 315, 260, 58],
    ["СПБ сервисный центр", 1105, 420, 260, 58],
    ["Мурманский сервисный центр", 1105, 525, 260, 58],
    ["Сахалинский сервисный центр", 1105, 630, 260, 58],
  ];
  for (const [label, x, y, w, h] of prodBoxes) items.push(boxSvg(label, x, y, w, h, "#f0fdf4", "#16a34a"));
  for (let i = 0; i < prodBoxes.length - 1; i++) {
    const [, x1, y1, w1, h1] = prodBoxes[i];
    const [, x2, y2, w2] = prodBoxes[i + 1];
    items.push(lineSvg(x1 + w1 / 2, y1 + h1, x2 + w2 / 2, y2));
  }

  items.push(`</svg>`);
  return items.join("\n");
}

function addBox(pptxDoc, slide, text, x, y, w, h, color = "2563EB", fill = "F8FAFC") {
  slide.addShape(pptxDoc.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: fill },
    line: { color, width: 1.25 },
  });
  slide.addText(text, {
    x: x + 0.05,
    y: y + 0.08,
    w: w - 0.1,
    h: h - 0.16,
    fontFace: "Arial",
    fontSize: 11,
    bold: true,
    color: "111827",
    align: "center",
    valign: "mid",
    fit: "shrink",
  });
}

function addArrow(pptxDoc, slide, x1, y1, x2, y2) {
  slide.addShape(pptxDoc.ShapeType.line, {
    x: x1,
    y: y1,
    w: x2 - x1,
    h: y2 - y1,
    line: { color: "64748B", width: 1.25, beginArrowType: "none", endArrowType: "triangle" },
  });
}

function bottom(box) {
  return [box[1] + box[3] / 2, box[2] + box[4]];
}

function top(box) {
  return [box[1] + box[3] / 2, box[2]];
}

function addDownArrow(pptxDoc, slide, from, to) {
  addArrow(pptxDoc, slide, ...bottom(from), ...top(to));
}

async function buildPptx() {
  const pptxDoc = new pptxgen();
  pptxDoc.layout = "LAYOUT_WIDE";
  pptxDoc.author = "OpenAI Codex";
  pptxDoc.subject = "Организационные блок-схемы";
  pptxDoc.title = "Структура технического департамента и производственного отдела";
  pptxDoc.company = "leaderstat/shema";
  pptxDoc.lang = "ru-RU";
  pptxDoc.theme = {
    headFontFace: "Arial",
    bodyFontFace: "Arial",
    lang: "ru-RU",
  };

  const slide = pptxDoc.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText("Блок-схема 1. Структура технического департамента", {
    x: 0.35,
    y: 0.25,
    w: 6.9,
    h: 0.35,
    fontFace: "Arial",
    fontSize: 17,
    bold: true,
    color: "1F2937",
    align: "center",
  });
  slide.addText("Блок-схема 2. Производственный отдел", {
    x: 8.1,
    y: 0.25,
    w: 4.8,
    h: 0.35,
    fontFace: "Arial",
    fontSize: 17,
    bold: true,
    color: "1F2937",
    align: "center",
  });

  const tb = {
    director: [2.55, 0.85, 3, 0.62],
    constructor: [0.45, 1.95, 2.35, 0.62],
    technologistHead: [3.05, 1.95, 2.25, 0.62],
    complex: [5.75, 1.95, 2.55, 0.62],
    pump: [0.35, 2.98, 2.55, 0.68],
    technologist: [3.05, 2.98, 2.25, 0.62],
    kip: [5.9, 2.98, 2.25, 0.62],
    calc: [0.35, 4.02, 2.55, 0.62],
    material: [3.05, 4.02, 2.25, 0.62],
    leadConstructor: [0.35, 5.05, 2.55, 0.62],
    staff: [0.35, 6.08, 2.55, 0.62],
  };
  for (const [key, label] of [
    ["director", "Директор технического департамента"],
    ["constructor", "Главный конструктор"],
    ["technologistHead", "Главный технолог"],
    ["complex", "Специалист комплексных решений"],
    ["pump", "Руководитель насосного подразделения"],
    ["technologist", "Технолог"],
    ["kip", "КИП"],
    ["calc", "Ведущий расчетчик"],
    ["material", "Материаловед"],
    ["leadConstructor", "Ведущий конструктор"],
    ["staff", "Штат конструкторов"],
  ]) addBox(pptxDoc, slide, label, ...tb[key]);
  addDownArrow(pptxDoc, slide, tb.director, tb.constructor);
  addDownArrow(pptxDoc, slide, tb.director, tb.technologistHead);
  addDownArrow(pptxDoc, slide, tb.director, tb.complex);
  addDownArrow(pptxDoc, slide, tb.constructor, tb.pump);
  addDownArrow(pptxDoc, slide, tb.technologistHead, tb.technologist);
  addDownArrow(pptxDoc, slide, tb.complex, tb.kip);
  addDownArrow(pptxDoc, slide, tb.pump, tb.calc);
  addDownArrow(pptxDoc, slide, tb.technologist, tb.material);
  addDownArrow(pptxDoc, slide, tb.calc, tb.leadConstructor);
  addDownArrow(pptxDoc, slide, tb.leadConstructor, tb.staff);

  const pb = [
    ["Руководитель производственного отдела", 9.05, 0.85, 3, 0.62],
    ["Производственный отдел", 9.25, 1.95, 2.6, 0.62],
    ["Полевой сервис", 9.25, 2.98, 2.6, 0.62],
    ["СПБ сервисный центр", 9.25, 4.02, 2.6, 0.62],
    ["Мурманский сервисный центр", 9.25, 5.05, 2.6, 0.62],
    ["Сахалинский сервисный центр", 9.25, 6.08, 2.6, 0.62],
  ];
  for (const [label, x, y, w, h] of pb) addBox(pptxDoc, slide, label, x, y, w, h, "16A34A", "F0FDF4");
  for (let i = 0; i < pb.length - 1; i++) {
    addDownArrow(pptxDoc, slide, pb[i], pb[i + 1]);
  }

  await pptxDoc.writeFile({ fileName: path.join(docsDir, "organizational-diagrams.pptx") });
}

async function buildVsdx() {
  const doc = await VisioDocument.create();
  doc.setMetadata({
    title: "Структура технического департамента и производственного отдела",
    author: "OpenAI Codex",
    company: "leaderstat/shema",
    description: "Редактируемые организационные блок-схемы для Microsoft Visio",
    keywords: "Visio VSDX organizational diagrams",
  });

  const page = doc.pages[0];
  doc.renamePage(page, "Организационные схемы");
  page.setSize(14.2, 7.5);

  await page.addShape({
    text: "Блок-схема 1. Структура технического департамента",
    x: 3.63,
    y: 7.05,
    width: 5.9,
    height: 0.3,
    lineColor: "#FFFFFF",
    fillColor: "#FFFFFF",
    fontColor: "#1F2937",
    fontFamily: "Arial",
    fontSize: 17,
    bold: true,
    horzAlign: "center",
    verticalAlign: "middle",
  });
  await page.addShape({
    text: "Блок-схема 2. Производственный отдел",
    x: 10.55,
    y: 7.05,
    width: 4.9,
    height: 0.3,
    lineColor: "#FFFFFF",
    fillColor: "#FFFFFF",
    fontColor: "#1F2937",
    fontFamily: "Arial",
    fontSize: 17,
    bold: true,
    horzAlign: "center",
    verticalAlign: "middle",
  });

  async function addVisioBox(label, x, y, w, h, lineColor = "#2563EB", fillColor = "#F8FAFC") {
    return page.addShape({
      text: label,
      x,
      y,
      width: w,
      height: h,
      geometry: "rounded-rectangle",
      cornerRadius: 0.08,
      fillColor,
      lineColor,
      fontColor: "#111827",
      fontFamily: "Arial",
      fontSize: 11,
      bold: true,
      horzAlign: "center",
      verticalAlign: "middle",
      textMarginLeft: 0.04,
      textMarginRight: 0.04,
      textMarginTop: 0.04,
      textMarginBottom: 0.04,
      connectionPoints: StandardConnectionPoints.cardinal,
    });
  }

  async function connect(from, to) {
    await page.connectShapes(from, to, ArrowHeads.None, ArrowHeads.Standard, {
      lineColor: "#64748B",
      lineWeight: 1.25,
      routing: "orthogonal",
    }, { name: "Bottom" }, { name: "Top" });
  }

  const tech = {
    director: await addVisioBox("Директор технического департамента", 4.05, 6.35, 3, 0.62),
    constructor: await addVisioBox("Главный конструктор", 1.6, 5.2, 2.35, 0.62),
    technologistHead: await addVisioBox("Главный технолог", 4.05, 5.2, 2.25, 0.62),
    complex: await addVisioBox("Специалист комплексных решений", 6.8, 5.2, 2.55, 0.62),
    pump: await addVisioBox("Руководитель насосного подразделения", 1.6, 4.1, 2.55, 0.68),
    technologist: await addVisioBox("Технолог", 4.05, 4.1, 2.25, 0.62),
    kip: await addVisioBox("КИП", 6.95, 4.1, 2.25, 0.62),
    calc: await addVisioBox("Ведущий расчетчик", 1.6, 3, 2.55, 0.62),
    material: await addVisioBox("Материаловед", 4.05, 3, 2.25, 0.62),
    leadConstructor: await addVisioBox("Ведущий конструктор", 1.6, 1.9, 2.55, 0.62),
    staff: await addVisioBox("Штат конструкторов", 1.6, 0.8, 2.55, 0.62),
  };
  await connect(tech.director, tech.constructor);
  await connect(tech.director, tech.technologistHead);
  await connect(tech.director, tech.complex);
  await connect(tech.constructor, tech.pump);
  await connect(tech.technologistHead, tech.technologist);
  await connect(tech.complex, tech.kip);
  await connect(tech.pump, tech.calc);
  await connect(tech.technologist, tech.material);
  await connect(tech.calc, tech.leadConstructor);
  await connect(tech.leadConstructor, tech.staff);

  const production = [
    await addVisioBox("Руководитель производственного отдела", 11.3, 6.35, 3, 0.62, "#16A34A", "#F0FDF4"),
    await addVisioBox("Производственный отдел", 11.3, 5.2, 2.6, 0.62, "#16A34A", "#F0FDF4"),
    await addVisioBox("Полевой сервис", 11.3, 4.1, 2.6, 0.62, "#16A34A", "#F0FDF4"),
    await addVisioBox("СПБ сервисный центр", 11.3, 3, 2.6, 0.62, "#16A34A", "#F0FDF4"),
    await addVisioBox("Мурманский сервисный центр", 11.3, 1.9, 2.6, 0.62, "#16A34A", "#F0FDF4"),
    await addVisioBox("Сахалинский сервисный центр", 11.3, 0.8, 2.6, 0.62, "#16A34A", "#F0FDF4"),
  ];
  for (let i = 0; i < production.length - 1; i++) {
    await connect(production[i], production[i + 1]);
  }

  await doc.save(path.join(docsDir, "organizational-diagrams.vsdx"));
}

fs.writeFileSync(path.join(docsDir, "organizational-diagrams.svg"), buildSvg(), "utf8");
Promise.all([buildPptx(), buildVsdx()]).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
