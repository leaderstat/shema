const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

const docsDir = path.join(__dirname, "..", "docs");
const svgPath = path.join(docsDir, "organizational-diagrams.svg");
const pptxPath = path.join(docsDir, "organizational-diagrams.pptx");

const requiredLabels = [
  "Директор технического департамента",
  "Главный конструктор",
  "Руководитель насосного подразделения",
  "Ведущий расчетчик",
  "Ведущий конструктор",
  "Штат конструкторов",
  "Главный технолог",
  "Технолог",
  "Материаловед",
  "Специалист комплексных решений",
  "КИП",
  "Руководитель производственного отдела",
  "Производственный отдел",
  "Полевой сервис",
  "СПБ сервисный центр",
  "Мурманский сервисный центр",
  "Сахалинский сервисный центр",
];

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

assert(fs.existsSync(svgPath), "SVG diagram is missing");
assert(fs.existsSync(pptxPath), "PPTX diagram is missing");

const svg = fs.readFileSync(svgPath, "utf8");
for (const label of requiredLabels) {
  assert(svg.includes(label), `SVG missing label: ${label}`);
}

const zip = new AdmZip(pptxPath);
const slide = zip.readAsText("ppt/slides/slide1.xml");
for (const label of requiredLabels) {
  assert(slide.includes(label), `PPTX missing label: ${label}`);
}

console.log(`Verified ${requiredLabels.length} labels in SVG and PPTX.`);
