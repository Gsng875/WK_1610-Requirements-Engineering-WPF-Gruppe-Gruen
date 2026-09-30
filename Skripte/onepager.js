const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "One-Pager ThesisFlow";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF";
const s = pres.addSlide();
s.background = { color: WHITE };

// ---------- linke Spalte (dunkel) ----------
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.7, h: 5.625, fill: { color: GREEN }, line: { color: GREEN } });

s.addText("ABSCHLUSSARBEITEN-PORTAL  ·  FACHBEREICH MND", {
  x: 0.45, y: 0.4, w: 2.9, h: 0.3, margin: 0, isTextBox: true,
  fontFace: "Calibri", fontSize: 8.5, color: MOSS, charSpacing: 1.5, bold: true,
});
s.addText("ThesisFlow", {
  x: 0.45, y: 0.75, w: 3.0, h: 0.7, margin: 0, isTextBox: true,
  fontFace: "Cambria", fontSize: 34, bold: true, color: WHITE,
});
s.addText("Anmelden. Betreuen.\nAbschließen.", {
  x: 0.45, y: 1.5, w: 3.0, h: 0.85, margin: 0, isTextBox: true,
  fontFace: "Cambria", fontSize: 19, italic: true, color: MOSS, valign: "top",
});

s.addText("UNSERE VISION", {
  x: 0.45, y: 2.6, w: 2.9, h: 0.25, margin: 0, isTextBox: true,
  fontFace: "Calibri", fontSize: 8.5, color: MOSS, charSpacing: 1.5, bold: true,
});
s.addText(
  "Jede Abschlussarbeit im Fachbereich MND wird vollständig digital angemeldet, freigegeben und verwaltet – ohne Papier, ohne Medienbrüche, von jedem Gerät aus. Studierende, Prüfende und Prüfungsamt sehen jederzeit, wo ihr Antrag steht.",
  { x: 0.45, y: 2.88, w: 2.9, h: 1.6, margin: 0, isTextBox: true,
    fontFace: "Calibri", fontSize: 11.5, color: WHITE, valign: "top", lineSpacingMultiple: 1.15 }
);

s.addText([
  { text: "Auftraggeber: ", options: { bold: true } },
  { text: "Fachbereich MND, THM\n" },
  { text: "Auftragnehmer: ", options: { bold: true } },
  { text: "Projektteam Requirements Engineering" },
], { x: 0.45, y: 4.7, w: 3.0, h: 0.55, margin: 0, isTextBox: true,
     fontFace: "Calibri", fontSize: 8.5, color: "D9E6CF", valign: "bottom" });

// ---------- rechte Seite ----------
s.addText("Elevator Pitch", {
  x: 4.0, y: 0.38, w: 5.5, h: 0.45, margin: 0, isTextBox: true,
  fontFace: "Cambria", fontSize: 22, bold: true, color: INK,
});

const cards = [
  ["Für wen", "Studierende, Erst- und Zweitprüfende – auch externe Korreferenten – sowie das Prüfungsamt des FB MND."],
  ["Das Problem", "Anmeldung per PDF-Antrag mit Unterschriften und E-Mail-Umläufen: langsam, fehleranfällig, ohne Statusübersicht."],
  ["Unsere Lösung", "Ein webbasiertes Portal, das den Antrag digital führt, auf Vollständigkeit prüft und an alle Beteiligten weiterleitet."],
  ["Der Nutzen", "Weniger Rückfragen, kürzere Durchlaufzeit und transparenter Status – am Desktop wie am Smartphone."],
  ["Anders als das PDF", "Das Portal kennt die Prüfungsordnungen (B.Sc./M.Sc.) und bindet externe Korreferenten direkt in den Prozess ein."],
  ["Ausblick", "Ausbaustufe 2: Verlängerungsanträge digital stellen, prüfen und genehmigen – auf derselben Plattform."],
];
const cw = 1.72, ch = 1.32, gap = 0.15, x0 = 4.0, y0 = 0.95;
cards.forEach((c, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  const x = x0 + col * (cw + gap), y = y0 + row * (ch + gap);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: TINT }, line: { color: TINT }, rectRadius: 0.08 });
  s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: y + 0.12, w: 0.28, h: 0.28, fill: { color: GREEN }, line: { color: GREEN } });
  s.addText(String(i + 1), { x: x + 0.12, y: y + 0.12, w: 0.28, h: 0.28, margin: 0, isTextBox: true,
    fontFace: "Calibri", fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle" });
  s.addText(c[0], { x: x + 0.48, y: y + 0.12, w: cw - 0.58, h: 0.28, margin: 0, isTextBox: true,
    fontFace: "Calibri", fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  s.addText(c[1], { x: x + 0.12, y: y + 0.48, w: cw - 0.24, h: ch - 0.56, margin: 0, isTextBox: true,
    fontFace: "Calibri", fontSize: 8.5, color: INK, valign: "top", lineSpacingMultiple: 1.1 });
});

// ---------- Kennzahlen-Zeile ----------
const stats = [
  ["1", "Portal statt drei\nPDF-Formulare"],
  ["2", "Ausbaustufen: Anmeldung,\ndann Verlängerung"],
  ["100 %", "nutzbar auf Desktop\nund Smartphone"],
];
stats.forEach((st, i) => {
  const x = x0 + i * (cw + gap), y = 3.95;
  s.addText(st[0], { x, y, w: 0.95, h: 0.6, margin: 0, isTextBox: true,
    fontFace: "Cambria", fontSize: 26, bold: true, color: GREEN, valign: "middle" });
  s.addText(st[1], { x: x + 0.95, y, w: cw - 0.95, h: 0.6, margin: 0, isTextBox: true,
    fontFace: "Calibri", fontSize: 8.5, color: MUTED, valign: "middle", lineSpacingMultiple: 1.05 });
});

s.addText("Requirements Engineering – Praxisteil  ·  Prof. Dr. Carsten Lucke  ·  One-Pager / Produktvision", {
  x: 4.0, y: 5.15, w: 5.5, h: 0.25, margin: 0, isTextBox: true,
  fontFace: "Calibri", fontSize: 8, color: "8A968A",
});

s.addNotes(
  "Elevator Pitch (30 Sek.): Für Studierende, Prüfende und das Prüfungsamt des Fachbereichs MND, die Abschlussarbeiten heute per PDF-Antrag, Unterschriften und E-Mail-Umläufen anmelden, ist ThesisFlow ein webbasiertes Portal, das Anmeldung, Freigabe und Verwaltung digital, mobil und nachvollziehbar macht. Anders als das PDF-Formular kennt es die Prüfungsordnungen, bindet externe Korreferenten direkt ein und wird in Stufe 2 um Verlängerungsanträge erweitert.\n\n" +
  "Alternative Taglines: „Deine Arbeit. Dein Antrag. Ein Klick.“ · „Vom Antrag zum Abschluss – digital.“ · „Weniger Formular, mehr Abschluss.“ · „Abschlussarbeiten. Einfach digital.“"
);

pres.writeFile({ fileName: "One-Pager_ThesisFlow.pptx" }).then(f => console.log("written", f));
