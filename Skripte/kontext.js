const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "Kontextdiagramm ThesisFlow";
const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF", GREY = "8A968A", LINE = "6B7F6B";

const s = pres.addSlide();
s.background = { color: WHITE };

s.addText("Kontextdiagramm ThesisFlow", { x: 0.45, y: 0.2, w: 6.5, h: 0.45, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 22, bold: true, color: INK, valign: "middle" });
s.addText("Systemkontext: Akteure, Nachbarsysteme und Datenflüsse", { x: 0.45, y: 0.62, w: 6.5, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10, color: MUTED });
s.addText("ABSCHLUSSARBEITEN-PORTAL MND  ·  FB MND, THM", { x: 6.0, y: 0.28, w: 3.55, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

// Kontextgrenze
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.35, y: 0.95, w: 9.3, h: 4.0, fill: { color: WHITE }, line: { color: MOSS, width: 1, dashType: "dash" }, rectRadius: 0.15 });
s.addText("Kontextgrenze", { x: 8.2, y: 1.0, w: 1.35, h: 0.2, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, italic: true, color: MOSS, align: "right" });

// System (Blackbox)
const C = { x: 3.9, y: 2.1, w: 2.2, h: 1.3 };
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: C.x, y: C.y, w: C.w, h: C.h, fill: { color: GREEN }, line: { color: GREEN }, rectRadius: 0.1, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
s.addText([
  { text: "ThesisFlow", options: { fontFace: "Cambria", fontSize: 18, bold: true, color: WHITE, breakLine: true } },
  { text: "Abschlussarbeiten-Portal", options: { fontFace: "Calibri", fontSize: 9, color: "D9E6CF", breakLine: true } },
  { text: "Systemgrenze", options: { fontFace: "Calibri", fontSize: 7.5, italic: true, color: MOSS } },
], { x: C.x, y: C.y, w: C.w, h: C.h, margin: 0, isTextBox: true, align: "center", valign: "middle" });

// Akteure und Systeme
const AW = 1.7, AH = 0.55;
const actor = (x, y, label, sys) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: AW, h: AH, fill: { color: sys ? WHITE : TINT }, line: { color: sys ? GREY : GREEN, width: sys ? 1 : 1.25, dashType: sys ? "dash" : "solid" }, rectRadius: 0.08 });
  s.addText(label, { x, y, w: AW, h: AH, margin: 0.04, isTextBox: true, fontFace: "Calibri", fontSize: 9.5, bold: true, color: sys ? MUTED : GREEN, align: "center", valign: "middle" });
};
const left = [["Studierende", 1.05], ["Erstprüfende", 2.475], ["Sekretariat / Prüfungsamt", 3.9]];
const right = [["Externe Korreferenten", 1.05], ["Interne Zweitprüfende", 2.475], ["Prüfungsausschuss", 3.9]];
left.forEach(a => actor(0.5, a[1], a[0], false));
right.forEach(a => actor(7.8, a[1], a[0], false));
actor(4.15, 1.0, "THM-Login (SSO)", true);
actor(4.15, 4.2, "E-Mail-Dienst", true);

// Verbindungen (Doppelpfeile)
const link = (x1, y1, x2, y2) => {
  const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
  const up = (x2 > x1 && y2 < y1) || (x2 < x1 && y2 > y1);
  s.addShape(pres.shapes.LINE, { x, y, w, h: h || 0.001, flipV: up, line: { color: LINE, width: 1.25, beginArrowType: "triangle", endArrowType: "triangle" } });
};
const cy = C.y + C.h / 2;
left.forEach(a => link(0.5 + AW, a[1] + AH / 2, C.x, cy));
right.forEach(a => link(7.8, a[1] + AH / 2, C.x + C.w, cy));
link(5.0, 1.0 + AH, 5.0, C.y);
link(5.0, C.y + C.h, 5.0, 4.2);

// Datenfluss-Labels (→ ins System, ← aus dem System)
const label = (x, y, w, into, out) => {
  s.addText([
    { text: "→ " + into, options: { breakLine: true } },
    { text: "← " + out },
  ], { x, y, w, h: 0.5, margin: 0.03, isTextBox: true, fontFace: "Calibri", fontSize: 7, color: INK, fill: { color: WHITE }, valign: "middle", lineSpacingMultiple: 1.05 });
};
label(2.3, 1.62, 1.5, "Antrag, Anlage ext. Prüfer, Urkundenkopie, Verlängerungsantrag", "Status, Rückfragen, Zulassungsbescheid");
label(2.3, 2.5, 1.5, "Themenbestätigung, Freigabe, Stellungnahme Verlängerung", "Betreuungsanfrage, Statusinfo");
label(2.3, 3.4, 1.5, "Vollständigkeitsprüfung, Fristen, Weiterleitung", "Antragsdaten, Fristenliste, Exporte");
label(6.2, 1.62, 1.5, "Qualifikationsnachweis, Zusage Korreferat", "Einladungslink, Antragsdaten");
label(6.2, 2.5, 1.5, "Zusage Korreferat", "Anfrage Korreferat, Statusinfo");
label(6.2, 3.4, 1.5, "Zulassungsentscheidung, Genehmigung Verlängerung, Anerkennung ext. Prüfer", "Entscheidungsvorlage");
label(5.1, 1.5, 1.45, "Identität, Rolle, Matrikelnummer", "Authentifizierungsanfrage");
label(5.1, 3.55, 1.45, "Zustellbestätigung", "Benachrichtigungen, Einladungen");

// Legende
const ly = 5.1;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.45, y: ly + 0.04, w: 0.3, h: 0.18, fill: { color: TINT }, line: { color: GREEN, width: 1 }, rectRadius: 0.04 });
s.addText("Akteur (Person / Rolle)", { x: 0.8, y: ly, w: 1.5, h: 0.26, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.3, y: ly + 0.04, w: 0.3, h: 0.18, fill: { color: WHITE }, line: { color: GREY, width: 1, dashType: "dash" }, rectRadius: 0.04 });
s.addText("Nachbarsystem", { x: 2.65, y: ly, w: 1.1, h: 0.26, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.LINE, { x: 3.85, y: ly + 0.13, w: 0.35, h: 0.001, line: { color: LINE, width: 1.25, beginArrowType: "triangle", endArrowType: "triangle" } });
s.addText("Datenfluss  (→ ins System, ← aus dem System)", { x: 4.25, y: ly, w: 2.6, h: 0.26, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addText("Requirements Engineering – Praxisteil  ·  Kontextdiagramm", { x: 6.9, y: ly, w: 2.65, h: 0.26, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: GREY, align: "right", valign: "middle" });

s.addNotes("Kontextdiagramm ThesisFlow: Das System ist eine Blackbox (Systemgrenze). Innerhalb der Kontextgrenze liegen die Akteure Studierende, Erstprüfende, interne Zweitprüfende, externe Korreferenten, Sekretariat/Prüfungsamt und Prüfungsausschuss sowie die Nachbarsysteme THM-Login (SSO) und E-Mail-Dienst. Pfeile zeigen Datenflüsse in beide Richtungen. Quellen: Projektbriefing, Anlage externer Prüfer (§ 18 Abs. 2 HHG, fünf Jahre Berufspraxis), PO B.Sc. § 6, PO M.Sc.");

const out = process.env.OUT || "Kontextdiagramm_ThesisFlow.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
