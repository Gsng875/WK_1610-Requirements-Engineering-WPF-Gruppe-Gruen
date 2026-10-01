// ThesisFlow – Abschlusspraesentation (Gesamtpraesentation vor dem Auftraggeber)
// Einheitliche Begriffe: ThesisFlow, Studierende:r, Erstpruefer:in, Zweitpruefer:in, Dekanat,
// Pruefungsausschussvorsitzende:r, THM-Login, Abschlussarbeit (zentrales Datenobjekt).
// Aufruf: OUT="../Abschlusspräsentation/ThesisFlow_Abschlusspraesentation.pptx" IMG="../Abschlusspräsentation/bilder" node abschluss.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "ThesisFlow – Abschlusspräsentation";
const IMG = process.env.IMG || path.join(__dirname, "..", "Abschlusspräsentation", "bilder");
const img = (n) => path.join(IMG, n);

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF",
  PALE = "D9E6CF", GREY = "8A968A", LGREY = "C9D1C9", BG = "F5F7F4", AMBER = "C9A227", AMBERPALE = "FFF7E0", BROWN = "7A5C00",
  BLUE = "4A7FB5", BLUEPALE = "DCE8F5";

// ---------- Helfer ----------
const T = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: "Calibri", color: INK, valign: "top" }, o));
const rbox = (s, x, y, w, h, fill, line, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 1 }, rectRadius: 0.08 }, o || {}));
const rect = (s, x, y, w, h, fill, line, lw) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: lw || 0.75 } });
const circle = (s, x, y, d, fill, line, lw) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: line || fill, width: lw || 1 } });
const seg = (s, x1, y1, x2, y2, head, color, width, dash) => {
  const horiz = Math.abs(y2 - y1) < 0.001;
  const line = { color: color || INK, width: width || 1.25 };
  if (dash) line.dashType = dash;
  if (head) { const end = horiz ? x2 > x1 : y2 > y1; if (end) line.endArrowType = "triangle"; else line.beginArrowType = "triangle"; }
  s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: horiz ? Math.abs(x2 - x1) : 0.001, h: horiz ? 0.001 : Math.abs(y2 - y1), line });
};
const flow = (s, pts, color) => { for (let i = 0; i < pts.length - 1; i++) seg(s, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], i === pts.length - 2, color); };
const bullets = (s, items, o) => T(s, items.map((t, i) => ({ text: t, options: { bullet: { indent: 11 }, breakLine: i < items.length - 1 } })), Object.assign({ fontSize: 9.5, paraSpaceAfter: 3 }, o));
const num = (s, x, y, n, d) => { d = d || 0.28; circle(s, x, y, d, GREEN); T(s, String(n), { x, y, w: d, h: d, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle" }); };
let page = 0;
const slide = (title, tag, sub) => {
  const s = pres.addSlide(); page++;
  s.background = { color: WHITE };
  T(s, title, { x: 0.45, y: 0.22, w: 7.0, h: 0.5, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  T(s, tag, { x: 6.6, y: 0.3, w: 2.95, h: 0.3, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });
  if (sub) T(s, sub, { x: 0.45, y: 0.72, w: 9.1, h: 0.22, fontSize: 9.5, color: MUTED });
  T(s, "ThesisFlow  ·  Gruppe Grün  ·  Auftraggeber: Fachbereich MND", { x: 0.45, y: 5.28, w: 5, h: 0.2, fontSize: 7.5, color: GREY });
  T(s, String(page + 1), { x: 9.05, y: 5.28, w: 0.5, h: 0.2, fontSize: 7.5, color: GREY, align: "right" });
  return s;
};
const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: GREEN }, fontSize: 8.5 } });
const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 8.5, color: INK }, o || {}) });
const b = (t) => c(t, { bold: true, color: GREEN, fill: { color: TINT } });
const table = (s, rows, o) => s.addTable(rows, Object.assign({ fontFace: "Calibri", border: { type: "solid", color: LGREY, pt: 0.75 }, valign: "middle", margin: [2, 5, 2, 5] }, o));

// BPMN-Bausteine
const pool = (s, y, lanes, name) => {
  const X = 0.45, W = 9.1, BAND = 0.28, H = lanes.reduce((a, l) => a + l.h, 0);
  rect(s, X, y, W, H, WHITE, INK, 1.25);
  const vt = (x, yy, bw, lh, text, bold) => T(s, text, { x: x + bw / 2 - lh / 2, y: yy + lh / 2 - bw / 2, w: lh, h: bw, rotate: 270, fontSize: 8.5, bold: !!bold, color: GREEN, align: "center", valign: "middle" });
  rect(s, X, y, BAND, H, TINT, INK, 1.25); vt(X, y, BAND, H, name, true);
  let yy = y; const cy = [];
  lanes.forEach((l, i) => {
    rect(s, X + BAND, yy, BAND, l.h, TINT, INK, 1); vt(X + BAND, yy, BAND, l.h, l.name);
    if (i > 0) s.addShape(pres.shapes.LINE, { x: X + BAND, y: yy, w: W - BAND, h: 0.001, line: { color: INK, width: 1 } });
    cy.push(yy + l.h / 2); yy += l.h;
  });
  return cy;
};
const TW = 1.25, TH = 0.5;
const task = (s, x, cy, text, sys) => { rbox(s, x, cy - TH / 2, TW, TH, sys ? BLUEPALE : WHITE, INK, { line: { color: INK, width: 1.25 }, rectRadius: 0.1 }); T(s, text, { x: x + 0.04, y: cy - TH / 2, w: TW - 0.08, h: TH, fontSize: 8, align: "center", valign: "middle" }); };
const startEv = (s, x, cy, label) => { circle(s, x, cy - 0.13, 0.26, WHITE, GREEN, 1.25); if (label) T(s, label, { x: x - 0.02, y: cy + 0.16, w: 0.52, h: 0.3, fontSize: 6.5, color: MUTED }); };
const endEv = (s, x, cy, label, below) => { circle(s, x, cy - 0.14, 0.28, WHITE, INK, 3); if (label) T(s, label, below ? { x: x - 0.46, y: cy + 0.3, w: 1.2, h: 0.2, fontSize: 7, bold: true, align: "center" } : { x: x + 0.36, y: cy - 0.2, w: 1.0, h: 0.4, fontSize: 7.5, bold: true, valign: "middle" }); };
const gw = (s, cx, cy, label) => { s.addShape(pres.shapes.DIAMOND, { x: cx - 0.2, y: cy - 0.2, w: 0.4, h: 0.4, fill: { color: AMBERPALE }, line: { color: INK, width: 1.25 } }); T(s, "X", { x: cx - 0.2, y: cy - 0.2, w: 0.4, h: 0.4, fontSize: 10, bold: true, align: "center", valign: "middle" }); if (label) T(s, label, { x: cx - 0.7, y: cy + 0.29, w: 1.4, h: 0.18, fontSize: 7, italic: true, color: MUTED, align: "center" }); };
const rules = (s, y, list) => {
  T(s, "Control: Regeln", { x: 0.45, y, w: 3, h: 0.24, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  list.forEach((r, i) => {
    const yy = y + 0.28 + i * 0.25;
    rbox(s, 0.45, yy, 0.38, 0.21, GREEN, GREEN, { rectRadius: 0.04 });
    T(s, "C" + (i + 1), { x: 0.45, y: yy, w: 0.38, h: 0.21, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, r, { x: 0.92, y: yy, w: 8.6, h: 0.21, fontSize: 8.5, valign: "middle" });
  });
};

// =====================================================================
// 1 Titel und Agenda
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 3.7, 5.625, GREEN);
  T(s, "PROJEKTPRÄSENTATION", { x: 0.45, y: 0.45, w: 3.0, h: 0.25, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, "ThesisFlow", { x: 0.45, y: 0.85, w: 3.0, h: 0.7, fontFace: "Cambria", fontSize: 36, bold: true, color: WHITE, valign: "middle" });
  T(s, "Anmelden. Betreuen.\nAbschließen.", { x: 0.45, y: 1.65, w: 3.0, h: 0.8, fontFace: "Cambria", fontSize: 18, italic: true, color: MOSS });
  T(s, "Webbasiertes Portal zur digitalen Anmeldung und Verwaltung von Abschlussarbeiten im Fachbereich MND", { x: 0.45, y: 2.75, w: 2.9, h: 0.9, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.1 });
  T(s, [
    { text: "Auftraggeber: ", options: { bold: true } }, { text: "Fachbereich MND, Prof. Dr. Carsten Lucke", options: { breakLine: true } },
    { text: "Auftragnehmer: ", options: { bold: true } }, { text: "Gruppe Grün", options: { breakLine: true } },
    { text: "Requirements Engineering, 02.10.2026" },
  ], { x: 0.45, y: 4.5, w: 3.0, h: 0.75, fontSize: 8.5, color: PALE, valign: "bottom", paraSpaceAfter: 2 });
  T(s, "Agenda", { x: 4.0, y: 0.42, w: 5.5, h: 0.45, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  const ag = ["Überblick und Produktvision", "System- und Kontextabgrenzung", "Stakeholder", "Story Map und Releaseplanung", "Gesamt-Datenmodell", "Dialoglandkarte", "Drei User Stories nach den 7 Dimensionen", "Testfälle und offene Punkte"];
  ag.forEach((a, i) => {
    const y = 1.02 + i * 0.5;
    rbox(s, 4.0, y, 5.55, 0.42, TINT);
    num(s, 4.1, y + 0.07, i + 1);
    T(s, a, { x: 4.5, y, w: 4.9, h: 0.42, fontSize: 11, valign: "middle" });
  });
}

// =====================================================================
// 2 Ueberblick und Produktvision
// =====================================================================
{
  const s = slide("Überblick und Produktvision", "1  ·  VISION");
  const cols = [
    ["Heute", ["Anmeldung per PDF-Antrag mit Ausdruck und Unterschriften", "Weitergabe per E-Mail, sieben Übergaben pro Antrag", "Kein Status für Studierende, Rückfragen nur telefonisch"], TINT, GREEN],
    ["Unsere Vision", ["Jede Abschlussarbeit im Fachbereich MND wird vollständig digital angemeldet, freigegeben und verwaltet.", "Ohne Papier, ohne Medienbrüche, von jedem Gerät aus."], GREEN, WHITE],
    ["Nutzen", ["Weniger Rückfragen und kürzere Durchlaufzeit", "Status für alle Beteiligten jederzeit sichtbar", "Komfortabel nutzbar am Desktop und am Smartphone"], TINT, GREEN],
  ];
  cols.forEach((col, i) => {
    const x = 0.45 + i * 3.075, w = 2.95, dark = col[3] === WHITE;
    rbox(s, x, 1.0, w, 2.85, col[2]);
    T(s, col[0], { x: x + 0.18, y: 1.1, w: w - 0.36, h: 0.36, fontFace: "Cambria", fontSize: 14, bold: true, color: dark ? MOSS : GREEN, valign: "middle" });
    bullets(s, col[1], { x: x + 0.2, y: 1.55, w: w - 0.38, h: 2.2, fontSize: 10.5, color: dark ? WHITE : INK, paraSpaceAfter: 6 });
  });
  rbox(s, 0.45, 4.05, 9.1, 1.05, AMBERPALE, AMBER);
  T(s, [
    { text: "Auftrag: ", options: { bold: true, color: BROWN } }, { text: "Der PDF-Antrag wird durch ein Webportal ersetzt. Externe Zweitprüfer:innen werden in die Anmeldung eingebunden.", options: { breakLine: true } },
    { text: "Zweite Ausbaustufe: ", options: { bold: true, color: BROWN } }, { text: "Verlängerungsanträge.", options: { breakLine: true } },
    { text: "Grundlage: ", options: { bold: true, color: BROWN } }, { text: "Antragsformulare der THM und Prüfungsordnungen B.Sc. und M.Sc. Wirtschaftsinformatik." },
  ], { x: 0.6, y: 4.05, w: 8.8, h: 1.05, fontSize: 9.5, valign: "middle", paraSpaceAfter: 3 });
}

// =====================================================================
// 3 System- und Kontextabgrenzung
// =====================================================================
{
  const s = slide("System- und Kontextabgrenzung", "2  ·  KONTEXT", "Soll-Perspektive: Wer und was mit ThesisFlow Daten austauscht");
  rbox(s, 0.45, 1.0, 9.1, 4.12, BG, MOSS, { line: { color: MOSS, width: 1.25, dashType: "dash" }, rectRadius: 0.12 });
  T(s, "Kontextgrenze", { x: 0.6, y: 1.04, w: 2, h: 0.18, fontSize: 7.5, italic: true, bold: true, color: MOSS });
  const C = { x: 4.1, y: 1.25, w: 1.8, h: 3.65 };
  rbox(s, C.x, C.y, C.w, C.h, GREEN, GREEN, { rectRadius: 0.1 });
  T(s, [{ text: "ThesisFlow", options: { fontFace: "Cambria", fontSize: 17, bold: true, color: WHITE, breakLine: true } }, { text: "Systemgrenze", options: { fontSize: 8, italic: true, color: MOSS } }], { x: C.x, y: C.y, w: C.w, h: C.h, align: "center", valign: "middle" });
  const ent = (x, y, name, sys) => { rbox(s, x, y, 1.9, 0.5, sys ? WHITE : TINT, sys ? GREY : GREEN, { line: { color: sys ? GREY : GREEN, width: 1.25, dashType: sys ? "dash" : "solid" } }); T(s, name, { x, y, w: 1.9, h: 0.5, fontSize: 9, bold: true, color: sys ? MUTED : GREEN, align: "center", valign: "middle" }); };
  const lab = (x, y, a, bb) => T(s, [{ text: "▶ " + a, options: { color: GREEN, breakLine: !!bb } }].concat(bb ? [{ text: "◀ " + bb, options: { color: BLUE } }] : []), { x, y, w: 1.42, h: 0.44, margin: 2, fontSize: 6.5, fill: { color: WHITE }, line: { color: LGREY, width: 0.5 }, valign: "middle" });
  const L = [["Studierende:r", "Antrag, Verlängerungsantrag", "Status, Abgabetermin"], ["Erstprüfer:in", "Themenfreigabe, Stellungnahme", "Themen, Aufgaben"], ["Zweitprüfer:in", "Zusage, Qualifikationsnachweis", "Anfrage"]];
  L.forEach((e, i) => { const y = 1.45 + i * 1.25, yc = y + 0.25; ent(s ? 0.6 : 0, y, e[0]); seg(s, 2.5, yc - 0.06, C.x, yc - 0.06, true, GREEN); seg(s, C.x, yc + 0.06, 2.5, yc + 0.06, true, BLUE); lab(2.59, yc - 0.22, e[1], e[2]); });
  const R = [["Dekanat", "Prüfergebnis, Fristprüfung", "Anträge, Fristen", false], ["Prüfungsausschuss-\nvorsitzende:r", "Entscheidung", "Entscheidungsvorlage", false], ["THM-Login", "Identität, Rolle", "Anmeldeanfrage", true], ["E-Mail-Dienst", null, "Benachrichtigungen", true]];
  R.forEach((e, i) => { const y = 1.35 + i * 0.9, yc = y + 0.25; ent(7.5, y, e[0], e[3]); if (e[1]) seg(s, 7.5, yc - 0.06, C.x + C.w, yc - 0.06, true, GREEN); seg(s, C.x + C.w, yc + 0.06, 7.5, yc + 0.06, true, BLUE); if (e[1]) lab(5.99, yc - 0.22, e[1], e[2]); else T(s, [{ text: "◀ " + e[2], options: { color: BLUE } }], { x: 5.99, y: yc - 0.14, w: 1.42, h: 0.28, margin: 2, fontSize: 6.5, fill: { color: WHITE }, line: { color: LGREY, width: 0.5 }, valign: "middle" }); });
  T(s, "▶ Eingabe   ◀ Ausgabe   gestrichelt: Nachbarsystem", { x: 5.2, y: 5.28, w: 3.7, h: 0.2, fontSize: 7.5, color: GREY, align: "right" });
}

// =====================================================================
// 4 Stakeholder
// =====================================================================
{
  const s = slide("Stakeholder", "3  ·  STAKEHOLDER", "Drei Personas für die Rollen, die ThesisFlow nutzen");
  const P = [
    { name: "Moritz Hoffmann", role: "Studierende:r", photo: "bachelor.png", full: true, ziel: "Antrag in einem Durchgang einreichen, Status jederzeit sehen, Verlängerung online beantragen", pain: "Wochenlang keine Rückmeldung zum Antrag, Papierformulare und unklare Fristen" },
    { name: "Prof. Dr. Andreas Becker", role: "Erstprüfer:in", initials: "AB", ziel: "Thema prüfen und freigeben, ohne Formular und Unterschrift", pain: "Formular per E-Mail hin- und herschicken" },
    { name: "Michael Krüger", role: "Dekanat", photo: "sekretariat.png", ziel: "Vollständigkeit und Fristen auf einen Blick prüfen", pain: "Unleserliche Scans, tägliche Telefonrückfragen" },
  ];
  P.forEach((p, i) => {
    const x = 0.45 + i * 3.075, w = 2.95, y = 1.02;
    rbox(s, x, y, w, 3.35, TINT);
    const d = 1.05;
    if (p.full) { const h = 1.15, ww = h * 360 / 836; s.addImage({ path: img(p.photo), x: x + (w - ww) / 2, y: y + 0.12, w: ww, h }); }
    else if (p.photo) { circle(s, x + (w - d) / 2 - 0.04, y + 0.13, d + 0.08, GREEN); s.addImage({ path: img(p.photo), x: x + (w - d) / 2, y: y + 0.17, w: d, h: d, rounding: true }); }
    else { circle(s, x + (w - d) / 2, y + 0.17, d, GREEN); T(s, p.initials, { x: x + (w - d) / 2, y: y + 0.17, w: d, h: d, fontFace: "Cambria", fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" }); }
    T(s, p.name, { x: x + 0.08, y: y + 1.35, w: w - 0.16, h: 0.3, fontFace: "Cambria", fontSize: 11.5, bold: true, align: "center", valign: "middle" });
    T(s, p.role, { x: x + 0.08, y: y + 1.65, w: w - 0.16, h: 0.34, fontSize: 8.5, color: MUTED, align: "center" });
    T(s, [{ text: "Ziel: ", options: { bold: true, color: GREEN } }, { text: p.ziel }], { x: x + 0.14, y: y + 2.05, w: w - 0.28, h: 0.6, fontSize: 8.5 });
    T(s, [{ text: "Problem heute: ", options: { bold: true, color: GREEN } }, { text: p.pain }], { x: x + 0.14, y: y + 2.7, w: w - 0.28, h: 0.6, fontSize: 8.5 });
  });
  rbox(s, 0.45, 4.5, 9.1, 0.62, AMBERPALE, AMBER);
  T(s, [{ text: "Weitere Stakeholder ohne eigene Persona: ", options: { bold: true, color: BROWN } }, { text: "Zweitprüfer:in (intern oder extern) und Prüfungsausschussvorsitzende:r. Sie handeln in einzelnen Schritten und sind im Kontextdiagramm und in der Story Map berücksichtigt." }], { x: 0.6, y: 4.5, w: 8.8, h: 0.62, fontSize: 9, valign: "middle" });
}

// =====================================================================
// 5 Story Map mit Releases
// =====================================================================
const REL = { 1: { fill: TINT, line: GREEN, txt: GREEN, name: "Release 1" }, 2: { fill: BLUEPALE, line: BLUE, txt: "1F3A5C", name: "Release 2" }, 3: { fill: AMBERPALE, line: AMBER, txt: BROWN, name: "Release 3" } };
{
  const s = slide("Story Map mit Releaseplanung", "4  ·  STORY MAP", "7 Activities, 29 User Tasks, 66 User Stories. Die vollständige Karte liegt als Story_Mapping_V2 vor.");
  const A = [
    ["Vorbereitung der Anmeldung", [["Thema abstimmen", 1], ["Zweitprüfung vorschlagen", 1], ["Externes Korreferat klären", 1], ["Starttermin abstimmen", 1]]],
    ["Antragstellung", [["Antragsdaten erfassen", 1], ["Unterlagen ergänzen", 1], ["Digital bestätigen", 1], ["Antrag absenden", 1, true]]],
    ["Prüfung des Antrags", [["Formalien prüfen", 1], ["Über Antrag entscheiden", 1]]],
    ["Ausgabe der Arbeit", [["Thema bestätigen und freigeben", 1, true], ["Bearbeitungszeit festlegen", 1], ["Thema ausgeben", 1]]],
    ["Durchführung der Arbeit", [["Zeitraum und Status einsehen", 2], ["Fristen verfolgen", 2], ["Verlängerung beantragen", 2, true], ["Stellungnahme abgeben", 2], ["Verlängerung entscheiden", 2]]],
    ["Abgabe der Arbeit", [["Arbeit abgeben", 3], ["Abgabe erfassen", 3], ["Eingang bestätigen", 3], ["Beteiligte informieren", 3]]],
    ["Abschluss der Dokumentation", [["Antragsstatus einsehen", 1], ["Antragsstatus einsehen (Dekanat)", 1], ["Rückmeldung erhalten", 1], ["Fehlende Unterlagen einreichen", 1], ["Dokumentation prüfen", 3], ["Abschlussbestätigung erhalten", 3], ["Dokumentation abrufen", 3]]],
  ];
  const cw = 1.24, gap = 0.07, y0 = 1.3;
  seg(s, 0.45, 1.13, 9.55, 1.13, true, GREEN, 2);
  A.forEach((a, i) => {
    const x = 0.45 + i * (cw + gap);
    circle(s, x + cw / 2 - 0.13, 1.0, 0.26, GREEN, WHITE, 1.5);
    T(s, String(i + 1), { x: x + cw / 2 - 0.13, y: 1.0, w: 0.26, h: 0.26, fontSize: 9, bold: true, color: WHITE, align: "center", valign: "middle" });
    rbox(s, x, y0, cw, 0.5, GREEN);
    T(s, a[0], { x: x + 0.03, y: y0, w: cw - 0.06, h: 0.5, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
    a[1].forEach((t, k) => {
      const r = REL[t[1]], yy = y0 + 0.58 + k * 0.4;
      rbox(s, x, yy, cw, 0.35, r.fill, r.line, { line: { color: r.line, width: t[2] ? 2 : 0.75 }, rectRadius: 0.05 });
      T(s, t[0], { x: x + 0.03, y: yy, w: cw - 0.06, h: 0.35, fontSize: 7, bold: !!t[2], color: r.txt, align: "center", valign: "middle" });
    });
  });
  [1, 2, 3].forEach((r, i) => { const x = 0.45 + i * 1.3; rbox(s, x, 4.85, 0.3, 0.2, REL[r].fill, REL[r].line, { rectRadius: 0.04 }); T(s, REL[r].name, { x: x + 0.36, y: 4.8, w: 0.9, h: 0.3, fontSize: 8, color: MUTED, valign: "middle" }); });
  T(s, "Dick umrandet: die drei User Tasks, aus denen die im Detail vorgestellten User Stories stammen.", { x: 4.5, y: 4.8, w: 5.05, h: 0.3, fontSize: 8, italic: true, color: MUTED, align: "right", valign: "middle" });
}

// =====================================================================
// 6 Releaseplanung
// =====================================================================
{
  const s = slide("Releaseplanung", "4  ·  RELEASES", "Drei Releases, priorisiert nach Auftrag, Nutzen und Abhängigkeiten");
  const R = [
    [1, "Anmeldung bis Ausgabe", ["Activities 1 bis 4: vorbereiten, Antrag stellen, prüfen, Thema freigeben und ausgeben", "Einbindung externer Zweitprüfer:innen", "Antragsstatus und Rückmeldung für Studierende"], "Kern des Auftrags. Ersetzt den PDF-Antrag vollständig und liefert die Daten, auf denen alles Weitere aufbaut."],
    [2, "Durchführung und Verlängerung", ["Bearbeitungszeitraum, Abgabetermin und Fristen einsehen", "Verlängerungsantrag stellen", "Stellungnahme und Entscheidung zur Verlängerung"], "Zweite Ausbaustufe laut Auftrag. Braucht Beginn und Abgabetermin aus Release 1."],
    [3, "Abgabe und Abschluss", ["Arbeit als PDF abgeben, Eingang bestätigen", "Beteiligte über die Abgabe informieren", "Dokumentation prüfen und abrufen"], "Erweiterung über den Auftrag hinaus. Umfang mit dem Auftraggeber abstimmen."],
  ];
  R.forEach((r, i) => {
    const x = 0.45 + i * 3.075, w = 2.95, rr = REL[r[0]];
    rbox(s, x, 1.02, w, 4.1, rr.fill, rr.line);
    T(s, rr.name.toUpperCase(), { x: x + 0.18, y: 1.1, w: w - 0.36, h: 0.22, fontSize: 8, bold: true, color: rr.txt, charSpacing: 1.5, valign: "middle" });
    T(s, r[1], { x: x + 0.18, y: 1.32, w: w - 0.36, h: 0.4, fontFace: "Cambria", fontSize: 14, bold: true, color: rr.txt, valign: "middle" });
    bullets(s, r[2], { x: x + 0.2, y: 1.85, w: w - 0.38, h: 1.75, fontSize: 9.5, paraSpaceAfter: 5 });
    rbox(s, x + 0.14, 3.72, w - 0.28, 1.25, WHITE, rr.line, { rectRadius: 0.05 });
    T(s, [{ text: "Begründung: ", options: { bold: true, color: rr.txt } }, { text: r[3] }], { x: x + 0.24, y: 3.72, w: w - 0.48, h: 1.25, fontSize: 9, valign: "middle" });
  });
}

// =====================================================================
// 7 Gesamt-Datenmodell
// =====================================================================
const LH = 0.172;
const klass = (s, x, y, w, name, attrs) => {
  const hn = 0.3, h = hn + attrs.length * LH + 0.1;
  rect(s, x, y, w, h, WHITE, INK, 1.25); rect(s, x, y, w, hn, TINT, INK, 1.25);
  T(s, name, { x, y, w, h: hn, fontSize: 9.5, bold: true, align: "center", valign: "middle" });
  const runs = [];
  attrs.forEach((a, i) => {
    const last = i === attrs.length - 1;
    runs.push({ text: (a[0] || "    ") + " ", options: { bold: true, color: GREEN } });
    runs.push({ text: a[1], options: { breakLine: !last } });
  });
  T(s, runs, { x: x + 0.07, y: y + hn + 0.04, w: w - 0.12, h: h - hn - 0.06, fontSize: 7.5, valign: "top" });
  return { x, y, w, h, b: y + h };
};
{
  const s = slide("Gesamt-Datenmodell", "5  ·  DATEN", "Vier Klassen. Entscheidungen und Stellungnahmen sind Attribute, keine eigenen Klassen.");
  const st = klass(s, 0.45, 1.05, 2.25, "Studierende:r", [["PK", "matrikelnummer : String *"], ["", "name : String *"], ["", "thm_email : String *"], ["", "studiengang : String"]]);
  const ar = klass(s, 3.45, 1.05, 2.5, "Abschlussarbeit", [["PK", "arbeit_id : Integer"], ["FK", "matrikelnummer : String"], ["FK", "erstpruefer_id : Integer"], ["FK", "zweitpruefer_id : Integer"], ["", "abschluss : Abschlussart"], ["", "thema : String"], ["", "firma : String"], ["", "wunsch_starttermin : Date"], ["", "status : ArbeitStatus"], ["", "thema_freigabe_datum : Date"], ["", "aenderungswunsch : Text"], ["", "beginn : Date"], ["", "abgabetermin : Date"]]);
  const pr = klass(s, 6.95, 1.05, 2.6, "Prüfer:in", [["PK", "pruefer_id : Integer"], ["", "name : String *"], ["", "extern : Boolean"], ["", "qualifikation_geprueft : Boolean"]]);
  const ve = klass(s, 6.95, 2.72, 2.6, "Verlängerungsantrag", [["PK", "antrag_id : Integer"], ["FK", "arbeit_id : Integer"], ["", "grund : Text"], ["", "nachweis : Datei (optional)"], ["", "datum : Date"], ["", "stellungnahme : Text"], ["", "entscheidung : Entscheidungsergebnis"], ["", "entscheidungsdatum : Date"], ["", "pa_vorsitzende_r : String"], ["", "neuer_abgabetermin : Date"], ["", "status : AntragStatus"]]);
  const m = (x, y, t, al) => T(s, t, { x, y, w: 0.42, h: 0.17, fontSize: 8, bold: true, align: al || "left", valign: "middle" });
  const l = (x, y, t, w) => T(s, t, { x, y, w: w || 0.75, h: 0.17, fontSize: 7.5, italic: true, color: MUTED, align: "center", valign: "middle" });
  seg(s, st.x + st.w, 1.55, ar.x, 1.55); l(st.x + st.w, 1.35, "schreibt ►"); m(st.x + st.w + 0.04, 1.58, "1"); m(ar.x - 0.46, 1.58, "1..n", "right");
  seg(s, ar.x + ar.w, 1.55, pr.x, 1.55); l(ar.x + ar.w, 1.35, "betreut von ►", 1.0); m(ar.x + ar.w + 0.04, 1.58, "0..n"); m(pr.x - 0.46, 1.58, "1..2", "right");
  seg(s, ar.x + ar.w, 3.2, ve.x, 3.2); l(ar.x + ar.w, 3.0, "hat ►", 1.0); m(ar.x + ar.w + 0.04, 3.23, "1"); m(ve.x - 0.46, 3.23, "0..n", "right");
  // Aufzaehlungstypen
  const en = [["Abschlussart", "Bachelor, Master"], ["ArbeitStatus", "eingereicht, rueckfrage, thema_freigegeben, zugelassen, ausgegeben"], ["AntragStatus", "eingereicht, stellungnahme_erteilt, entschieden"], ["Entscheidungsergebnis", "genehmigt, abgelehnt"]];
  T(s, "«enumeration»", { x: 0.45, y: 2.2, w: 2.25, h: 0.18, fontSize: 7.5, italic: true, color: MUTED });
  T(s, en.map((e, i) => [{ text: e[0] + ": ", options: { bold: true, color: GREEN } }, { text: e[1], options: { breakLine: i < en.length - 1 } }]).flat(), { x: 0.45, y: 2.4, w: 2.75, h: 1.45, fontSize: 7.5, paraSpaceAfter: 3 });
  rbox(s, 0.45, 3.98, 6.2, 1.12, AMBERPALE, AMBER);
  T(s, [
    { text: "* aus dem Nachbarsystem THM-Login. ", options: { bold: true, color: BROWN } }, { text: "Alle übrigen Daten gehören ThesisFlow.", options: { breakLine: true } },
    { text: "Prüfer:in: ", options: { bold: true, color: BROWN } }, { text: "höchstens zwei je Abschlussarbeit, Erst- und Zweitprüfer:in.", options: { breakLine: true } },
    { text: "Formeln: ", options: { bold: true, color: BROWN } }, { text: "abgabetermin = beginn + Bearbeitungszeit laut Prüfungsordnung.  neuer_abgabetermin = abgabetermin + Verlängerung." },
  ], { x: 0.57, y: 3.98, w: 5.96, h: 1.12, fontSize: 8.5, valign: "middle", paraSpaceAfter: 3 });
}

// =====================================================================
// 8 Dialoglandkarte
// =====================================================================
{
  const s = slide("Dialoglandkarte", "6  ·  DIALOGE", "Zugriff nach THM-Login. Die Startseite richtet sich nach der Rolle.");
  const dlg = (x, y, w, t, kind) => { const f = kind === "story" ? AMBERPALE : kind === "top" ? GREEN : TINT, ln = kind === "story" ? AMBER : kind === "top" ? GREEN : TINT; rbox(s, x, y, w, 0.36, f, ln, { rectRadius: 0.05, line: { color: ln, width: kind === "story" ? 1.5 : 1 } }); T(s, t, { x: x + 0.04, y, w: w - 0.08, h: 0.36, fontSize: 8.5, bold: kind !== undefined, color: kind === "top" ? WHITE : INK, align: "center", valign: "middle" }); };
  dlg(4.0, 1.02, 2.0, "THM-Login", "top");
  seg(s, 5.0, 1.38, 5.0, 1.55, true, MUTED);
  dlg(4.0, 1.55, 2.0, "Startseite je Rolle", "top");
  seg(s, 5.0, 1.91, 5.0, 2.08, false, MUTED);
  const cols = [
    ["Studierende:r", [["Meine Abschlussarbeit"], ["Antrag stellen", "story"], ["Status und Fristen"], ["Verlängerung beantragen", "story"]]],
    ["Erstprüfer:in", [["Meine Themen"], ["Thema prüfen und freigeben", "story"], ["Bearbeitungszeit festlegen"], ["Stellungnahme zur Verlängerung"]]],
    ["Dekanat", [["Antragsliste"], ["Formalien prüfen"], ["Frist prüfen, Ausgabe bestätigen"]]],
    ["Prüfungsausschussvorsitzende:r", [["Entscheidungsliste"], ["Über Antrag entscheiden"], ["Über Verlängerung entscheiden"]]],
  ];
  const cw = 2.15, gap = 0.166;
  seg(s, 0.45 + cw / 2, 2.08, 0.45 + 3 * (cw + gap) + cw / 2, 2.08, false, MUTED);
  cols.forEach((col, i) => {
    const x = 0.45 + i * (cw + gap), xc = x + cw / 2;
    seg(s, xc, 2.08, xc, 2.25, true, MUTED);
    rbox(s, x, 2.25, cw, 0.36, WHITE, GREEN, { rectRadius: 0.05, line: { color: GREEN, width: 1.25 } });
    T(s, col[0], { x, y: 2.25, w: cw, h: 0.36, fontSize: 9, bold: true, color: GREEN, align: "center", valign: "middle" });
    col[1].forEach((d, k) => { const y = 2.78 + k * 0.5; seg(s, xc, y - 0.14, xc, y, true, MUTED); dlg(x + 0.1, y, cw - 0.2, d[0], d[1]); });
  });
  rbox(s, 0.45, 4.85, 0.3, 0.2, AMBERPALE, AMBER, { rectRadius: 0.04 });
  T(s, "Dialoge der drei im Detail vorgestellten User Stories", { x: 0.82, y: 4.8, w: 5, h: 0.3, fontSize: 8, color: MUTED, valign: "middle" });
}

// =====================================================================
// Drei User Stories: je drei Folien
// =====================================================================
const storyHead = (n, title, part) => slide(`Story ${n}: ${title}`, `7  ·  STORY ${n} VON 3`, part);
const storyA = (n, o) => {
  const s = storyHead(n, o.title, "User und Interface");
  rbox(s, 0.45, 1.02, 3.7, 1.2, TINT);
  T(s, "„" + o.story + "“", { x: 0.6, y: 1.02, w: 3.4, h: 1.2, fontFace: "Cambria", fontSize: 10.5, italic: true, color: GREEN, valign: "middle" });
  T(s, o.ref, { x: 0.45, y: 2.27, w: 3.7, h: 0.2, fontSize: 7.5, italic: true, color: GREY });
  T(s, "User", { x: 0.45, y: 2.6, w: 3.7, h: 0.26, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  const d = 0.8, py = 2.95;
  if (o.photo && o.full) { const h = 0.95, w = h * 360 / 836; s.addImage({ path: img(o.photo), x: 0.65, y: py - 0.05, w, h }); }
  else if (o.photo) s.addImage({ path: img(o.photo), x: 0.5, y: py, w: d, h: d, rounding: true });
  else { circle(s, 0.5, py, d, GREEN); T(s, o.initials, { x: 0.5, y: py, w: d, h: d, fontFace: "Cambria", fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" }); }
  T(s, [{ text: o.persona, options: { bold: true, fontSize: 10, breakLine: true } }, { text: o.role, options: { fontSize: 8.5, color: MUTED } }], { x: 1.45, y: py, w: 2.7, h: d, valign: "middle" });
  T(s, [
    { text: "Ziel: ", options: { bold: true, color: GREEN } }, { text: o.ziel, options: { breakLine: true } },
    { text: "Problem heute: ", options: { bold: true, color: GREEN } }, { text: o.pain, options: { breakLine: true } },
    { text: "Weitere Beteiligte: ", options: { bold: true, color: GREEN } }, { text: o.others },
  ], { x: 0.45, y: 3.95, w: 3.7, h: 1.2, fontSize: 8.5, paraSpaceAfter: 3 });
  T(s, "Interface", { x: 4.4, y: 1.0, w: 5.15, h: 0.26, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  if (o.image && fs.existsSync(img(o.image))) {
    const w = 5.15, h = w * o.ratio;
    s.addImage({ path: img(o.image), x: 4.4, y: 1.35, w, h });
    T(s, o.caption, { x: 4.4, y: 1.4 + h, w, h: 0.4, fontSize: 8, italic: true, color: MUTED });
  } else {
    rbox(s, 4.4, 1.32, 5.15, 3.8, BG, GREY, { line: { color: GREY, width: 1.25, dashType: "dash" }, rectRadius: 0.1 });
    T(s, o.placeholder, { x: 4.6, y: 2.75, w: 4.75, h: 0.9, fontSize: 11, color: MUTED, align: "center", valign: "middle" });
  }
  return s;
};
const storyC = (n, o) => {
  const s = storyHead(n, o.title, "Data, Environment und Quality Attribute");
  T(s, "Data: Ausschnitt aus dem Gesamt-Datenmodell", { x: 0.45, y: 1.0, w: 4.5, h: 0.26, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  table(s, [[hdr("Klasse"), hdr("Attribute"), hdr("Zugriff")]].concat(o.data.map(r => [b(r[0]), c(r[1]), c(r[2])])), { x: 0.45, y: 1.32, w: 4.5, colW: [1.25, 2.35, 0.9], rowH: 0.42 });
  if (o.dataNote) T(s, o.dataNote, { x: 0.45, y: 1.4 + 0.42 * (o.data.length + 1), w: 4.5, h: 0.5, fontSize: 8.5, italic: true, color: MUTED });
  T(s, "Environment", { x: 5.2, y: 1.0, w: 4.35, h: 0.26, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  bullets(s, o.env, { x: 5.22, y: 1.32, w: 4.3, h: 0.95, fontSize: 9, paraSpaceAfter: 2 });
  T(s, "Quality Attribute", { x: 5.2, y: 2.4, w: 4.35, h: 0.26, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  table(s, [[hdr("Attribut"), hdr("Anforderung"), hdr("Prüfbarkeit")]].concat(o.quality.map(r => [b(r[0]), c(r[1], { fontSize: 8 }), c(r[2], { fontSize: 8 })])), { x: 5.2, y: 2.72, w: 4.35, colW: [1.0, 1.75, 1.6], rowH: 0.6 });
  return s;
};

// ---------- Story 1: Antrag einreichen ----------
const S1 = {
  title: "Antrag einreichen",
  story: "Als Studierende:r möchte ich meinen Antrag verbindlich über das Portal einreichen, damit die Anmeldung meiner Abschlussarbeit offiziell gestartet wird.",
  ref: "US 2.4.2  ·  Activity 2 Antragstellung  ·  User Task 2.4 Antrag absenden  ·  Release 1",
  photo: "bachelor.png", full: true, persona: "Moritz Hoffmann", role: "Studierende:r",
  ziel: "Antrag in einem Durchgang korrekt einreichen.", pain: "Unsicher, ob der Antrag vollständig und angekommen ist.", others: "Erstprüfer:in, Dekanat.",
  placeholder: "Interface-Entwurf „Antrag stellen“\nwird durch das Projektteam ergänzt.",
  data: [["Studierende:r", "matrikelnummer, name, studiengang", "lesen"], ["Prüfer:in", "name, extern", "lesen"], ["Abschlussarbeit", "thema, firma, wunsch_starttermin, Erst- und Zweitprüfer:in", "anlegen"], ["Abschlussarbeit", "status = eingereicht", "ändern"]],
  dataNote: "Name und Matrikelnummer kommen aus dem THM-Login.",
  env: ["Webportal im Browser, am Desktop und am Smartphone", "Zugriff nach THM-Login", "Nutzung zu Hause und unterwegs"],
  quality: [["Benutzbarkeit", "Schritt-für-Schritt-Führung mit Prüfung der Pflichtangaben.", "Eine Testperson reicht den Antrag ohne Hilfe ein."], ["Zuverlässigkeit", "Beim Absenden gehen keine Angaben verloren.", "Gespeicherte Angaben stimmen mit der Eingabe überein."], ["Transparenz", "Der Eingang wird sofort bestätigt.", "Bestätigung erscheint nach dem Absenden."]],
};
storyA(1, S1);
{
  const s = storyHead(1, S1.title, "Action und Control");
  const cy = pool(s, 1.0, [{ name: "Studierende:r", h: 1.15 }, { name: "Portal", h: 1.35 }], "Antrag einreichen");
  const x1 = 1.75, x2 = 3.25, xg = 4.95, x3 = 5.45, x4 = 7.0, xe = 8.6;
  startEv(s, 1.2, cy[0], "Angaben erfasst");
  flow(s, [[1.46, cy[0]], [x1, cy[0]]]);
  task(s, x1, cy[0], "Angaben und Unterlagen prüfen");
  flow(s, [[x1 + TW, cy[0]], [x2, cy[0]]]);
  task(s, x2, cy[0], "Antrag absenden");
  flow(s, [[x2 + TW / 2, cy[0] + TH / 2], [x2 + TW / 2, cy[1] - TH / 2]]);
  task(s, x2, cy[1], "Antragseingang verarbeiten", true);
  flow(s, [[x2 + TW, cy[1]], [xg - 0.2, cy[1]]]);
  gw(s, xg, cy[1], "Pflichtangaben vollständig?");
  flow(s, [[xg + 0.2, cy[1]], [x3, cy[1]]]); T(s, "ja", { x: xg + 0.2, y: cy[1] - 0.19, w: 0.2, h: 0.16, fontSize: 7, color: MUTED });
  task(s, x3, cy[1], "Status „eingereicht“ setzen", true);
  flow(s, [[x3 + TW, cy[1]], [x4, cy[1]]]);
  task(s, x4, cy[1], "Eingang bestätigen, Erstprüfer:in informieren", true);
  flow(s, [[x4 + TW, cy[1]], [xe, cy[1]]]);
  endEv(s, xe, cy[1], "Antrag eingereicht", true);
  flow(s, [[xg, cy[1] - 0.2], [xg, cy[0]], [x3, cy[0]]]); T(s, "nein", { x: xg + 0.04, y: cy[0] - 0.19, w: 0.35, h: 0.16, fontSize: 7, color: MUTED });
  task(s, x3, cy[0], "Fehlende Angaben ergänzen");
  flow(s, [[x3 + TW / 2, cy[0] - TH / 2], [x3 + TW / 2, cy[0] - 0.45], [x2 + TW / 2, cy[0] - 0.45], [x2 + TW / 2, cy[0] - TH / 2]]);
  rules(s, 3.65, [
    "Das System muss den Zugriff erst nach THM-Login gewähren.",
    "Falls Pflichtangaben fehlen, muss das System das Absenden verhindern und die fehlenden Angaben anzeigen.",
    "Sobald der Antrag abgesendet ist, muss das System den Status „eingereicht“ setzen.",
    "Sobald der Status gesetzt ist, muss das System der Studierenden den Eingang bestätigen.",
    "Das System muss die Erstprüfer:in über den eingereichten Antrag informieren.",
  ]);
}
storyC(1, S1);

// ---------- Story 2: Thema bestaetigen und freigeben ----------
const S2 = {
  title: "Thema bestätigen und freigeben",
  story: "Als Erstprüfer:in möchte ich das Thema der Abschlussarbeit bestätigen, damit die Arbeit mit dem abgestimmten Thema offiziell freigegeben wird.",
  ref: "US 4.1.1  ·  Activity 4 Ausgabe der Arbeit  ·  User Task 4.1  ·  Release 1",
  initials: "AB", persona: "Prof. Dr. Andreas Becker", role: "Erstprüfer:in",
  ziel: "Thema prüfen und freigeben, ohne Formular und Unterschrift.", pain: "Unteren Formularteil ausfüllen, unterschreiben, per E-Mail senden.", others: "Studierende:r, Dekanat.",
  image: "interface_themenfreigabe.png", ratio: 0.5625, caption: "Prototyp des Projektteams: „Meine Themen“ und „Thema prüfen und freigeben“.",
  placeholder: "Screenshot aus dem Prototyp\n„Thema prüfen und freigeben“\nwird durch das Projektteam eingefügt.",
  data: [["Abschlussarbeit", "thema, wunsch_starttermin, status", "lesen"], ["Studierende:r", "name, studiengang", "lesen"], ["Abschlussarbeit", "thema_freigabe_datum, status = thema_freigegeben", "ändern"], ["Abschlussarbeit", "aenderungswunsch, status = rueckfrage", "ändern"]],
  dataNote: "Die Themenfreigabe ist ein Attribut der Abschlussarbeit, keine eigene Klasse.",
  env: ["Webportal im Browser, am Desktop und am Smartphone", "Zugriff nach THM-Login", "Nutzung im Büro und zwischen Lehrveranstaltungen"],
  quality: [["Sicherheit", "Nur die zugeordnete Erstprüfer:in kann das Thema freigeben.", "Versuch mit anderem Konto wird verhindert."], ["Zuverlässigkeit", "Die Freigabe bleibt unverändert gespeichert.", "Datum und Thema sind nach erneutem Aufruf identisch."], ["Benutzbarkeit", "Freigabe ohne Schulung auf beiden Gerätetypen.", "Nutzungstest mit Erstprüfenden."]],
};
storyA(2, S2);
{
  const s = storyHead(2, S2.title, "Action und Control");
  const cy = pool(s, 1.0, [{ name: "Erstprüfer:in", h: 1.5 }, { name: "Portal", h: 1.3 }], "Themenfreigabe");
  const yA = cy[0] - 0.45, yE = cy[0] + 0.25, yP = cy[1] - 0.1;
  const x1 = 1.75, x2 = 3.2, xg = 4.8, x3 = 5.25, x4 = 7.2, xe1 = 6.65, xe2 = 8.62;
  startEv(s, 1.2, yP, "Antrag eingereicht");
  flow(s, [[1.46, yP], [x1, yP]]);
  task(s, x1, yP, "Erstprüfer:in informieren", true);
  flow(s, [[x1 + TW / 2, yP - TH / 2], [x1 + TW / 2, yE + TH / 2]]);
  task(s, x1, yE, "Thema aufrufen");
  flow(s, [[x1 + TW, yE], [x2, yE]]);
  task(s, x2, yE, "Thema prüfen");
  flow(s, [[x2 + TW, yE], [xg - 0.2, yE]]);
  gw(s, xg, yE, "Thema wie abgestimmt?");
  flow(s, [[xg + 0.2, yE], [x3, yE]]); T(s, "ja", { x: xg + 0.2, y: yE - 0.19, w: 0.2, h: 0.16, fontSize: 7, color: MUTED });
  task(s, x3, yE, "Thema bestätigen und freigeben");
  flow(s, [[x3 + TW / 2, yE + TH / 2], [x3 + TW / 2, yP - TH / 2]]);
  task(s, x3, yP, "Freigabe speichern, Status setzen", true);
  flow(s, [[x3 + TW, yP], [xe1, yP]]);
  endEv(s, xe1, yP, "Thema freigegeben", true);
  flow(s, [[xg, yE - 0.2], [xg, yA], [x4, yA]]); T(s, "nein", { x: xg + 0.04, y: yA - 0.19, w: 0.35, h: 0.16, fontSize: 7, color: MUTED });
  task(s, x4, yA, "Änderungswunsch eintragen");
  flow(s, [[x4 + TW / 2, yA + TH / 2], [x4 + TW / 2, yP - TH / 2]]);
  task(s, x4, yP, "Speichern, Studierende:n informieren", true);
  flow(s, [[x4 + TW, yP], [xe2, yP]]);
  endEv(s, xe2, yP, "Rückfrage gesendet", true);
  rules(s, 3.9, [
    "Das System muss den Zugriff erst nach THM-Login gewähren.",
    "Das System muss ausschließlich der zugeordneten Erstprüfer:in die Möglichkeit bieten, das Thema freizugeben.",
    "Sobald das Thema freigegeben ist, muss das System das Freigabedatum speichern und den Status „thema_freigegeben“ setzen.",
    "Falls das Thema nicht bestätigt wird, muss das System den Änderungswunsch speichern und die Studierende:n informieren.",
  ]);
}
storyC(2, S2);

// ---------- Story 3: Verlaengerungsantrag stellen ----------
const S3 = {
  title: "Verlängerungsantrag stellen",
  story: "Als Studierende:r möchte ich einen Verlängerungsantrag direkt im Portal stellen, damit kein separater Papierantrag notwendig ist.",
  ref: "US 5.3.1 bis 5.3.3  ·  Activity 5 Durchführung der Arbeit  ·  User Task 5.3  ·  Release 2",
  photo: "bachelor.png", full: true, persona: "Moritz Hoffmann", role: "Studierende:r",
  ziel: "Verlängerung bei Bedarf online beantragen.", pain: "Papierantrag und unklare Fristen.", others: "Erstprüfer:in, Prüfungsausschussvorsitzende:r.",
  image: "wireframes_verlaengerung.png", ratio: 985 / 2070, caption: "Wireframes des Projektteams. Ergänzung nach Rückmeldung: optionales Uploadfeld für einen Nachweis, Anzeige von altem und neuem Abgabetermin.",
  data: [["Abschlussarbeit", "thema, firma, abgabetermin", "lesen"], ["Verlängerungsantrag", "grund, nachweis (optional), datum, status = eingereicht", "anlegen"], ["Verlängerungsantrag", "stellungnahme, entscheidung, entscheidungsdatum, pa_vorsitzende_r", "ändern"], ["Verlängerungsantrag", "neuer_abgabetermin", "ändern"]],
  dataNote: "Je Abschlussarbeit sind mehrere Verlängerungsanträge möglich (0..n). Der alte Abgabetermin bleibt erhalten.",
  env: ["Webportal im Browser, am Desktop und am Smartphone", "Zugriff nach THM-Login", "Nutzung unterwegs und zu Hause"],
  quality: [["Benutzbarkeit", "Antrag mit wenigen Angaben, Thema und Firma sind vorbelegt.", "Eine Testperson stellt den Antrag ohne Hilfe."], ["Transparenz", "Status, alter und neuer Abgabetermin sind sichtbar.", "Beide Termine erscheinen nach der Entscheidung."], ["Datenschutz", "Der Nachweis ist nur für Berechtigte einsehbar.", "Rollen- und Rechtetest."]],
};
storyA(3, S3);
{
  const s = storyHead(3, S3.title, "Action und Control");
  const cy = pool(s, 1.0, [{ name: "Studierende:r", h: 0.9 }, { name: "Portal", h: 0.9 }, { name: "Erstprüfer:in", h: 0.9 }], "Verlängerungsantrag");
  const x1 = 1.75, x2 = 3.3, x3 = 4.85, x4 = 6.4, xe = 8.2;
  startEv(s, 1.2, cy[0], "");
  flow(s, [[1.46, cy[0]], [x1, cy[0]]]);
  task(s, x1, cy[0], "Verlängerungsantrag starten");
  flow(s, [[x1 + TW, cy[0]], [x2, cy[0]]]);
  task(s, x2, cy[0], "Grund angeben, Nachweis optional");
  flow(s, [[x2 + TW, cy[0]], [x3, cy[0]]]);
  task(s, x3, cy[0], "Antrag absenden");
  flow(s, [[x3 + TW / 2, cy[0] + TH / 2], [x3 + TW / 2, cy[1] - TH / 2]]);
  task(s, x3, cy[1], "Antragseingang verarbeiten", true);
  flow(s, [[x3 + TW, cy[1]], [x4, cy[1]]]);
  task(s, x4, cy[1], "Status „eingereicht“, Erstprüfer:in informieren", true);
  flow(s, [[x4 + TW / 2, cy[1] + TH / 2], [x4 + TW / 2, cy[2] - TH / 2]]);
  task(s, x4, cy[2], "Antrag erhalten");
  flow(s, [[x4 + TW, cy[2]], [xe, cy[2]]]);
  endEv(s, xe, cy[2], "weiter mit Stellungnahme");
  rules(s, 3.85, [
    "Das System muss den Zugriff erst nach THM-Login gewähren.",
    "Falls der Grund fehlt, muss das System das Absenden verhindern.",
    "Das System muss der Prüfungsausschussvorsitzenden die Entscheidung erst ermöglichen, wenn die Stellungnahme der Erstprüfer:in vorliegt.",
    "Sobald genehmigt ist, muss das System berechnen: neuer_abgabetermin = abgabetermin + Verlängerung.",
  ]);
}
storyC(3, S3);

// =====================================================================
// Testfaelle
// =====================================================================
{
  const s = slide("Testfälle", "8  ·  TESTS", "Nach dem Schema Given, When, Then. Jeder Testfall prüft eine Control-Regel.");
  const rows = [
    ["T1", "Story 1, C3", "Alle Pflichtangaben sind erfasst.", "Die Studierende sendet den Antrag ab.", "Status ist „eingereicht“, der Eingang wird bestätigt."],
    ["T2", "Story 1, C2", "Der gewünschte Starttermin fehlt.", "Die Studierende sendet den Antrag ab.", "Der Antrag wird nicht gesendet, die fehlende Angabe wird angezeigt."],
    ["T3", "Story 2, C3", "Die Abschlussarbeit ist eingereicht, die zugeordnete Erstprüfer:in ist angemeldet.", "Sie gibt das Thema frei.", "Freigabedatum ist gespeichert, Status ist „thema_freigegeben“."],
    ["T4", "Story 2, C2", "Eine nicht zugeordnete Prüfer:in ist angemeldet.", "Sie ruft das Thema auf.", "Das System verweigert den Zugriff."],
    ["T5", "Story 3, C2", "Eine Abschlussarbeit läuft, das Feld Grund ist leer.", "Die Studierende sendet den Verlängerungsantrag ab.", "Der Antrag wird nicht gesendet."],
    ["T6", "Story 3, C4", "Abgabetermin 15.01., Verlängerung um 14 Tage ist genehmigt.", "Die Entscheidung wird gespeichert.", "Neuer Abgabetermin ist der 29.01., der alte bleibt sichtbar."],
  ];
  table(s, [[hdr("Nr."), hdr("Bezug"), hdr("Given"), hdr("When"), hdr("Then")]].concat(rows.map(r => [b(r[0]), c(r[1]), c(r[2]), c(r[3]), c(r[4])])), { x: 0.45, y: 1.05, w: 9.1, colW: [0.45, 1.0, 2.75, 2.2, 2.7], rowH: [0.3, 0.55, 0.55, 0.62, 0.55, 0.55, 0.62] });
}

// =====================================================================
// Offene Punkte und naechste Schritte
// =====================================================================
{
  const s = slide("Offene Punkte und nächste Schritte", "8  ·  AUSBLICK");
  rbox(s, 0.45, 1.0, 4.49, 4.1, AMBERPALE, AMBER);
  T(s, "Fragen an den Auftraggeber", { x: 0.6, y: 1.08, w: 4.2, h: 0.32, fontSize: 12, bold: true, color: BROWN, valign: "middle" });
  bullets(s, [
    "Wie lang darf eine Verlängerung laut Prüfungsordnung sein?",
    "Über welchen Weg sollen Beteiligte benachrichtigt werden?",
    "Liefert die Prüfungsverwaltung den Modulstand für die Zulassung über eine Schnittstelle?",
    "Gehört die Abgabe der Arbeit (Release 3) zum Auftrag?",
  ], { x: 0.62, y: 1.5, w: 4.15, h: 3.5, fontSize: 10.5, paraSpaceAfter: 7 });
  rbox(s, 5.065, 1.0, 4.49, 4.1, TINT);
  T(s, "Nächste Schritte", { x: 5.215, y: 1.08, w: 4.2, h: 0.32, fontSize: 12, bold: true, color: GREEN, valign: "middle" });
  bullets(s, [
    "Offene Fragen im Interview klären.",
    "User Stories von Release 1 nach den 7 Dimensionen ausarbeiten.",
    "Prototyp um die Dialoge der Studierenden erweitern.",
    "Testfälle für alle Regeln von Release 1 ergänzen.",
  ], { x: 5.235, y: 1.5, w: 4.15, h: 3.5, fontSize: 10.5, paraSpaceAfter: 7 });
}

const out = process.env.OUT || "ThesisFlow_Abschlusspraesentation.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
