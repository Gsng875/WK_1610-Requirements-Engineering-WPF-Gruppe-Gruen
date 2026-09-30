// Kontextdiagramm ThesisFlow (Datenfluss-Notation, Soll-Perspektive)
// Aufbau nach Vorlesung Kap. 2: Systemgrenze, Kontextgrenze, Grauzone;
// Kontextaspekte: Personen, Systeme im Betrieb, Prozesse, Ereignisse, Dokumente;
// Schnittstellen als Quellen (Input) und Senken (Output).
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "Kontextdiagramm ThesisFlow";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF", GREY = "8A968A", LGREY = "C9D1C9", PALE = "F5F7F4";
const IN = "2C5F2D", OUT = "4A7FB5"; // Farbe für Input-/Output-Flüsse
const s = pres.addSlide();
s.background = { color: WHITE };

// ---------- Kopf ----------
s.addText("Kontextdiagramm ThesisFlow", { x: 0.5, y: 0.22, w: 8, h: 0.5, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 24, bold: true, color: INK, valign: "middle" });
s.addText("Systemkontext in der Soll-Perspektive: Quellen und Senken des Abschlussarbeiten-Portals des FB MND", { x: 0.5, y: 0.7, w: 9, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED });
s.addText("REQUIREMENTS ENGINEERING  ·  GRUPPE GRÜN  ·  STAND 30.09.2026", { x: 8.8, y: 0.3, w: 4.05, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

// ---------- Kontextgrenze ----------
const KX = 0.5, KY = 1.05, KW = 10.15, KH = 5.65;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: KX, y: KY, w: KW, h: KH, fill: { color: PALE }, line: { color: MOSS, width: 1.25, dashType: "dash" }, rectRadius: 0.2 });
s.addText("Kontextgrenze  ·  Systemkontext", { x: KX + 0.2, y: KY + 0.08, w: 3, h: 0.22, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, italic: true, bold: true, color: MOSS });

// ---------- System ----------
const C = { x: 4.6, y: 3.05, w: 2.0, h: 1.3 };
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: C.x, y: C.y, w: C.w, h: C.h, fill: { color: GREEN }, line: { color: GREEN }, rectRadius: 0.12, shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.3 } });
s.addText([
  { text: "ThesisFlow", options: { fontFace: "Cambria", fontSize: 18, bold: true, color: WHITE, breakLine: true } },
  { text: "Portal für Anmeldung, Freigabe,", options: { fontFace: "Calibri", fontSize: 8.5, color: "D9E6CF", breakLine: true } },
  { text: "Ausgabe und Verlängerung", options: { fontFace: "Calibri", fontSize: 8.5, color: "D9E6CF", breakLine: true } },
  { text: "Systemgrenze", options: { fontFace: "Calibri", fontSize: 7.5, italic: true, color: MOSS } },
], { x: C.x, y: C.y, w: C.w, h: C.h, margin: 0.05, isTextBox: true, align: "center", valign: "middle" });

// ---------- Kontextelemente ----------
// kind: person | system | document | event
const EW = 1.9, EH = 0.55;
const styles = {
  person: { fill: TINT, line: GREEN, dash: "solid", txt: GREEN, tag: "Person / Rolle" },
  system: { fill: WHITE, line: GREY, dash: "dash", txt: MUTED, tag: "System im Betrieb" },
  document: { fill: WHITE, line: GREEN, dash: "sysDot", txt: GREEN, tag: "Dokument" },
  event: { fill: "FFF7E0", line: "C9A227", dash: "solid", txt: "7A5C00", tag: "Ereignis" },
};
const el = (x, y, label, kind) => {
  const st = styles[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: EW, h: EH, fill: { color: st.fill }, line: { color: st.line, width: 1.25, dashType: st.dash }, rectRadius: 0.08 });
  s.addText([
    { text: label, options: { fontSize: 9.5, bold: true, color: st.txt, breakLine: true } },
    { text: st.tag, options: { fontSize: 7, color: MUTED } },
  ], { x, y, w: EW, h: EH, margin: 0.04, isTextBox: true, fontFace: "Calibri", align: "center", valign: "middle" });
};

// Pfeil: eine Richtung, eigene Farbe
const arrow = (x1, y1, x2, y2, color) => {
  const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.abs(x2 - x1) || 0.001, h = Math.abs(y2 - y1) || 0.001;
  // Linie läuft von (x,y) nach (x+w,y+h); flipV, wenn Start unten links / Ende oben rechts
  const flipV = (x2 >= x1 && y2 < y1) || (x2 < x1 && y2 >= y1);
  // Richtung: Pfeilspitze am Ende (x2,y2). Bei gespiegelter Geometrie beginArrow statt endArrow.
  const endIsP2 = (x2 >= x1); // ohne Flip: Ende der Form ist rechts
  const opts = { color, width: 1.25 };
  if (endIsP2) opts.endArrowType = "triangle"; else opts.beginArrowType = "triangle";
  s.addShape(pres.shapes.LINE, { x, y, w, h, flipV, line: opts });
};

// Beschriftung mit zwei Richtungen
const label = (x, y, w, into, out) => {
  const runs = [];
  if (into) runs.push({ text: "▶ " + into, options: { color: IN, breakLine: !!out } });
  if (out) runs.push({ text: "◀ " + out, options: { color: OUT } });
  s.addText(runs, { x, y, w, h: 0.46, margin: 0.03, isTextBox: true, fontFace: "Calibri", fontSize: 7, fill: { color: WHITE }, line: { color: LGREY, width: 0.5 }, valign: "middle", lineSpacingMultiple: 1.05 });
};

// Linke Spalte: Personen (Quellen und Senken)
const left = [
  ["Studierende", "person", 1.35, "Antrag (Thema, Prüfende, Starttermin), Anlage ext. Prüfer, Verlängerungsantrag", "Status, Zulassungsbescheid, Ausgabe (Thema, Beginn, Abgabe)"],
  ["Erstprüfende", "person", 2.5, "Betreuungszusage, Ausgabe mit Beginn und Abgabe, Befürwortung Verlängerung", "Antragsdaten, Aufgabenhinweis"],
  ["Interne Zweitprüfende", "person", 3.65, "Zusage Korreferat", "Korreferatsanfrage, Antragsdaten"],
  ["Externe Korreferenten", "person", 4.8, "Zusage, Qualifikationsnachweis, Urkundenkopie", "Einladungslink (ohne THM-Konto), Antragsdaten"],
];
// Rechte Spalte: Organisation und Systeme
const right = [
  ["Sekretariat / Dekanat MND", "person", 1.35, "Ergebnis formale Prüfung, Fristprüfung", "Prüfliste, Antragsdaten, Fristenübersicht"],
  ["Prüfungsausschuss-Vorsitz", "person", 2.5, "Entscheidung Zulassung und ext. Prüfer, Begründung", "Entscheidungsvorlage"],
  ["Studienausschuss WI", "person", 3.65, "Entscheidung Verlängerung", "Verlängerungsantrag mit Befürwortung"],
  ["E-Mail-Dienst THM", "system", 4.8, "Zustellstatus", "Benachrichtigungen, Einladungen, Bescheide"],
];
const LX = 0.8, RX = 8.45;
const cyL = C.y + C.h / 2;
// Endpunkte an der Systemkante verteilen
const ends = [3.2, 3.55, 3.9, 4.25];
left.forEach((e, i) => {
  el(LX, e[2], e[0], e[1]);
  const yc = e[2] + EH / 2;
  arrow(LX + EW, yc - 0.06, C.x, ends[i] - 0.05, IN);   // Input ins System
  arrow(C.x, ends[i] + 0.05, LX + EW, yc + 0.06, OUT);  // Output aus dem System
  const my = (yc + ends[i]) / 2;
  label(2.85, my - 0.23, 1.6, e[3], e[4]);
});
right.forEach((e, i) => {
  el(RX, e[2], e[0], e[1]);
  const yc = e[2] + EH / 2;
  arrow(RX, yc - 0.06, C.x + C.w, ends[i] - 0.05, IN);
  arrow(C.x + C.w, ends[i] + 0.05, RX, yc + 0.06, OUT);
  const my = (yc + ends[i]) / 2;
  label(6.75, my - 0.23, 1.6, e[3], e[4]);
});

// Oben: THM-Login (System)
el(4.65, 1.3, "THM-Login (SSO)", "system");
arrow(5.5, 1.3 + EH, 5.5, C.y, IN);
arrow(5.7, C.y, 5.7, 1.3 + EH, OUT);
label(5.8, 2.2, 1.5, "Identität, Rolle, Matrikelnummer", "Authentifizierungsanfrage");

// Unten: Dokumente und Ereignis
const bottom = [
  [2.35, "Prüfungsordnungen / Allg. Best.", "document", "Zulassungsregeln, Fristen (Konfiguration)", null],
  [4.65, "Studierendenakte", "document", null, "abgeschlossener Vorgang als Akten-PDF"],
  [6.95, "Fristablauf / Rückgabefrist", "event", "Abgabetermin erreicht, 4-Wochen-Frist Themenrückgabe", null],
];
const bx = [4.95, 5.6, 6.25];
bottom.forEach((b, i) => {
  el(b[0], 5.85, b[1], b[2]);
  const xc = b[0] + EW / 2;
  if (b[3]) arrow(xc, 5.85, bx[i], C.y + C.h, IN);
  if (b[4]) arrow(bx[i], C.y + C.h, xc, 5.85, OUT);
});
label(3.0, 5.0, 1.55, bottom[0][3], null);
label(5.55, 5.15, 1.5, null, bottom[1][4]);
label(7.25, 5.0, 1.55, bottom[2][3], null);

// ---------- Grauzone ----------
const GX = 10.95, GY = 1.05, GW = 1.9, GH = 5.65;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX, y: GY, w: GW, h: GH, fill: { color: WHITE }, line: { color: LGREY, width: 1, dashType: "sysDot" }, rectRadius: 0.2 });
s.addText([
  { text: "Grauzone", options: { bold: true, fontSize: 9.5, color: MUTED, breakLine: true } },
  { text: "Beziehung zum System noch offen (Kontextabgrenzung)", options: { fontSize: 7.5, color: GREY } },
], { x: GX + 0.1, y: GY + 0.12, w: GW - 0.2, h: 0.6, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top" });
const grey = [
  ["Prüfungsverwaltung THM", "Modulstand für die Zulassungsprüfung: manuell oder Schnittstelle?"],
  ["Unternehmen der Studierenden", "Firma bei externer Thesis und Verlängerung: nur Datenfeld oder Beteiligter?"],
  ["Moodle / Lernplattform", "Kolloquium und Betreuung laufen dort, nicht im Portal?"],
  ["Postadresse / Papierakte", "Entfällt in der Soll-Perspektive, falls Akte digital geführt wird."],
];
grey.forEach((g, i) => {
  const y = GY + 0.85 + i * 1.18;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX + 0.12, y, w: GW - 0.24, h: 1.05, fill: { color: PALE }, line: { color: LGREY, width: 0.75, dashType: "dash" }, rectRadius: 0.06 });
  s.addText([
    { text: g[0], options: { bold: true, fontSize: 8.5, color: MUTED, breakLine: true } },
    { text: g[1], options: { fontSize: 7, color: GREY } },
  ], { x: GX + 0.18, y: y + 0.04, w: GW - 0.36, h: 0.97, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.05 });
});

// ---------- Legende ----------
const ly = 6.85;
const leg = (x, kind, text) => {
  const st = styles[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ly + 0.05, w: 0.32, h: 0.2, fill: { color: st.fill }, line: { color: st.line, width: 1, dashType: st.dash }, rectRadius: 0.04 });
  s.addText(text, { x: x + 0.38, y: ly, w: 1.5, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
};
leg(0.5, "person", "Person / Rolle");
leg(2.0, "system", "System im Betrieb");
leg(3.6, "document", "Dokument");
leg(4.8, "event", "Ereignis");
s.addShape(pres.shapes.LINE, { x: 6.0, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: IN, width: 1.25, endArrowType: "triangle" } });
s.addText("Input in das System (Quelle)", { x: 6.45, y: ly, w: 1.7, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.LINE, { x: 8.2, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: OUT, width: 1.25, beginArrowType: "triangle" } });
s.addText("Output aus dem System (Senke)", { x: 8.65, y: ly, w: 1.9, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addText("Quellen: Zulassungsantrag WS23/24, Anlage ext. Prüfer, PO B.Sc./M.Sc. WI 2024, Vorlesung Kap. 2", { x: 10.5, y: ly, w: 2.35, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 6.5, color: GREY, align: "right", valign: "middle" });

s.addNotes(
  "Kontextdiagramm ThesisFlow in Datenfluss-Notation (Soll-Perspektive, vgl. Vorlesung Kap. 2 System und Systemkontext abgrenzen).\n" +
  "Systemgrenze: das Portal ThesisFlow. Kontextgrenze: alle Personen, Systeme, Dokumente und Ereignisse, die Anforderungen beeinflussen.\n" +
  "Grüne Pfeile = Input (Quellen), blaue Pfeile = Output (Senken). Die Grauzone rechts enthält Aspekte, deren Beziehung zum System noch zu klären ist (Interviewfragen an den Auftraggeber).\n" +
  "Grundlage: 7-Schritte-Ablauf des Zulassungsantrags (Studierende, 2. Prüfer, Sekretariat, PA-Vorsitz, 1. Prüfer, Sekretariat, Studierende) und Verlängerungsantrag an den Studienausschuss WI."
);

const out = process.env.OUT || "Kontextdiagramm_ThesisFlow_v2.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
