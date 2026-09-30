// Kontextdiagramm ThesisFlow (Datenfluss-Notation, Soll-Perspektive)
// Aufbau nach Vorlesung Kap. 2: Systemgrenze, Kontextgrenze, Grauzone;
// Kontextaspekte: Personen, Systeme im Betrieb, Prozesse, Ereignisse, Dokumente;
// Schnittstellen als Quellen (Input) und Senken (Output).
// Layout: waagerechte Verbindungen links/rechts, Elbow-Verbindungen unten, keine Kreuzungen.
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "Kontextdiagramm ThesisFlow";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF", GREY = "8A968A", LGREY = "C9D1C9", PALE = "F5F7F4";
const IN = "2C5F2D", OUT = "4A7FB5";
const s = pres.addSlide();
s.background = { color: WHITE };

// ---------- Kopf ----------
s.addText("Kontextdiagramm ThesisFlow", { x: 0.5, y: 0.2, w: 8, h: 0.48, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 24, bold: true, color: INK, valign: "middle" });
s.addText("Systemkontext in der Soll-Perspektive: Quellen und Senken des Abschlussarbeiten-Portals des FB MND", { x: 0.5, y: 0.66, w: 9, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED });
s.addText("REQUIREMENTS ENGINEERING  ·  GRUPPE GRÜN  ·  STAND 30.09.2026", { x: 8.8, y: 0.28, w: 4.05, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

// ---------- Kontextgrenze ----------
const KX = 0.5, KY = 1.0, KW = 10.6, KH = 5.75;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: KX, y: KY, w: KW, h: KH, fill: { color: PALE }, line: { color: MOSS, width: 1.25, dashType: "dash" }, rectRadius: 0.2 });
s.addText("Kontextgrenze  ·  Systemkontext", { x: KX + 0.2, y: KY + 0.07, w: 3, h: 0.22, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, italic: true, bold: true, color: MOSS });

// ---------- Geometrie ----------
const EW = 1.9, EH = 0.55;
const LX = 0.8, LLX = 2.85, LLW = 1.9;          // linke Entitäten, linke Labels
const C = { x: 4.9, y: 1.3, w: 2.0, h: 3.75 };  // System
const RLX = 7.05, RLW = 1.9, RX = 9.1;          // rechte Labels, rechte Entitäten

// ---------- System ----------
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: C.x, y: C.y, w: C.w, h: C.h, fill: { color: GREEN }, line: { color: GREEN }, rectRadius: 0.12, shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.3 } });
s.addText([
  { text: "ThesisFlow", options: { fontFace: "Cambria", fontSize: 20, bold: true, color: WHITE, breakLine: true } },
  { text: "Portal für Anmeldung, Freigabe, Ausgabe und Verlängerung von Abschlussarbeiten", options: { fontFace: "Calibri", fontSize: 9, color: "D9E6CF", breakLine: true } },
  { text: " ", options: { fontSize: 6, breakLine: true } },
  { text: "Systemgrenze", options: { fontFace: "Calibri", fontSize: 8, italic: true, color: MOSS } },
], { x: C.x + 0.1, y: C.y, w: C.w - 0.2, h: C.h, margin: 0, isTextBox: true, align: "center", valign: "middle" });

// ---------- Elemente ----------
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
// waagerechte/senkrechte Linie mit Pfeilspitze am Ende (x2,y2)
const seg = (x1, y1, x2, y2, color, head) => {
  const horiz = Math.abs(y2 - y1) < 0.001;
  const x = Math.min(x1, x2), y = Math.min(y1, y2);
  const w = horiz ? Math.abs(x2 - x1) : 0.001, h = horiz ? 0.001 : Math.abs(y2 - y1);
  const opts = { color, width: 1.25 };
  if (head) { const endIsP2 = horiz ? x2 > x1 : y2 > y1; if (endIsP2) opts.endArrowType = "triangle"; else opts.beginArrowType = "triangle"; }
  s.addShape(pres.shapes.LINE, { x, y, w, h, line: opts });
};
const label = (x, y, w, into, out) => {
  const runs = [];
  if (into) runs.push({ text: "▶ " + into, options: { color: IN, breakLine: !!out } });
  if (out) runs.push({ text: "◀ " + out, options: { color: OUT } });
  s.addText(runs, { x, y, w, h: 0.5, margin: 0.04, isTextBox: true, fontFace: "Calibri", fontSize: 7, fill: { color: WHITE }, line: { color: LGREY, width: 0.5 }, valign: "middle", lineSpacingMultiple: 1.05 });
};

// ---------- linke Spalte (5 Slots) ----------
const left = [
  ["Studierende", "person", "Antrag (Thema, Prüfende, Starttermin), Anlage ext. Prüfer, Verlängerungsantrag", "Status, Zulassungsbescheid, Ausgabe (Thema, Beginn, Abgabe)"],
  ["Erstprüfende", "person", "Betreuungszusage, Ausgabe mit Beginn und Abgabe, Befürwortung Verlängerung", "Antragsdaten, Aufgabenhinweis"],
  ["Interne Zweitprüfende", "person", "Zusage Korreferat", "Korreferatsanfrage, Antragsdaten"],
  ["Externe Korreferenten", "person", "Zusage, Qualifikationsnachweis, Urkundenkopie", "Einladungslink (ohne THM-Konto), Antragsdaten"],
  ["Sekretariat / Dekanat MND", "person", "Ergebnis formale Prüfung, Fristprüfung", "Prüfliste, Antragsdaten, Fristenübersicht"],
];
const ysL = [1.45, 2.2, 2.95, 3.7, 4.45];
left.forEach((e, i) => {
  const y = ysL[i], yc = y + EH / 2;
  el(LX, y, e[0], e[1]);
  seg(LX + EW, yc - 0.07, C.x, yc - 0.07, IN, true);
  seg(C.x, yc + 0.07, LX + EW, yc + 0.07, OUT, true);
  label(LLX, yc - 0.25, LLW, e[2], e[3]);
});

// ---------- rechte Spalte (4 Slots) ----------
const right = [
  ["Prüfungsausschuss-Vorsitz", "person", "Entscheidung Zulassung und ext. Prüfer, Begründung", "Entscheidungsvorlage"],
  ["Studienausschuss WI", "person", "Entscheidung Verlängerung", "Verlängerungsantrag mit Befürwortung"],
  ["THM-Login (SSO)", "system", "Identität, Rolle, Matrikelnummer", "Authentifizierungsanfrage"],
  ["E-Mail-Dienst THM", "system", "Zustellstatus", "Benachrichtigungen, Einladungen, Bescheide"],
];
const ysR = [1.8, 2.65, 3.5, 4.35];
right.forEach((e, i) => {
  const y = ysR[i], yc = y + EH / 2;
  el(RX, y, e[0], e[1]);
  seg(RX, yc - 0.07, C.x + C.w, yc - 0.07, IN, true);
  seg(C.x + C.w, yc + 0.07, RX, yc + 0.07, OUT, true);
  label(RLX, yc - 0.25, RLW, e[2], e[3]);
});

// ---------- unten: Dokumente und Ereignis (Elbow-Verbindungen) ----------
const BY = 5.95, ELB = 5.35; // Entitäten-Oberkante, Höhe der waagerechten Elbow-Strecke
const bottom = [
  [2.6, "Prüfungsordnungen / Allg. Best.", "document", "Zulassungsregeln, Fristen (Konfiguration)", null, 5.35],
  [4.95, "Studierendenakte", "document", null, "abgeschlossener Vorgang als Akten-PDF", 5.9],
  [7.3, "Fristablauf / Rückgabefrist", "event", "Abgabetermin erreicht, 4-Wochen-Frist Themenrückgabe", null, 6.45],
];
bottom.forEach(b => {
  const [x, name, kind, into, out, sx] = b;
  el(x, BY, name, kind);
  const xc = x + EW / 2;
  const color = into ? IN : OUT;
  if (Math.abs(xc - sx) < 0.01) {
    // gerade Verbindung
    if (into) seg(xc, BY, xc, C.y + C.h, IN, true); else seg(xc, C.y + C.h, xc, BY, OUT, true);
  } else if (into) {
    seg(xc, BY, xc, ELB, IN, false);
    seg(xc, ELB, sx, ELB, IN, false);
    seg(sx, ELB, sx, C.y + C.h, IN, true);
  } else {
    seg(sx, C.y + C.h, sx, ELB, OUT, false);
    seg(sx, ELB, xc, ELB, OUT, false);
    seg(xc, ELB, xc, BY, OUT, true);
  }
  label(xc - 0.85, ELB + 0.07, 1.7, into, out);
});

// ---------- Grauzone ----------
const GX = 11.3, GY = 1.0, GW = 1.55, GH = 5.75;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX, y: GY, w: GW, h: GH, fill: { color: WHITE }, line: { color: LGREY, width: 1, dashType: "sysDot" }, rectRadius: 0.2 });
s.addText([
  { text: "Grauzone", options: { bold: true, fontSize: 9.5, color: MUTED, breakLine: true } },
  { text: "Beziehung zum System noch offen (Kontextabgrenzung)", options: { fontSize: 7, color: GREY } },
], { x: GX + 0.1, y: GY + 0.1, w: GW - 0.2, h: 0.6, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top" });
const grey = [
  ["Prüfungsverwaltung THM", "Modulstand für die Zulassungsprüfung: manuell oder Schnittstelle?"],
  ["Unternehmen der Studierenden", "Bei externer Thesis und Verlängerung: nur Datenfeld oder Beteiligter?"],
  ["Moodle / Lernplattform", "Kolloquium und Betreuung laufen dort, nicht im Portal?"],
  ["Papierakte / Postadresse", "Entfällt, falls die Akte digital geführt wird."],
];
grey.forEach((g, i) => {
  const y = GY + 0.8 + i * 1.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX + 0.1, y, w: GW - 0.2, h: 1.08, fill: { color: PALE }, line: { color: LGREY, width: 0.75, dashType: "dash" }, rectRadius: 0.06 });
  s.addText([
    { text: g[0], options: { bold: true, fontSize: 8, color: MUTED, breakLine: true } },
    { text: g[1], options: { fontSize: 7, color: GREY } },
  ], { x: GX + 0.16, y: y + 0.05, w: GW - 0.32, h: 1.0, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.05 });
});

// ---------- Legende ----------
const ly = 6.9;
const leg = (x, kind, text) => {
  const st = styles[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ly + 0.05, w: 0.32, h: 0.2, fill: { color: st.fill }, line: { color: st.line, width: 1, dashType: st.dash }, rectRadius: 0.04 });
  s.addText(text, { x: x + 0.38, y: ly, w: 1.4, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
};
leg(0.5, "person", "Person / Rolle");
leg(1.9, "system", "System im Betrieb");
leg(3.4, "document", "Dokument");
leg(4.5, "event", "Ereignis");
s.addShape(pres.shapes.LINE, { x: 5.6, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: IN, width: 1.25, endArrowType: "triangle" } });
s.addText("Input in das System (Quelle)", { x: 6.05, y: ly, w: 1.7, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.LINE, { x: 7.8, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: OUT, width: 1.25, beginArrowType: "triangle" } });
s.addText("Output aus dem System (Senke)", { x: 8.25, y: ly, w: 1.9, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addText("Quellen: Zulassungsantrag WS23/24, Anlage ext. Prüfer, PO B.Sc./M.Sc. WI 2024, Vorlesung Kap. 2", { x: 10.2, y: ly, w: 2.65, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 6.5, color: GREY, align: "right", valign: "middle" });

s.addNotes(
  "Kontextdiagramm ThesisFlow in Datenfluss-Notation (Soll-Perspektive, vgl. Vorlesung Kap. 2 System und Systemkontext abgrenzen).\n" +
  "Systemgrenze: das Portal ThesisFlow. Kontextgrenze: alle Personen, Systeme, Dokumente und Ereignisse, die Anforderungen beeinflussen.\n" +
  "Grüne Pfeile = Input (Quellen), blaue Pfeile = Output (Senken). Die Grauzone rechts enthält Aspekte, deren Beziehung zum System noch zu klären ist (Interviewfragen an den Auftraggeber).\n" +
  "Grundlage: 7-Schritte-Ablauf des Zulassungsantrags (Studierende, 2. Prüfer, Sekretariat, PA-Vorsitz, 1. Prüfer, Sekretariat, Studierende) und Verlängerungsantrag an den Studienausschuss WI."
);

// =====================================================================
// Folie 2: Ist-Zustand als Laufweg des Antragsformulars (kein System)
// Quelle: "Hinweise zur Bearbeitung" im Zulassungsantrag WS23/24 (7 Schritte),
// THM-Seite "Abschlussarbeit" FB MND, Verlaengerungsantrag.
// =====================================================================
(function buildIst() {
  const RED = "9B3B2E", REDPALE = "FBF1EF", ORANGE = "C9A227";
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("Kontextdiagramm Ist-Zustand", { x: 0.5, y: 0.2, w: 8, h: 0.48, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 24, bold: true, color: INK, valign: "middle" });
  s.addText("Heute gibt es kein System: das PDF-Formular wandert in sieben Übergaben per E-Mail, Papier und Scan zwischen den Beteiligten", { x: 0.5, y: 0.66, w: 10, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED });
  s.addText("REQUIREMENTS ENGINEERING  ·  GRUPPE GRÜN  ·  STAND 30.09.2026", { x: 8.8, y: 0.28, w: 4.05, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

  // Kontextgrenze
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.98, w: 10.4, h: 4.95, fill: { color: PALE }, line: { color: MOSS, width: 1.25, dashType: "dash" }, rectRadius: 0.2 });
  s.addText("Kontextgrenze  ·  Ist-Prozess (Laufweg des Antragsformulars)", { x: 0.7, y: 1.02, w: 5, h: 0.22, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, italic: true, bold: true, color: MOSS });

  const EW = 1.9, EH = 0.55;
  const box = (x, y, label, kind, sub) => {
    const st = styles[kind];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: EW, h: EH, fill: { color: st.fill }, line: { color: st.line, width: 1.25, dashType: st.dash }, rectRadius: 0.08 });
    s.addText([
      { text: label, options: { fontSize: 9.5, bold: true, color: st.txt, breakLine: true } },
      { text: sub || st.tag, options: { fontSize: 7, color: MUTED } },
    ], { x, y, w: EW, h: EH, margin: 0.04, isTextBox: true, fontFace: "Calibri", align: "center", valign: "middle" });
  };
  const line = (x1, y1, x2, y2, color, head, dash) => {
    const horiz = Math.abs(y2 - y1) < 0.001;
    const x = Math.min(x1, x2), y = Math.min(y1, y2);
    const w = horiz ? Math.abs(x2 - x1) : 0.001, h = horiz ? 0.001 : Math.abs(y2 - y1);
    const opts = { color, width: 1.5 };
    if (dash) opts.dashType = dash;
    if (head) { const endIsP2 = horiz ? x2 > x1 : y2 > y1; if (endIsP2) opts.endArrowType = "triangle"; else opts.beginArrowType = "triangle"; }
    s.addShape(pres.shapes.LINE, { x, y, w, h, line: opts });
  };
  // Polyline: Punkte, Pfeilspitze am letzten Segment
  const path = (pts, color, dash) => {
    for (let i = 0; i < pts.length - 1; i++) line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], color, i === pts.length - 2, dash);
  };
  const step = (x, y, w, n, text, color) => {
    const c = color || GREEN;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.03, w: 0.22, h: 0.22, fill: { color: c }, line: { color: c } });
    s.addText(String(n), { x, y: y + 0.03, w: 0.22, h: 0.22, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
    s.addText(text, { x: x + 0.27, y, w: w - 0.27, h: 0.28, margin: 0.02, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: INK, fill: { color: WHITE }, valign: "middle" });
  };

  // Beteiligte
  box(1.0, 1.5, "Erstprüfende", "person");
  box(1.0, 3.2, "Studierende", "person");
  box(1.0, 4.9, "Interne Zweitprüfende", "person");
  box(3.4, 4.9, "Externe Korreferenten", "person");
  box(5.6, 3.2, "Sekretariat / Dekanat MND", "person", "Drehscheibe, prüft und legt ab");
  box(8.7, 1.5, "Prüfungsausschuss-Vorsitz", "person");
  box(8.7, 3.2, "Studienausschuss WI", "person", "nur bei Verlängerung");

  // 1: Erstprüfende -> Studierende
  path([[1.95, 2.05], [1.95, 3.2]], GREEN);
  step(2.05, 2.45, 1.7, 1, "Formular per E-Mail");
  // 2: Studierende -> Zweitprüfende (intern)
  path([[1.95, 3.75], [1.95, 4.9]], GREEN);
  step(2.05, 4.15, 2.2, 2, "Antrag ausgefüllt, unterschrieben");
  // 2b: Studierende -> Externe Korreferenten
  path([[2.9, 3.65], [4.35, 3.65], [4.35, 4.9]], GREEN);
  step(3.0, 3.75, 2.0, "2", "Formular Ext. Prüfende + Urkunde");
  // 3: Zweitprüfende (intern + extern) -> Sekretariat
  line(1.95, 5.45, 1.95, 5.62, GREEN, false);
  line(4.35, 5.45, 4.35, 5.62, GREEN, false);
  path([[1.95, 5.62], [6.2, 5.62], [6.2, 3.75]], GREEN);
  step(6.3, 4.35, 2.3, 3, "unterschrieben, per E-Mail ans Dekanat");
  // 4: Sekretariat -> PA-Vorsitz
  path([[7.1, 3.2], [7.1, 1.775], [8.7, 1.775]], GREEN);
  step(7.2, 2.3, 1.55, 4, "geprüft: Zulassung, Qualifikation");
  // 5: PA-Vorsitz -> Erstprüfende (über den oberen Rand)
  path([[9.65, 1.5], [9.65, 1.3], [1.95, 1.3], [1.95, 1.5]], GREEN);
  step(4.4, 1.32, 3.2, 5, "Entscheidung PA; 1. Prüfer trägt Ausgabe ein (Beginn, Abgabe)");
  // 6: Erstprüfende -> Sekretariat
  path([[2.9, 1.9], [5.35, 1.9], [5.35, 3.35], [5.6, 3.35]], GREEN);
  step(3.0, 2.0, 1.8, 6, "Formular mit Fristen zurück");
  // 7: Sekretariat -> Studierende
  path([[5.6, 3.5], [2.9, 3.5]], GREEN);
  step(3.0, 3.05, 2.3, 7, "Info Thema + Frist; Formular bleibt in der Akte");
  // V: Verlängerung
  path([[7.5, 3.5], [8.7, 3.5]], ORANGE, "dash");
  step(7.55, 3.75, 1.2, "V", "Verlängerung", ORANGE);

  // Rahmenbedingungen (Systeme, Dokumente, Ereignisse ohne eigene Pfeile)
  const chips = [
    ["E-Mail (Outlook)", "system", "Transport aller Übergaben, keine Nachverfolgung"],
    ["THM-Intranet", "system", "PDF-Formulare nur nach Login"],
    ["Sprechzeiten Sekretariat", "event", "Rückfragen Mo bis Fr 9.30 bis 11.30 Uhr"],
    ["Papierakte", "document", "Formular wird abgeheftet, nicht an Studierende"],
  ];
  s.addText("Rahmenbedingungen im Kontext", { x: 0.5, y: 6.0, w: 3, h: 0.2, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, italic: true, color: MOSS });
  chips.forEach((c, i) => {
    const st = styles[c[1]], x = 0.5 + i * 2.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.22, w: 2.5, h: 0.5, fill: { color: st.fill }, line: { color: st.line, width: 1, dashType: st.dash }, rectRadius: 0.06 });
    s.addText([
      { text: c[0], options: { bold: true, fontSize: 8.5, color: st.txt, breakLine: true } },
      { text: c[2], options: { fontSize: 7, color: MUTED } },
    ], { x, y: 6.22, w: 2.5, h: 0.5, margin: 0.05, isTextBox: true, fontFace: "Calibri", valign: "middle" });
  });

  // Schwachstellen
  const GX = 11.2, GY = 0.98, GW = 1.65, GH = 5.74;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX, y: GY, w: GW, h: GH, fill: { color: WHITE }, line: { color: RED, width: 1, dashType: "sysDot" }, rectRadius: 0.2 });
  s.addText([
    { text: "Schwachstellen (Ist)", options: { bold: true, fontSize: 9.5, color: RED, breakLine: true } },
    { text: "Beobachtet aus Formular und THM-Seite", options: { fontSize: 7, color: GREY } },
  ], { x: GX + 0.1, y: GY + 0.1, w: GW - 0.2, h: 0.6, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top" });
  const weak = [
    ["Medienbrüche", "PDF ausfüllen, drucken, unterschreiben, scannen, mailen: sieben Übergaben pro Antrag."],
    ["Kein Status", "Studierende erfahren erst in Schritt 7 vom Ergebnis; Nachfragen nur telefonisch."],
    ["Manuelle Prüfung", "Zulassung, Qualifikation und Fristen prüft das Sekretariat von Hand (Schritte 3 und 6)."],
    ["Externe ohne Zugang", "Externe Korreferenten brauchen Zusatzformular, Post oder Scan (Schritt 2)."],
  ];
  weak.forEach((g, i) => {
    const y = GY + 0.8 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX + 0.1, y, w: GW - 0.2, h: 1.08, fill: { color: REDPALE }, line: { color: RED, width: 0.75, dashType: "dash" }, rectRadius: 0.06 });
    s.addText([
      { text: g[0], options: { bold: true, fontSize: 8, color: RED, breakLine: true } },
      { text: g[1], options: { fontSize: 7, color: MUTED } },
    ], { x: GX + 0.16, y: y + 0.05, w: GW - 0.32, h: 1.0, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.05 });
  });

  // Legende
  const ly = 6.9;
  const leg = (x, kind, text) => {
    const st = styles[kind];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ly + 0.05, w: 0.32, h: 0.2, fill: { color: st.fill }, line: { color: st.line, width: 1, dashType: st.dash }, rectRadius: 0.04 });
    s.addText(text, { x: x + 0.38, y: ly, w: 1.4, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
  };
  leg(0.5, "person", "Person / Rolle");
  leg(1.9, "system", "System im Betrieb");
  leg(3.4, "document", "Dokument");
  leg(4.5, "event", "Ereignis / Einschränkung");
  s.addShape(pres.shapes.LINE, { x: 6.2, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: GREEN, width: 1.5, endArrowType: "triangle" } });
  s.addText("Übergabe des Formulars, Schritte 1 bis 7", { x: 6.65, y: ly, w: 2.2, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: 8.9, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: ORANGE, width: 1.5, dashType: "dash", endArrowType: "triangle" } });
  s.addText("Verlängerungsantrag", { x: 9.35, y: ly, w: 1.3, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
  s.addText("Quelle: Zulassungsantrag WS23/24, Hinweise zur Bearbeitung", { x: 10.7, y: ly, w: 2.15, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 6.5, color: GREY, align: "right", valign: "middle" });

  s.addNotes(
    "Ist-Zustand: kein System, sondern der Laufweg des PDF-Antragsformulars. Die Schritte 1 bis 7 entsprechen den Hinweisen zur Bearbeitung im Zulassungsantrag: " +
    "1 Erstprüfer mailt das Formular, 2 Studierende füllen aus und unterschreiben (externe Korreferenten zusätzlich mit Formular Externe Prüfende und Urkunde), " +
    "3 Zweitprüfer unterschreibt und mailt ans Dekanat, 4 Sekretariat prüft Zulassung und Qualifikation, PA-Vorsitz entscheidet, 5 Erstprüfer trägt die Ausgabe ein, " +
    "6 Formular zurück ans Sekretariat, Fristprüfung und Ablage, 7 Sekretariat informiert Studierende. Verlängerung läuft als eigener Papierantrag an den Studienausschuss WI. " +
    "Die Abgrenzung des Systems erfolgt laut Vorlesung in der Soll-Perspektive (Folie 1); diese Folie belegt die Schwachstellen."
  );
})();

const out = process.env.OUT || "Kontextdiagramm_ThesisFlow_v2.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
