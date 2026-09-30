// Kontextdiagramm Ist-Zustand (Datenfluss-Notation): heutiger manueller Anmeldeprozess
// Gleiches Layout wie kontext2.js (Soll), aber ohne IT-System in der Mitte und mit
// Schwachstellen statt Grauzone. Quellen: Zulassungsantrag WS23/24 (Hinweise zur Bearbeitung),
// THM-Seite "Abschlussarbeit" FB MND, Anlage externe Pruefende.
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "Kontextdiagramm Ist-Zustand";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF", GREY = "8A968A", LGREY = "C9D1C9", PALE = "F5F7F4";
const IN = "2C5F2D", OUT = "4A7FB5", RED = "9B3B2E", REDPALE = "FBF1EF";
const s = pres.addSlide();
s.background = { color: WHITE };

// ---------- Kopf ----------
s.addText("Kontextdiagramm Ist-Zustand", { x: 0.5, y: 0.2, w: 8, h: 0.48, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 24, bold: true, color: INK, valign: "middle" });
s.addText("Heutiger Anmeldeprozess für Abschlussarbeiten im FB MND: PDF-Formular, E-Mail, Unterschriften, Papierakte, kein IT-System", { x: 0.5, y: 0.66, w: 9.5, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED });
s.addText("REQUIREMENTS ENGINEERING  ·  GRUPPE GRÜN  ·  STAND 30.09.2026", { x: 8.8, y: 0.28, w: 4.05, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

// ---------- Kontextgrenze ----------
const KX = 0.5, KY = 1.0, KW = 10.6, KH = 5.75;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: KX, y: KY, w: KW, h: KH, fill: { color: PALE }, line: { color: MOSS, width: 1.25, dashType: "dash" }, rectRadius: 0.2 });
s.addText("Kontextgrenze  ·  Ist-Prozess", { x: KX + 0.2, y: KY + 0.07, w: 3, h: 0.22, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, italic: true, bold: true, color: MOSS });

// ---------- Geometrie ----------
const EW = 1.9, EH = 0.55;
const LX = 0.8, LLX = 2.85, LLW = 1.9;
const C = { x: 4.9, y: 1.3, w: 2.0, h: 3.75 };
const RLX = 7.05, RLW = 1.9, RX = 9.1;

// ---------- Mitte: kein System, manueller Prozess ----------
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: C.x, y: C.y, w: C.w, h: C.h, fill: { color: "6B7F6B" }, line: { color: "4A5C4A", width: 1.5, dashType: "dash" }, rectRadius: 0.12, shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.25 } });
s.addText([
  { text: "Ist-Prozess", options: { fontFace: "Cambria", fontSize: 20, bold: true, color: WHITE, breakLine: true } },
  { text: "Anmeldung, Freigabe, Ausgabe und Verlängerung per PDF-Formular, E-Mail-Umlauf und Unterschriften in 7 Schritten", options: { fontFace: "Calibri", fontSize: 9, color: "E4EAE4", breakLine: true } },
  { text: " ", options: { fontSize: 6, breakLine: true } },
  { text: "kein System: manueller Ablauf, Sekretariat als Drehscheibe", options: { fontFace: "Calibri", fontSize: 8, italic: true, color: "D9E6CF" } },
], { x: C.x + 0.1, y: C.y, w: C.w - 0.2, h: C.h, margin: 0, isTextBox: true, align: "center", valign: "middle" });

// ---------- Elemente ----------
const styles = {
  person: { fill: TINT, line: GREEN, dash: "solid", txt: GREEN, tag: "Person / Rolle" },
  system: { fill: WHITE, line: GREY, dash: "dash", txt: MUTED, tag: "System im Betrieb" },
  document: { fill: WHITE, line: GREEN, dash: "sysDot", txt: GREEN, tag: "Dokument" },
  event: { fill: "FFF7E0", line: "C9A227", dash: "solid", txt: "7A5C00", tag: "Ereignis / Einschränkung" },
};
const el = (x, y, label, kind) => {
  const st = styles[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: EW, h: EH, fill: { color: st.fill }, line: { color: st.line, width: 1.25, dashType: st.dash }, rectRadius: 0.08 });
  s.addText([
    { text: label, options: { fontSize: 9.5, bold: true, color: st.txt, breakLine: true } },
    { text: st.tag, options: { fontSize: 7, color: MUTED } },
  ], { x, y, w: EW, h: EH, margin: 0.04, isTextBox: true, fontFace: "Calibri", align: "center", valign: "middle" });
};
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

// ---------- linke Spalte ----------
const left = [
  ["Studierende", "person", "Antrag ausgedruckt + unterschrieben; Formular ext. Prüfende + Urkunde; Verlängerungsantrag", "PDF-Formular per E-Mail vom 1. Prüfer; Info über Ausgabe; kein Status"],
  ["Erstprüfende", "person", "Formular per E-Mail an Studierende; Ausgabe eintragen + unterschreiben; Befürwortung Verlängerung", "Antrag vom PA-Vorsitz zurück (Papier/Scan)"],
  ["Interne Zweitprüfende", "person", "Unterschrift Korreferat; Weiterleitung per E-Mail an das Dekanat", "Antrag von Studierenden (Papier oder Scan)"],
  ["Externe Korreferenten", "person", "Unterschrift; Formular Externe Prüfende; Urkunde per Post/Scan", "Antrag von Studierenden; Rückfragen zur Qualifikation"],
  ["Sekretariat / Dekanat MND", "person", "Prüfung Zulassung + Qualifikation; Fristprüfung; Ablage; Info an Studierende", "Antrag per E-Mail an das Dekanat; Telefonrückfragen; unleserliche Scans"],
];
const ysL = [1.45, 2.2, 2.95, 3.7, 4.45];
left.forEach((e, i) => {
  const y = ysL[i], yc = y + EH / 2;
  el(LX, y, e[0], e[1]);
  seg(LX + EW, yc - 0.07, C.x, yc - 0.07, IN, true);
  seg(C.x, yc + 0.07, LX + EW, yc + 0.07, OUT, true);
  label(LLX, yc - 0.25, LLW, e[2], e[3]);
});

// ---------- rechte Spalte ----------
const right = [
  ["Prüfungsausschuss-Vorsitz", "person", "Entscheidung mit Unterschrift (Zulassung, ext. Prüfer)", "Antrag vom Dekanat (Papier)"],
  ["Studienausschuss WI", "person", "Entscheidung Verlängerung", "Verlängerungsantrag mit Befürwortung (Papier)"],
  ["THM-Website / Intranet", "system", "PDF-Formulare, nur nach Login", "Download durch Studierende"],
  ["E-Mail (Outlook)", "system", "Anhänge: Formular, Scans, Nachweise", "Versand zwischen allen Beteiligten, keine Nachverfolgung"],
];
const ysR = [1.8, 2.65, 3.5, 4.35];
right.forEach((e, i) => {
  const y = ysR[i], yc = y + EH / 2;
  el(RX, y, e[0], e[1]);
  seg(RX, yc - 0.07, C.x + C.w, yc - 0.07, IN, true);
  seg(C.x + C.w, yc + 0.07, RX, yc + 0.07, OUT, true);
  label(RLX, yc - 0.25, RLW, e[2], e[3]);
});

// ---------- unten ----------
const BY = 5.95, ELB = 5.35;
const bottom = [
  [2.6, "Prüfungsordnungen / Allg. Best.", "document", "Regeln und Fristen werden manuell nachgeschlagen", null, 5.35],
  [4.95, "Papierakte Studierende", "document", null, "Formular wird abgeheftet, geht nicht an Studierende", 5.9],
  [7.3, "Sprechzeiten Sekretariat", "event", "Rückfragen nur Mo bis Fr 9.30 bis 11.30 Uhr, Raum A2.1.01", null, 6.45],
];
bottom.forEach(b => {
  const [x, name, kind, into, out, sx] = b;
  el(x, BY, name, kind);
  const xc = x + EW / 2;
  if (Math.abs(xc - sx) < 0.01) {
    if (into) seg(xc, BY, xc, C.y + C.h, IN, true); else seg(xc, C.y + C.h, xc, BY, OUT, true);
  } else if (into) {
    seg(xc, BY, xc, ELB, IN, false); seg(xc, ELB, sx, ELB, IN, false); seg(sx, ELB, sx, C.y + C.h, IN, true);
  } else {
    seg(sx, C.y + C.h, sx, ELB, OUT, false); seg(sx, ELB, xc, ELB, OUT, false); seg(xc, ELB, xc, BY, OUT, true);
  }
  label(xc - 0.85, ELB + 0.07, 1.7, into, out);
});

// ---------- Schwachstellen ----------
const GX = 11.3, GY = 1.0, GW = 1.55, GH = 5.75;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX, y: GY, w: GW, h: GH, fill: { color: WHITE }, line: { color: RED, width: 1, dashType: "sysDot" }, rectRadius: 0.2 });
s.addText([
  { text: "Schwachstellen (Ist)", options: { bold: true, fontSize: 9.5, color: RED, breakLine: true } },
  { text: "Beobachtet aus Formular und THM-Seite", options: { fontSize: 7, color: GREY } },
], { x: GX + 0.1, y: GY + 0.1, w: GW - 0.2, h: 0.6, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top" });
const weak = [
  ["Medienbrüche", "PDF ausfüllen, drucken, unterschreiben, scannen, mailen: sieben Übergaben pro Antrag."],
  ["Kein Status", "Studierende erfahren erst am Ende vom Ergebnis; Nachfragen nur telefonisch."],
  ["Manuelle Prüfung", "Zulassung, Qualifikation und Fristen werden im Sekretariat von Hand geprüft."],
  ["Externe ohne Zugang", "Externe Korreferenten brauchen Zusatzformular, Post oder Scan."],
];
weak.forEach((g, i) => {
  const y = GY + 0.8 + i * 1.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX + 0.1, y, w: GW - 0.2, h: 1.08, fill: { color: REDPALE }, line: { color: RED, width: 0.75, dashType: "dash" }, rectRadius: 0.06 });
  s.addText([
    { text: g[0], options: { bold: true, fontSize: 8, color: RED, breakLine: true } },
    { text: g[1], options: { fontSize: 7, color: MUTED } },
  ], { x: GX + 0.16, y: y + 0.05, w: GW - 0.32, h: 1.0, margin: 0, isTextBox: true, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.05 });
});

// ---------- Legende ----------
const ly = 6.9;
const leg = (x, kind, text) => {
  const st = styles[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ly + 0.05, w: 0.32, h: 0.2, fill: { color: st.fill }, line: { color: st.line, width: 1, dashType: st.dash }, rectRadius: 0.04 });
  s.addText(text, { x: x + 0.38, y: ly, w: 1.5, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
};
leg(0.5, "person", "Person / Rolle");
leg(1.9, "system", "System im Betrieb");
leg(3.4, "document", "Dokument");
leg(4.5, "event", "Ereignis / Einschränkung");
s.addShape(pres.shapes.LINE, { x: 6.2, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: IN, width: 1.25, endArrowType: "triangle" } });
s.addText("Input in den Prozess", { x: 6.65, y: ly, w: 1.3, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.LINE, { x: 8.0, y: ly + 0.15, w: 0.4, h: 0.001, line: { color: OUT, width: 1.25, beginArrowType: "triangle" } });
s.addText("Output aus dem Prozess", { x: 8.45, y: ly, w: 1.5, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: MUTED, valign: "middle" });
s.addText("Quellen: Zulassungsantrag WS23/24 (Hinweise zur Bearbeitung), THM-Seite Abschlussarbeit FB MND, Anlage ext. Prüfer", { x: 10.0, y: ly, w: 2.85, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 6.5, color: GREY, align: "right", valign: "middle" });

s.addNotes(
  "Kontextdiagramm des Ist-Zustands: heutiger manueller Anmeldeprozess ohne IT-System. Dient dem Vergleich mit dem Soll-Kontextdiagramm ThesisFlow. " +
  "Laut Vorlesung Kap. 2 erfolgt die Abgrenzung in der Soll-Perspektive; das Ist-Bild macht die Schwachstellen sichtbar, die ThesisFlow beseitigen soll.\n" +
  "Ablauf (Hinweise zur Bearbeitung im Zulassungsantrag): 1. Prüfer mailt Formular, Studierende füllen aus und unterschreiben, 2. Prüfer unterschreibt, Dekanat prüft, PA-Vorsitz entscheidet, 1. Prüfer gibt aus, Dekanat prüft Frist und legt ab, Dekanat informiert Studierende."
);

const out = process.env.OUT || "Kontextdiagramm_Ist-Zustand.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
