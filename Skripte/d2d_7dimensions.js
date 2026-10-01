// 7 Dimensions (Discover to Deliver) fuer genau eine User Story: US 4.1.1
// "Als Erstpruefer:in moechte ich das Thema der Abschlussarbeit bestaetigen, damit die Arbeit
//  mit dem abgestimmten Thema offiziell freigegeben wird."  (Story Mapping/4-Ausgabe-der-Arbeit.md)
// Interface wird vom Projektteam ergaenzt: nur Platzhalterfolie, kein Mockup.
// Aufruf: OUT="../D2D User Story/7_Dimensions_Erstpruefer_Themenfreigabe.pptx" node d2d_7dimensions.js
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "7 Dimensions – Erstprüfer:in: Themenfreigabe";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF",
  PALE = "D9E6CF", GREY = "8A968A", LGREY = "C9D1C9", BG = "F5F7F4", AMBER = "C9A227", AMBERPALE = "FFF7E0", BROWN = "7A5C00";
const STORY = "Als Erstprüfer:in möchte ich das Thema der Abschlussarbeit bestätigen, damit die Arbeit mit dem abgestimmten Thema offiziell freigegeben wird.";
const THREAD = ["Thema aufrufen", "Thema prüfen", "Thema bestätigen", "Thema freigeben"];

// ---------- Helfer ----------
const T = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: "Calibri", color: INK, valign: "top" }, o));
const rbox = (s, x, y, w, h, fill, line, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 1 }, rectRadius: 0.08 }, o || {}));
const rect = (s, x, y, w, h, fill, line, lw) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: lw || 0.75 } });
const circle = (s, x, y, d, fill, line, lw) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: line || fill, width: lw || 1 } });
const seg = (s, x1, y1, x2, y2, head, color, width) => {
  const horiz = Math.abs(y2 - y1) < 0.001;
  const x = Math.min(x1, x2), y = Math.min(y1, y2);
  const line = { color: color || INK, width: width || 1.25 };
  if (head) { const end = horiz ? x2 > x1 : y2 > y1; if (end) line.endArrowType = "triangle"; else line.beginArrowType = "triangle"; }
  s.addShape(pres.shapes.LINE, { x, y, w: horiz ? Math.abs(x2 - x1) : 0.001, h: horiz ? 0.001 : Math.abs(y2 - y1), line });
};
const slide = (n, title, source) => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  T(s, title, { x: 0.45, y: 0.22, w: 6.5, h: 0.5, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  T(s, n ? `7 DIMENSIONS  ·  ${n} VON 7` : "7 DIMENSIONS", { x: 6.6, y: 0.3, w: 2.95, h: 0.3, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });
  T(s, "US 4.1.1  ·  Erstprüfer:in:  " + THREAD.join("  ›  "), { x: 0.45, y: 0.72, w: 9.1, h: 0.22, fontSize: 9, color: MUTED });
  T(s, "ThesisFlow  ·  Requirements Engineering, Gruppe Grün", { x: 0.45, y: 5.28, w: 4.2, h: 0.2, fontSize: 7.5, color: GREY });
  if (source) T(s, source, { x: 4.6, y: 5.28, w: 4.95, h: 0.2, fontSize: 7.5, italic: true, color: GREY, align: "right" });
  return s;
};

// =====================================================================
// Folie 1: Titel + User Story
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 3.7, 5.625, GREEN);
  T(s, "7 DIMENSIONS  ·  DISCOVER TO DELIVER", { x: 0.45, y: 0.42, w: 3.0, h: 0.25, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, "User Story 4.1.1", { x: 0.45, y: 0.78, w: 3.0, h: 0.5, fontFace: "Cambria", fontSize: 26, bold: true, color: WHITE, valign: "middle" });
  T(s, "Thema bestätigen und freigeben", { x: 0.45, y: 1.32, w: 3.0, h: 0.7, fontFace: "Cambria", fontSize: 17, italic: true, color: MOSS });
  T(s, "EINORDNUNG IM STORY MAPPING", { x: 0.45, y: 2.55, w: 3.0, h: 0.22, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, [
    { text: "Activity 4: ", options: { bold: true, color: PALE } }, { text: "Ausgabe der Arbeit", options: { color: WHITE, breakLine: true } },
    { text: "User Task 4.1: ", options: { bold: true, color: PALE } }, { text: "Thema bestätigen und freigeben", options: { color: WHITE, breakLine: true } },
    { text: "Rolle: ", options: { bold: true, color: PALE } }, { text: "Erstprüfer:in", options: { color: WHITE } },
  ], { x: 0.45, y: 2.82, w: 3.0, h: 0.85, fontSize: 10, paraSpaceAfter: 4 });
  T(s, "ThesisFlow  ·  Abschlussarbeiten-Portal FB MND\nRequirements Engineering, Gruppe Grün", { x: 0.45, y: 4.85, w: 3.0, h: 0.45, fontSize: 8, color: PALE, valign: "bottom" });

  T(s, "Betrachtete User Story", { x: 4.0, y: 0.42, w: 5.5, h: 0.45, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  rbox(s, 4.0, 1.0, 5.55, 1.25, TINT);
  T(s, "„" + STORY + "“", { x: 4.2, y: 1.0, w: 5.15, h: 1.25, fontFace: "Cambria", fontSize: 14, italic: true, color: GREEN, valign: "middle" });
  T(s, "Quelle: Story Mapping/4-Ausgabe-der-Arbeit.md", { x: 4.0, y: 2.3, w: 5.55, h: 0.2, fontSize: 7.5, italic: true, color: GREY });

  T(s, "Roter Faden der Präsentation", { x: 4.0, y: 2.75, w: 5.5, h: 0.3, fontSize: 11, bold: true, color: GREEN, valign: "middle" });
  const cw = 1.3, gap = 0.115;
  THREAD.forEach((t, i) => {
    const x = 4.0 + i * (cw + gap);
    rbox(s, x, 3.12, cw, 0.5, GREEN);
    T(s, t, { x, y: 3.12, w: cw, h: 0.5, fontSize: 9.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    if (i < THREAD.length - 1) seg(s, x + cw, 3.37, x + cw + gap, 3.37, true, GREEN, 1.5);
  });

  T(s, "Die sieben Dimensionen", { x: 4.0, y: 3.95, w: 5.5, h: 0.3, fontSize: 11, bold: true, color: GREEN, valign: "middle" });
  const dims = ["User", "Interface", "Action", "Data", "Control", "Environment", "Quality Attribute"];
  const dw = [0.58, 0.76, 0.64, 0.56, 0.68, 0.94, 1.09];
  let dx = 4.0;
  dims.forEach((d, i) => {
    rbox(s, dx, 4.32, dw[i], 0.36, TINT, MOSS, { rectRadius: 0.05 });
    T(s, [{ text: (i + 1) + " ", options: { bold: true, color: GREEN } }, { text: d }], { x: dx, y: 4.32, w: dw[i], h: 0.36, fontSize: 8, align: "center", valign: "middle" });
    dx += dw[i] + 0.05;
  });
  s.addNotes("Betrachtet wird genau eine User Story: US 4.1.1 aus dem Story Mapping (Activity 4 Ausgabe der Arbeit, User Task 4.1 Thema bestätigen und freigeben). Formulierung wörtlich aus dem Repository. Roter Faden: Thema aufrufen, prüfen, bestätigen, freigeben.");
}

// =====================================================================
// Folie 2: User (zwei Personas)
// =====================================================================
{
  const s = slide(1, "1. User", "vgl. Vorlesung RE, Kap. 3 Stakeholder; Praxisteil, Personas");
  const persona = (x, p) => {
    const w = 4.49, y = 1.08;
    rbox(s, x, y, w, 4.07, TINT);
    rbox(s, x, y, w, 0.62, GREEN);
    rect(s, x, y + 0.3, w, 0.32, GREEN);
    T(s, p.name, { x: x + 0.15, y: y + 0.04, w: w - 0.3, h: 0.32, fontFace: "Cambria", fontSize: 14, bold: true, color: WHITE, valign: "middle" });
    T(s, p.role, { x: x + 0.15, y: y + 0.34, w: w - 0.3, h: 0.24, fontSize: 9, color: PALE, valign: "middle" });
    p.rows.forEach((r, i) => {
      const ry = y + 0.72 + i * 0.55;
      T(s, r[0], { x: x + 0.15, y: ry, w: 1.05, h: 0.5, fontSize: 8.5, bold: true, color: GREEN });
      T(s, r[1], { x: x + 1.22, y: ry, w: w - 1.37, h: 0.5, fontSize: 8.5 });
      if (i < p.rows.length - 1) s.addShape(pres.shapes.LINE, { x: x + 0.15, y: ry + 0.51, w: w - 0.3, h: 0.001, line: { color: PALE, width: 0.75 } });
    });
  };
  persona(0.45, {
    name: "Prof. Dr. Andreas Becker", role: "Persona B  ·  Erstprüfer:in (THM), handelnde Rolle der Story",
    rows: [
      ["Kontext", "Hat Thema und Starttermin mit der Studierenden abgestimmt. Der Antrag wurde vom Prüfungsausschuss angenommen."],
      ["Ziel", "Das Thema bestätigen, damit die Arbeit mit dem abgestimmten Thema offiziell freigegeben wird."],
      ["Aufgaben", "Antrag aufrufen, Thema prüfen, Thema bestätigen und freigeben."],
      ["Erwartungen", "Thema eindeutig einem Antrag zugeordnet. Freigabe ohne Papierformular, am Desktop und am Smartphone."],
      ["Pain Points", "Heute: unteren Formularteil ausfüllen, unterschreiben und per E-Mail an das Dekanat senden."],
      ["Wissen", "Kennt Prüfungsordnung und Ablauf. Erfahrung mit dem Portal: noch zu validieren."],
    ],
  });
  persona(5.06, {
    name: "Moritz Hoffmann", role: "Persona A  ·  Studierende:r B.Sc. Wirtschaftsinformatik, betroffene Rolle",
    rows: [
      ["Kontext", "Erste Abschlussarbeit. Hat das Thema mit der Erstprüfer:in abgestimmt und den Antrag eingereicht."],
      ["Ziel", "Thema und Bearbeitungsfrist verbindlich kennen und mit der Arbeit beginnen."],
      ["Aufgaben", "In dieser Story keine eigene Aktion. Wartet auf die Themenfreigabe."],
      ["Erwartungen", "Bearbeitungsstand des Antrags jederzeit sehen. Über die Ausgabe des Themas informiert werden."],
      ["Pain Points", "Heute: wochenlang keine Rückmeldung zum Antrag, unsicher, ob alles vollständig ist."],
      ["Wissen", "Unsicher bei Formalien, im Umgang mit digitalen Anwendungen geübt."],
    ],
  });
  s.addNotes("Persona A stammt aus der Aufgabe Personas (Aufgabe 2). Persona B ist aus dem Ist-Prozess im Zulassungsantrag abgeleitet (1. Prüfer(in) füllt den unteren Formularteil aus und sendet an das Dekanat); Name und Systemerfahrung sind fiktiv und zu validieren. Bezüge: US 1.1.1 (Thema abstimmen), US 3.2.3 (Entscheidung mitgeteilt bekommen), US 4.3.2 (über Ausgabe informiert werden).");
}

// =====================================================================
// Folie 3: Interface (Platzhalter, wird vom Projektteam ergaenzt)
// =====================================================================
{
  const s = slide(2, "2. Interface", "vgl. Praxisteil, D2D Interface: GUI-Design, Dialoglandkarte");
  rbox(s, 0.45, 1.1, 9.1, 4.0, BG, LGREY, { line: { color: GREY, width: 1.25, dashType: "dash" }, rectRadius: 0.12 });
  T(s, "Interface-Entwurf wird durch das Projektteam ergänzt.", { x: 0.45, y: 2.55, w: 9.1, h: 0.5, fontFace: "Cambria", fontSize: 18, bold: true, color: MUTED, align: "center", valign: "middle" });
  T(s, "Platzhalter für den Dialog, in dem die Erstprüfer:in das Thema aufruft, prüft, bestätigt und freigibt.", { x: 0.45, y: 3.1, w: 9.1, h: 0.35, fontSize: 10.5, color: GREY, align: "center", valign: "middle" });
  s.addNotes("Diese Dimension erstellt das Projektteam selbst. Hier bewusst kein UI, kein Mockup, keine Schaltflächen.");
}

// =====================================================================
// Folie 4: Action (BPMN)
// =====================================================================
{
  const s = slide(3, "3. Action", "vgl. Vorlesung RE, Kap. 6 Funktionsperspektive; Darstellung als BPMN");
  const PX = 0.45, PY = 1.08, PW = 9.1, BAND = 0.3;
  const H1 = 1.95, H2 = 1.5, PH = H1 + H2;
  // Pool und Lanes
  rect(s, PX, PY, PW, PH, WHITE, INK, 1.25);
  rect(s, PX, PY, BAND, PH, TINT, INK, 1.25);
  rect(s, PX + BAND, PY, BAND, H1, TINT, INK, 1);
  rect(s, PX + BAND, PY + H1, BAND, H2, TINT, INK, 1);
  s.addShape(pres.shapes.LINE, { x: PX + BAND, y: PY + H1, w: PW - BAND, h: 0.001, line: { color: INK, width: 1 } });
  const vtext = (x, y, bw, lh, text, bold) => T(s, text, { x: x + bw / 2 - lh / 2, y: y + lh / 2 - bw / 2, w: lh, h: bw, rotate: 270, fontSize: 9, bold: !!bold, color: GREEN, align: "center", valign: "middle" });
  vtext(PX, PY, BAND, PH, "ThesisFlow: Themenfreigabe", true);
  vtext(PX + BAND, PY, BAND, H1, "Erstprüfer:in");
  vtext(PX + BAND, PY + H1, BAND, H2, "Portal");

  const TW = 1.3, TH = 0.52;
  const task = (x, cy, text, h) => { h = h || TH; rbox(s, x, cy - h / 2, TW, h, WHITE, INK, { line: { color: INK, width: 1.25 }, rectRadius: 0.1 }); T(s, text, { x: x + 0.05, y: cy - h / 2, w: TW - 0.1, h, fontSize: 8.5, align: "center", valign: "middle" }); };
  const THP = 0.76;
  const cyA = PY + 0.5, cyE = PY + 1.3, cyP = PY + H1 + 0.55;
  const xStart = 1.22, x1 = 1.75, x2 = 3.3, xG = 4.98, x3 = 5.4, xE1 = 6.9, x4 = 7.45, xE2 = 8.95;

  // Startereignis (Portal): Thema liegt vor
  circle(s, xStart, cyP - 0.14, 0.28, WHITE, GREEN, 1.25);
  T(s, "Thema von Studierende:r eingereicht", { x: xStart - 0.1, y: cyP + 0.42, w: 1.2, h: 0.36, fontSize: 7, color: MUTED, align: "left" });
  seg(s, xStart + 0.28, cyP, x1, cyP, true);
  task(x1, cyP, "Erstprüfer:in informieren");
  seg(s, x1 + TW / 2, cyP - TH / 2, x1 + TW / 2, cyE + TH / 2, true);
  task(x1, cyE, "Thema aufrufen");
  seg(s, x1 + TW, cyE, x2, cyE, true);
  task(x2, cyE, "Thema prüfen");
  seg(s, x2 + TW, cyE, xG - 0.22, cyE, true);
  // Exklusives Gateway
  s.addShape(pres.shapes.DIAMOND, { x: xG - 0.22, y: cyE - 0.22, w: 0.44, h: 0.44, fill: { color: AMBERPALE }, line: { color: INK, width: 1.25 } });
  T(s, "X", { x: xG - 0.22, y: cyE - 0.22, w: 0.44, h: 0.44, fontSize: 11, bold: true, align: "center", valign: "middle" });
  T(s, "Thema wie abgestimmt?", { x: xG - 0.75, y: cyE + 0.25, w: 1.5, h: 0.2, fontSize: 7.5, italic: true, color: MUTED, align: "center" });
  // ja: bestaetigen und freigeben -> Portal speichert und leitet weiter
  seg(s, xG + 0.22, cyE, x3, cyE, true);
  T(s, "ja", { x: xG + 0.2, y: cyE - 0.2, w: 0.2, h: 0.18, fontSize: 7.5, color: MUTED });
  task(x3, cyE, "Thema bestätigen und freigeben");
  seg(s, x3 + TW / 2, cyE + TH / 2, x3 + TW / 2, cyP - THP / 2, true);
  task(x3, cyP, "Themenfreigabe speichern, Antrag weiterleiten", THP);
  seg(s, x3 + TW, cyP, xE1, cyP, true);
  circle(s, xE1, cyP - 0.15, 0.3, WHITE, INK, 3);
  T(s, "Thema freigegeben", { x: xE1 - 0.4, y: cyP + 0.46, w: 1.1, h: 0.2, fontSize: 7.5, bold: true, align: "center", valign: "middle" });
  // nein: Aenderungswunsch eintragen -> Portal speichert und informiert Studierende:n
  seg(s, xG, cyE - 0.22, xG, cyA, false);
  seg(s, xG, cyA, x4, cyA, true);
  T(s, "nein", { x: xG + 0.05, y: cyA - 0.2, w: 0.4, h: 0.18, fontSize: 7.5, color: MUTED });
  task(x4, cyA, "Änderungswunsch eintragen");
  seg(s, x4 + TW / 2, cyA + TH / 2, x4 + TW / 2, cyP - THP / 2, true);
  task(x4, cyP, "Änderungswunsch speichern, Studierende:n informieren", THP);
  seg(s, x4 + TW, cyP, xE2, cyP, true);
  circle(s, xE2, cyP - 0.15, 0.3, WHITE, INK, 3);
  T(s, "Rückfrage gesendet", { x: xE2 - 0.8, y: cyP + 0.46, w: 1.15, h: 0.2, fontSize: 7.5, bold: true, align: "right", valign: "middle" });

  rbox(s, 0.45, 4.67, 9.1, 0.46, AMBERPALE, AMBER);
  T(s, [
    { text: "Ausschnitt: ", options: { bold: true, color: BROWN } }, { text: "nur US 4.1.1.   " },
    { text: "Davor: ", options: { bold: true, color: BROWN } }, { text: "Studierende:r reicht den Antrag mit Thema ein (Activity 2).   " },
    { text: "Danach: ", options: { bold: true, color: BROWN } }, { text: "Bearbeitungszeit festlegen (User Task 4.2)." },
  ], { x: 0.57, y: 4.67, w: 8.86, h: 0.46, fontSize: 8.5, valign: "middle" });
  s.addNotes("BPMN-Ausschnitt nur für US 4.1.1. Start: Die oder der Studierende hat den Antrag mit Thema eingereicht; das Portal informiert die Erstprüfer:in. Erstprüfer:in ruft das Thema auf und prüft es. Ja-Zweig: Thema bestätigen und freigeben, das Portal speichert die Themenfreigabe und leitet den Antrag weiter (Regel C6). Nein-Zweig: Erstprüfer:in trägt einen Änderungswunsch ein, das Portal speichert ihn und informiert die Studierende:n. Systemschritte liegen in der Lane Portal, Endereignisse sind als Zustände benannt. Editierbare Quelle: quellen/Action_Themenfreigabe.bpmn.");
}

// =====================================================================
// Folie 5: Data (UML-Klassendiagramm)
// =====================================================================
{
  const s = slide(4, "4. Data", "vgl. Vorlesung RE, Kap. 6 Strukturperspektive, UML-Klassendiagramm");
  const cls = (x, y, w, name, attrs) => {
    const hn = 0.32, ha = Math.max(attrs.length * 0.2 + 0.12, 0.4), h = hn + ha;
    rect(s, x, y, w, h, WHITE, INK, 1.25);
    rect(s, x, y, w, hn, TINT, INK, 1.25);
    T(s, name, { x, y, w, h: hn, fontSize: 10, bold: true, align: "center", valign: "middle" });
    T(s, attrs.map((a, i) => ({ text: a, options: { breakLine: i < attrs.length - 1 } })), { x: x + 0.08, y: y + hn + 0.05, w: w - 0.16, h: ha - 0.08, fontSize: 8, valign: "top" });
    return { x, y, w, h, b: y + h };
  };
  const W = 2.2, Y1 = 1.12, Y2 = 3.45;
  const stud = cls(0.5, Y1, W, "Studierende:r", ["- name : String", "- matrikelnummer : String", "- studiengang : String"]);
  const antr = cls(3.9, Y1, W, "Antrag", ["- studienabschluss : Abschlussart", "- gewuenschterStarttermin : Datum", "- zulassung : Entscheidung", "- status : Antragsstatus"]);
  const erst = cls(7.3, Y1, W, "Erstprüfer:in", ["- name : String"]);
  const thema = cls(2.3, Y2, W, "Thema", ["- wortlaut : String"]);
  const frei = cls(5.5, Y2, W, "Themenfreigabe", ["- zeitpunkt : Datum", "- vorschlagBestaetigt : Boolean"]);
  const mult = (x, y, t, al) => T(s, t, { x, y, w: 0.45, h: 0.18, fontSize: 8.5, bold: true, align: al || "left", valign: "middle" });
  const lab = (x, y, t, w) => T(s, t, { x, y, w: w || 1.2, h: 0.18, fontSize: 8, italic: true, color: MUTED, align: "center", valign: "middle" });
  // Studierende:r – Antrag
  const ya = Y1 + 0.55;
  seg(s, stud.x + W, ya, antr.x, ya, false);
  lab(stud.x + W, ya - 0.22, "stellt");
  mult(stud.x + W + 0.05, ya + 0.03, "1"); mult(antr.x - 0.5, ya + 0.03, "1..*", "right");
  // Antrag – Erstpruefer:in
  seg(s, antr.x + W, ya, erst.x, ya, false);
  lab(antr.x + W, ya - 0.22, "ist zugeordnet");
  mult(antr.x + W + 0.05, ya + 0.03, "0..*"); mult(erst.x - 0.5, ya + 0.03, "1", "right");
  // Antrag <>– Thema (Komposition)
  const xt = 4.2;
  seg(s, xt, antr.b + 0.2, xt, Y2, false);
  s.addShape(pres.shapes.DIAMOND, { x: xt - 0.08, y: antr.b, w: 0.16, h: 0.22, fill: { color: INK }, line: { color: INK, width: 1 } });
  mult(xt - 0.55, antr.b + 0.2, "1", "right"); mult(xt - 0.55, Y2 - 0.2, "1", "right");
  T(s, "enthält", { x: xt - 1.0, y: (antr.b + Y2) / 2 - 0.02, w: 0.9, h: 0.18, fontSize: 8, italic: true, color: MUTED, align: "right", valign: "middle" });
  // Antrag – Themenfreigabe
  const xf = 5.8;
  seg(s, xf, antr.b, xf, Y2, false);
  mult(xf + 0.06, antr.b + 0.03, "1"); mult(xf + 0.06, Y2 - 0.2, "0..1");
  T(s, "hat", { x: xf + 0.06, y: (antr.b + Y2) / 2 - 0.05, w: 0.5, h: 0.18, fontSize: 8, italic: true, color: MUTED, valign: "middle" });
  // Erstpruefer:in – Themenfreigabe
  const xe = 7.5;
  seg(s, xe, erst.b, xe, Y2, false);
  mult(xe + 0.06, erst.b + 0.03, "1"); mult(xe + 0.06, Y2 - 0.2, "0..*");
  T(s, "erteilt", { x: xe + 0.06, y: (erst.b + Y2) / 2 - 0.05, w: 0.6, h: 0.18, fontSize: 8, italic: true, color: MUTED, valign: "middle" });

  rbox(s, 0.45, 4.5, 9.1, 0.62, TINT);
  T(s, [
    { text: "Lesart: ", options: { bold: true, color: GREEN } },
    { text: "Ein Antrag enthält genau ein Thema (Komposition) und hat höchstens eine Themenfreigabe. Sie wird von genau einer Erstprüfer:in erteilt, die dem Antrag zugeordnet ist. Wertebereich von „status“ noch zu validieren." },
  ], { x: 0.6, y: 4.5, w: 8.8, h: 0.62, fontSize: 9, valign: "middle" });
  s.addNotes("Minimales UML-Klassendiagramm nur für US 4.1.1. Attribute aus dem Zulassungsantrag: Name, Matrikelnummer, Studiengang, Studienabschluss (Bachelor/Master), gewünschter Starttermin, Thema, 1. Prüfer(in), Entscheidung des Prüfungsausschusses (angenommen/abgelehnt), Bestätigung der Vorschläge des Kandidaten, Zeitpunkt der Ausgabe. Zu validieren: Multiplizität Studierende:r zu Antrag (1..*: Bachelor und Master oder erneuter Antrag) und der Wertebereich des Antragsstatus. Editierbare Quelle: quellen/Data_Klassendiagramm_Themenfreigabe.puml.");
}

// =====================================================================
// Folie 6: Control (Regeln nach Satzschablone)
// =====================================================================
{
  const s = slide(5, "5. Control", "vgl. Vorlesung RE, Kap. 5 Satzschablonen");
  const rules = [
    ["Das System muss der Erstprüfer:in die Möglichkeit bieten, den ihr zugeordneten Antrag mit dem Thema aufzurufen.", "User Story 4.1.1", true],
    ["Falls der Prüfungsausschuss den Antrag angenommen hat, muss das System der Erstprüfer:in die Möglichkeit bieten, das Thema zu bestätigen und freizugeben.", "Zulassungsantrag, Schritte 4 und 5; US 3.2.1", true],
    ["Das System muss ausschließlich der dem Antrag zugeordneten Erstprüfer:in die Möglichkeit bieten, das Thema freizugeben.", "Zulassungsantrag: 1. Prüfer(in) gibt aus", true],
    ["Sobald die Erstprüfer:in das Thema freigegeben hat, muss das System die Themenfreigabe mit Thema und Zeitpunkt speichern.", "Allg. PO THM § 17 (3), laut Zulassungsantrag", true],
    ["Falls die Erstprüfer:in das Thema nicht bestätigt, muss das System ihr die Möglichkeit bieten, einen Änderungswunsch einzutragen.", "Aktivitätsdiagramm des Teams", true],
    ["Sobald die Themenfreigabe gespeichert ist, muss das System den Antrag für das Festlegen der Bearbeitungszeit bereitstellen.", "abgeleitet aus User Task 4.1 und 4.2, zu validieren", false],
  ];
  const y0 = 1.08, rh = 0.54, gap = 0.06;
  rules.forEach((r, i) => {
    const y = y0 + i * (rh + gap);
    rbox(s, 0.45, y, 9.1, rh, TINT);
    rbox(s, 0.55, y + 0.1, 0.5, rh - 0.2, GREEN, GREEN, { rectRadius: 0.05 });
    T(s, "C" + (i + 1), { x: 0.55, y: y + 0.1, w: 0.5, h: rh - 0.2, fontSize: 10.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, r[0], { x: 1.2, y, w: 5.95, h: rh, fontSize: 9.5, valign: "middle" });
    T(s, r[1], { x: 7.3, y, w: 2.15, h: rh, fontSize: 7.5, italic: true, color: r[2] ? MUTED : BROWN, valign: "middle" });
  });
  T(s, [
    { text: "Satzschablone:  ", options: { bold: true, color: GREEN } },
    { text: "[Wann? Unter welcher Bedingung?]  muss das System  [wem?]  die Möglichkeit bieten,  [Objekt]  [Prozesswort]." },
  ], { x: 0.45, y: 4.72, w: 9.1, h: 0.4, fontSize: 9, color: MUTED, valign: "middle" });
  s.addNotes("Sechs Regeln, je eine Aussage, formuliert nach der Satzschablone der Vorlesung (Kap. 5). C1 bis C5 sind durch Repository-Artefakte belegt. C6 ist eine logische Ableitung aus der Reihenfolge der User Tasks 4.1 und 4.2 und noch zu validieren. Zu klären: Das Aktivitätsdiagramm des Teams zeigt die Freigabe der Prüfenden vor der Prüfung durch das Dekanat, das Story Mapping erst nach der Entscheidung des Prüfungsausschusses. C2 folgt dem Story Mapping und dem Zulassungsantrag.");
}

// =====================================================================
// Folie 7: Environment (Systemkontext, Soll-Perspektive)
// =====================================================================
{
  const s = slide(6, "6. Environment", "vgl. Vorlesung RE, Kap. 2 System und Systemkontext abgrenzen");
  const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: GREEN }, fontSize: 9 } });
  const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 8.5, color: INK }, o || {}) });
  const b = (t) => c(t, { bold: true, color: GREEN, fill: { color: TINT } });
  const rows = [
    [hdr("Bereich"), hdr("Relevante Umgebung"), hdr("Bedeutung für die User Story")],
    [b("Personen"), c("Erstprüfer:in, Studierende:r, Prüfungsausschuss-Vorsitz, Sekretariat / Dekanat MND"), c("Die Erstprüfer:in handelt. Die Studierende:r hat das Thema vorgeschlagen. Der Ausschuss entscheidet vorher, das Sekretariat prüft danach die Bearbeitungsfrist.")],
    [b("Systeme"), c("THM-Login für die Anmeldung, E-Mail-Dienst der THM (beide laut Kontextdiagramm)"), c("Die Anmeldung identifiziert die Erstprüfer:in. Der Weg der Benachrichtigung ist noch zu validieren.")],
    [b("Prozesse"), c("Activity 4 „Ausgabe der Arbeit“ im Anmeldeprozess"), c("Die Themenfreigabe liegt zwischen der Entscheidung des Prüfungsausschusses und dem Festlegen der Bearbeitungszeit.")],
    [b("Ereignisse"), c("Antrag vom Prüfungsausschuss angenommen; Thema freigegeben"), c("Das erste Ereignis löst die Story aus, das zweite beendet sie.")],
    [b("Dokumente"), c("Antrag mit Thema; Prüfungsordnungen B.Sc. und M.Sc.; Allg. PO der THM § 17 (3)"), c("Der Antrag liefert das zu prüfende Thema. Die Prüfungsordnung verlangt, Thema und Zeitpunkt der Ausgabe aktenkundig zu machen.")],
    [b("Technik"), c("Webbasierte Anwendung im Browser, am Desktop und am Smartphone"), c("Die Themenfreigabe muss auf beiden Gerätetypen komfortabel möglich sein.")],
  ];
  s.addTable(rows, { x: 0.45, y: 1.08, w: 9.1, colW: [1.15, 3.35, 4.6], rowH: [0.32, 0.66, 0.58, 0.58, 0.5, 0.66, 0.5], fontFace: "Calibri", border: { type: "solid", color: LGREY, pt: 0.75 }, valign: "middle", margin: [3, 6, 3, 6] });
  s.addNotes("Systemkontext in der Soll-Perspektive nach Kap. 2 der Vorlesung: Personen, Systeme im Betrieb, Prozesse, Ereignisse, Dokumente; ergänzt um die technische Umgebung aus dem Projektbriefing (webbasiert, Desktop und Smartphone). THM-Login und E-Mail-Dienst stehen im Kontextdiagramm des Teams (Aufgabe 3); das Aktivitätsdiagramm nennt die Anmeldung mit THM-Konto. Weitere Systeme werden nicht angenommen.");
}

// =====================================================================
// Folie 8: Quality Attribute
// =====================================================================
{
  const s = slide(7, "7. Quality Attribute", "vgl. Vorlesung RE, Kap. 1 Qualitätsanforderungen");
  const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: GREEN }, fontSize: 9 } });
  const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 9, color: INK }, o || {}) });
  const b = (t) => c(t, { bold: true, color: GREEN, fill: { color: TINT } });
  const rows = [
    [hdr("Qualitätsattribut"), hdr("Konkrete Anforderung"), hdr("Prüfbarkeit")],
    [b("Sicherheit"), c("Das System muss die Themenfreigabe und die Antragsdaten vor Zugriff durch nicht berechtigte Personen schützen."), c("Aufruf und Freigabeversuch mit einem nicht zugeordneten Konto werden vom System verhindert.")],
    [b("Zuverlässigkeit"), c("Eine erteilte Themenfreigabe muss vollständig und unverändert gespeichert bleiben."), c("Nach erneutem Aufruf sind Thema und Zeitpunkt der Themenfreigabe identisch vorhanden.")],
    [b("Benutzbarkeit"), c("Das System muss der Erstprüfer:in die Möglichkeit bieten, die Themenfreigabe ohne Schulung am Desktop und am Smartphone durchzuführen."), c("Nutzungstest mit Erstprüfenden auf beiden Gerätetypen. Zielwert noch zu validieren.")],
    [b("Effizienz"), c("Die Themenfreigabe sollte für die Erstprüfer:in mit weniger Arbeitsschritten möglich sein als im heutigen Formularprozess."), c("Vergleich der Arbeitsschritte heute und im Portal. Zielwert noch zu validieren.")],
  ];
  s.addTable(rows, { x: 0.45, y: 1.08, w: 9.1, colW: [1.5, 4.2, 3.4], rowH: [0.32, 0.74, 0.66, 0.82, 0.74], fontFace: "Calibri", border: { type: "solid", color: LGREY, pt: 0.75 }, valign: "middle", margin: [3, 6, 3, 6] });
  rbox(s, 0.45, 4.55, 9.1, 0.57, AMBERPALE, AMBER);
  T(s, [
    { text: "Abgrenzung: ", options: { bold: true, color: BROWN } },
    { text: "Qualitätsanforderungen sind getrennt von den funktionalen Regeln C1 bis C6 dokumentiert. Messwerte sind im Projekt nicht vorgegeben und deshalb als „noch zu validieren“ gekennzeichnet." },
  ], { x: 0.6, y: 4.55, w: 8.8, h: 0.57, fontSize: 9, valign: "middle" });
  s.addNotes("Vier Qualitätsattribute aus den Kategorien der Vorlesung (Sicherheit, Zuverlässigkeit, Benutzbarkeit, Effizienz). Bezug zum Projektbriefing: komfortabel nutzbar auf Desktop und Smartphone; Prozess für alle Beteiligten effizienter abwickeln. Es werden keine Grenzwerte erfunden.");
}

// =====================================================================
// Folie 9: Konsistenz und zu validierende Annahmen
// =====================================================================
{
  const s = slide(0, "Konsistenz und offene Punkte", "Begriffe wie im Story Mapping und im Zulassungsantrag");
  T(s, "Ein Begriff, eine Bedeutung", { x: 0.45, y: 1.08, w: 4.5, h: 0.3, fontSize: 11.5, bold: true, color: GREEN, valign: "middle" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: GREEN }, fontSize: 8.5 } });
  const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 8.5, color: INK }, o || {}) });
  const b = (t) => c(t, { bold: true, color: GREEN, fill: { color: TINT } });
  s.addTable([
    [hdr("Begriff"), hdr("User"), hdr("Action"), hdr("Data"), hdr("Control")],
    [b("Erstprüfer:in"), c("Persona B"), c("Lane"), c("Klasse"), c("C1 bis C5")],
    [b("Antrag"), c("Kontext"), c("aufrufen"), c("Klasse"), c("C1, C2, C6")],
    [b("Thema"), c("Ziel"), c("prüfen"), c("Klasse"), c("C2 bis C5")],
    [b("Themenfreigabe"), c("Aufgaben"), c("speichern"), c("Klasse"), c("C4, C6")],
    [b("Studierende:r"), c("Persona A"), c("Endereignis"), c("Klasse"), c("–")],
  ], { x: 0.45, y: 1.45, w: 4.55, colW: [1.25, 0.85, 0.85, 0.7, 0.9], rowH: 0.34, fontFace: "Calibri", border: { type: "solid", color: LGREY, pt: 0.75 }, valign: "middle", margin: [2, 5, 2, 5] });

  rbox(s, 5.2, 1.08, 4.35, 4.04, AMBERPALE, AMBER);
  T(s, "Zu validierende Annahmen", { x: 5.35, y: 1.15, w: 4.05, h: 0.3, fontSize: 11.5, bold: true, color: BROWN, valign: "middle" });
  const items = [
    "Zeitpunkt der Themenfreigabe: Story Mapping und Zulassungsantrag legen sie nach der Entscheidung des Prüfungsausschusses, das Aktivitätsdiagramm davor.",
    "C6: Nach der Themenfreigabe folgt das Festlegen der Bearbeitungszeit.",
    "Wertebereich des Antragsstatus und Multiplizität Studierende:r zu Antrag.",
    "Weg der Benachrichtigung der Erstprüfer:in.",
    "Persona B ist fiktiv, Zielwerte der Qualitätsanforderungen fehlen.",
  ];
  T(s, items.map((t, i) => ({ text: t, options: { bullet: { indent: 11 }, breakLine: i < items.length - 1 } })), { x: 5.37, y: 1.52, w: 4.0, h: 3.5, fontSize: 9.5, paraSpaceAfter: 6 });
  T(s, "Interface: wird durch das Projektteam ergänzt und anschließend auf dieselben Begriffe geprüft.", { x: 0.45, y: 3.65, w: 4.55, h: 0.5, fontSize: 9, italic: true, color: MUTED });
  s.addNotes("Abschlussfolie: zeigt, dass alle Dimensionen dieselben Begriffe verwenden, und sammelt die Annahmen, die mit dem Auftraggeber zu validieren sind.");
}

const out = process.env.OUT || "7_Dimensions_Erstpruefer_Themenfreigabe.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
