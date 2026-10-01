// D2D-Beschreibung (Discover to Deliver, 7 Product Dimensions) fuer User Story 4.1.1
// "Als Erstpruefer:in moechte ich das Thema der Abschlussarbeit bestaetigen, damit die Arbeit
//  mit dem abgestimmten Thema offiziell freigegeben wird." (Story Mapping V2, Activity 4, User Task 4.1)
// Aufruf: OUT="../D2D User Story/D2D_US_4.1.1_Thema_bestaetigen.pptx" node d2d_us411.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "D2D User Story 4.1.1";

const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF",
  PALE = "D9E6CF", GREY = "8A968A", LGREY = "C9D1C9", BG = "F5F7F4", BLUE = "4A7FB5", BLUEPALE = "DCE8F5",
  AMBER = "C9A227", AMBERPALE = "FFF7E0";
const DIMS = ["User", "Interface", "Action", "Data", "Control", "Environment", "Quality Attribute"];
const STORY = "Als Erstprüfer:in möchte ich das Thema der Abschlussarbeit bestätigen, damit die Arbeit mit dem abgestimmten Thema offiziell freigegeben wird.";
const PHOTO = path.join(process.env.PHOTO_DIR || path.join(__dirname, "photos"), "bachelor.png");

// ---------- Helfer ----------
const T = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: "Calibri", color: INK, valign: "top" }, o));
const box = (s, x, y, w, h, fill, line, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 1 }, rectRadius: 0.08 }, o || {}));
const rect = (s, x, y, w, h, fill, line) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 0.75 } });
const circle = (s, x, y, d, fill, line) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: line || fill, width: 1 } });
const num = (s, x, y, n, d) => { d = d || 0.28; circle(s, x, y, d, GREEN); T(s, String(n), { x, y, w: d, h: d, fontSize: d > 0.3 ? 12 : 10, bold: true, color: WHITE, align: "center", valign: "middle" }); };
// Linie mit optionaler Pfeilspitze am Ende (x2,y2); nur waagerecht/senkrecht
const seg = (s, x1, y1, x2, y2, head, color, dash) => {
  const horiz = Math.abs(y2 - y1) < 0.001;
  const x = Math.min(x1, x2), y = Math.min(y1, y2);
  const line = { color: color || MUTED, width: 1.25 };
  if (dash) line.dashType = dash;
  if (head) { const end = horiz ? x2 > x1 : y2 > y1; if (end) line.endArrowType = "triangle"; else line.beginArrowType = "triangle"; }
  s.addShape(pres.shapes.LINE, { x, y, w: horiz ? Math.abs(x2 - x1) : 0.001, h: horiz ? 0.001 : Math.abs(y2 - y1), line });
};
const bullets = (s, items, o) => T(s, items.map((t, i) => ({ text: t, options: { bullet: { indent: 11 }, breakLine: i < items.length - 1 } })), Object.assign({ fontSize: 9.5, paraSpaceAfter: 3 }, o));
const header = (title, dimIdx, sub) => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  T(s, title, { x: 0.45, y: 0.25, w: 6.6, h: 0.5, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  const tag = dimIdx ? `D2D  ·  DIMENSION ${dimIdx} VON 7  ·  ${DIMS[dimIdx - 1].toUpperCase()}` : "D2D  ·  DISCOVER TO DELIVER";
  T(s, tag, { x: 5.6, y: 0.32, w: 3.95, h: 0.3, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });
  if (sub) T(s, sub, { x: 0.45, y: 0.72, w: 9.1, h: 0.22, fontSize: 9.5, color: MUTED });
  T(s, "US 4.1.1 Thema bestätigen  ·  ThesisFlow  ·  Requirements Engineering, Gruppe Grün", { x: 0.45, y: 5.28, w: 7, h: 0.2, fontSize: 7.5, color: GREY });
  return s;
};

// =====================================================================
// Folie 1: Uebersicht
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 3.7, 5.625, GREEN);
  T(s, "D2D  ·  7 PRODUCT DIMENSIONS", { x: 0.45, y: 0.4, w: 3.0, h: 0.3, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, "User Story 4.1.1", { x: 0.45, y: 0.75, w: 3.0, h: 0.5, fontFace: "Cambria", fontSize: 26, bold: true, color: WHITE, valign: "middle" });
  T(s, "Thema bestätigen", { x: 0.45, y: 1.25, w: 3.0, h: 0.4, fontFace: "Cambria", fontSize: 18, italic: true, color: MOSS, valign: "middle" });
  T(s, "USER STORY", { x: 0.45, y: 2.0, w: 3.0, h: 0.22, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, "„" + STORY + "“", { x: 0.45, y: 2.25, w: 2.95, h: 1.25, fontSize: 11.5, color: WHITE, lineSpacingMultiple: 1.12 });
  T(s, "EINORDNUNG IM STORY MAPPING", { x: 0.45, y: 3.7, w: 3.0, h: 0.22, fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5 });
  T(s, [
    { text: "Activity 4: ", options: { bold: true, color: PALE } }, { text: "Ausgabe der Arbeit", options: { color: WHITE, breakLine: true } },
    { text: "User Task 4.1: ", options: { bold: true, color: PALE } }, { text: "Thema bestätigen und freigeben", options: { color: WHITE, breakLine: true } },
    { text: "Rolle: ", options: { bold: true, color: PALE } }, { text: "Erstprüfer:in", options: { color: WHITE } },
  ], { x: 0.45, y: 3.95, w: 3.0, h: 0.75, fontSize: 9.5, paraSpaceAfter: 3 });
  T(s, "ThesisFlow  ·  Requirements Engineering, Gruppe Grün", { x: 0.45, y: 5.05, w: 3.0, h: 0.25, fontSize: 8, color: PALE, valign: "bottom" });

  T(s, "Die sieben Dimensionen", { x: 4.0, y: 0.38, w: 5.5, h: 0.45, fontFace: "Cambria", fontSize: 22, bold: true, valign: "middle" });
  const rows = [
    ["User", "Personas für Erstprüfer:in und Studierende:r"],
    ["Interface", "Dialog „Thema bestätigen“, Dialoglandkarte, Schnittstellen"],
    ["Action", "Aktivitätsdiagramm: erhalten, lesen, prüfen, bestätigen, freigeben"],
    ["Data", "Ausschnitt des Datenmodells und Datenzugriffe"],
    ["Control", "Geschäftsregeln aus Zulassungsantrag und Prüfungsordnung"],
    ["Environment", "Einsatzumgebung: Geräte, Zugang, Organisation"],
    ["Quality Attribute", "Messbare Qualitätsziele für die Story"],
  ];
  rows.forEach((r, i) => {
    const y = 0.98 + i * 0.58;
    box(s, 4.0, y, 5.5, 0.5, TINT);
    num(s, 4.12, y + 0.11, i + 1);
    T(s, r[0], { x: 4.52, y, w: 1.45, h: 0.5, fontSize: 11, bold: true, color: GREEN, valign: "middle" });
    T(s, r[1], { x: 5.95, y, w: 3.45, h: 0.5, fontSize: 9.5, valign: "middle" });
  });
  T(s, "Quellen: Story Mapping V2, Zulassungsantrag WS23/24, Praxisteil S. 18 bis 33", { x: 4.0, y: 5.15, w: 5.5, h: 0.22, fontSize: 7.5, color: GREY });
}

// =====================================================================
// Folien 2 und 3: User (Personas)
// =====================================================================
const persona = (p) => {
  const s = header(p.title, 1);
  const LX = 0.45, LY = 0.95, LW = 2.75, LH = 4.2;
  box(s, LX, LY, LW, LH, TINT, TINT, { rectRadius: 0.12 });
  if (p.photo && fs.existsSync(p.photo)) {
    const h = 1.55, w = h * 360 / 836;
    s.addImage({ path: p.photo, x: LX + (LW - w) / 2, y: LY + 0.15, w, h });
  } else {
    const d = 1.3;
    circle(s, LX + (LW - d) / 2, LY + 0.25, d, GREEN);
    T(s, p.initials, { x: LX + (LW - d) / 2, y: LY + 0.25, w: d, h: d, fontFace: "Cambria", fontSize: 30, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  T(s, p.name, { x: LX + 0.1, y: LY + 1.78, w: LW - 0.2, h: 0.35, fontFace: "Cambria", fontSize: 15, bold: true, align: "center", valign: "middle" });
  T(s, p.role, { x: LX + 0.15, y: LY + 2.13, w: LW - 0.3, h: 0.36, fontSize: 9, color: MUTED, align: "center" });
  T(s, p.quote, { x: LX + 0.2, y: LY + 2.52, w: LW - 0.4, h: 0.62, fontFace: "Cambria", fontSize: 10, italic: true, color: GREEN, align: "center" });
  T(s, p.facts.map((f, i) => ({ text: f, options: { breakLine: i < p.facts.length - 1 } })), { x: LX + 0.2, y: LY + 3.2, w: LW - 0.4, h: 0.95, fontSize: 8.5, paraSpaceAfter: 2 });

  const RX = 3.5, RW = 2.95, GAPX = 0.15, RY = 0.95, RH = 1.62, GAPY = 0.08;
  p.sections.forEach((sec, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = RX + col * (RW + GAPX), y = RY + row * (RH + GAPY);
    num(s, x, y + 0.03, i + 1, 0.26);
    T(s, sec[0], { x: x + 0.36, y, w: RW - 0.36, h: 0.32, fontSize: 12, bold: true, color: GREEN, valign: "middle" });
    bullets(s, sec[1], { x: x + 0.05, y: y + 0.38, w: RW - 0.05, h: RH - 0.4, fontSize: 9.5 });
  });
  // Bezug zur Story
  box(s, RX, 4.42, 6.05, 0.73, AMBERPALE, AMBER);
  T(s, [{ text: "Bezug zur User Story 4.1.1:  ", options: { bold: true, color: "7A5C00" } }, { text: p.link }], { x: RX + 0.12, y: 4.42, w: 5.8, h: 0.73, fontSize: 9.5, valign: "middle" });
};

persona({
  title: "User: Persona Erstprüfer:in", initials: "AB", name: "Prof. Dr. Andreas Becker",
  role: "Professor für Wirtschaftsinformatik, FB MND, Erstprüfer",
  quote: "„Ich will ein Thema freigeben, ohne ein Formular dreimal anzufassen.“",
  facts: ["Alter: 51 Jahre", "An der THM seit: 12 Jahren", "Betreut: rund 10 Abschlussarbeiten pro Semester", "Geräte: Laptop im Büro, Smartphone unterwegs", "Hinweis: fiktive Persona, Werte sind Annahmen"],
  sections: [
    ["Ziele", ["Thema auf einen Blick prüfen und verbindlich freigeben", "Überblick über alle betreuten Arbeiten und offenen Aufgaben", "Mehr Zeit für Betreuung statt für Formalien"]],
    ["Eigenschaften und Verhalten", ["Fachlich genau, wenig Zeit zwischen Lehre und Forschung", "Arbeitet viel per E-Mail, erwartet klare Aufgabenlisten", "Skeptisch bei Zusatzaufwand, offen, wenn es Zeit spart"]],
    ["Frustrationen", ["Formular per E-Mail hin- und herschicken", "Unteren Formularteil ausfüllen, unterschreiben, scannen", "Unklar, bei wem ein Antrag gerade liegt"]],
    ["Nutzungsüberwindung", ["Benachrichtigung mit direktem Link zum Antrag", "Bestätigung in wenigen Klicks, auch mobil", "Login mit THM-Konto, kein neues Passwort"]],
  ],
  link: "Hauptnutzer der Story. Er prüft das Thema und gibt die Arbeit frei. Ohne seine Bestätigung geht der Vorgang nicht weiter.",
});
persona({
  title: "User: Persona Studierende:r", photo: PHOTO, name: "Moritz Hoffmann",
  role: "Student B.Sc. Wirtschaftsinformatik, 6. Semester",
  quote: "„Ich will einfach wissen, ob mein Antrag angekommen ist und was noch fehlt.“",
  facts: ["Alter: 23 Jahre", "Wohnort: Gießen, WG", "Nebenjob: Werkstudent, 12 h/Woche", "Geräte: Smartphone, Laptop zu Hause", "Persona aus Aufgabe „Personas“ übernommen"],
  sections: [
    ["Ziele", ["Antrag in einem Durchgang korrekt einreichen", "Bearbeitungsstand jederzeit am Handy sehen", "Bachelor in Regelstudienzeit abschließen"]],
    ["Eigenschaften und Verhalten", ["Erste Abschlussarbeit, unsicher bei Formalien", "Digital-native, erwartet App-Komfort", "Plant knapp, reagiert auf Erinnerungen"]],
    ["Frustrationen", ["Wochenlang keine Rückmeldung zum Antrag", "Unsicher, ob alles vollständig ist", "Sekretariat nur zu Öffnungszeiten erreichbar"]],
    ["Nutzungsüberwindung", ["Benachrichtigung bei jedem Statuswechsel", "Login mit THM-Konto, keine neue Registrierung", "Mobil genauso komfortabel wie am Desktop"]],
  ],
  link: "Mittelbar betroffen. Er hat das Thema vorgeschlagen und wartet auf die Freigabe, um Thema und Frist verbindlich zu kennen.",
});

// =====================================================================
// Folie 4: Interface
// =====================================================================
{
  const s = header("Interface: Dialog „Thema bestätigen“", 2, "GUI-Entwurf für Desktop und Smartphone, Dialoglandkarte und Schnittstellen zu Nachbarsystemen");
  // ---- Browser-Mockup ----
  const BX = 0.45, BY = 1.05, BW = 5.0, BH = 4.05;
  box(s, BX, BY, BW, BH, WHITE, LGREY, { rectRadius: 0.06 });
  rect(s, BX + 0.01, BY + 0.01, BW - 0.02, 0.26, "E9EEE9");
  [0, 1, 2].forEach(i => circle(s, BX + 0.1 + i * 0.14, BY + 0.09, 0.09, LGREY));
  box(s, BX + 0.6, BY + 0.05, 3.2, 0.17, WHITE, LGREY, { rectRadius: 0.04 });
  T(s, "thesisflow.thm.de/antraege/2026-0142/thema", { x: BX + 0.68, y: BY + 0.05, w: 3.1, h: 0.17, fontSize: 6.5, color: MUTED, valign: "middle" });
  rect(s, BX + 0.01, BY + 0.27, BW - 0.02, 0.3, GREEN);
  T(s, "ThesisFlow", { x: BX + 0.15, y: BY + 0.27, w: 1.5, h: 0.3, fontFace: "Cambria", fontSize: 10.5, bold: true, color: WHITE, valign: "middle" });
  T(s, "Prof. Dr. A. Becker  ·  Abmelden", { x: BX + 2.9, y: BY + 0.27, w: 1.95, h: 0.3, fontSize: 7, color: PALE, align: "right", valign: "middle" });
  T(s, "Meine Aufgaben  ›  Antrag Moritz Hoffmann  ›  Thema bestätigen", { x: BX + 0.15, y: BY + 0.63, w: 4.6, h: 0.18, fontSize: 7, color: MUTED, valign: "middle" });
  // Stepper
  const steps = ["Antrag", "Prüfung", "Zulassung", "Thema bestätigen", "Bearbeitungszeit"];
  const sx0 = BX + 0.45, sgap = 0.98, sy = BY + 0.98;
  seg(s, sx0 + 0.08, sy + 0.08, sx0 + 4 * sgap + 0.08, sy + 0.08, false, LGREY);
  steps.forEach((st, i) => {
    const done = i < 3, act = i === 3;
    circle(s, sx0 + i * sgap, sy, 0.16, done ? GREEN : act ? AMBER : WHITE, done ? GREEN : act ? AMBER : LGREY);
    T(s, st, { x: sx0 + i * sgap - 0.42, y: sy + 0.19, w: 1.0, h: 0.18, fontSize: 6.5, bold: act, color: act ? "7A5C00" : MUTED, align: "center", valign: "middle" });
  });
  // Detailkarte
  const CX = BX + 0.2, CY = BY + 1.48, CW = BW - 0.4, CH = 1.62;
  box(s, CX, CY, CW, CH, BG, LGREY, { rectRadius: 0.05 });
  T(s, "Antrag auf Zulassung und Ausgabe der Abschlussarbeit", { x: CX + 0.12, y: CY + 0.06, w: CW - 0.24, h: 0.22, fontSize: 8.5, bold: true, color: GREEN, valign: "middle" });
  const rowsD = [
    ["Studierende:r", "Moritz Hoffmann, B.Sc. Wirtschaftsinformatik"],
    ["Thema", "Digitale Anmeldung von Abschlussarbeiten am Fachbereich MND"],
    ["Zweitprüfer:in", "Dr. Jana Weiß, extern, Qualifikation geprüft"],
    ["Entscheidung PA", "zugelassen am 28.09.2026"],
    ["Wunschstarttermin", "15.10.2026"],
  ];
  rowsD.forEach((r, i) => {
    const y = CY + 0.33 + i * 0.25;
    T(s, r[0], { x: CX + 0.12, y, w: 1.05, h: 0.22, fontSize: 7.5, color: MUTED, valign: "middle" });
    T(s, r[1], { x: CX + 1.2, y, w: CW - 1.32, h: 0.22, fontSize: 7.5, bold: i === 1, valign: "middle" });
  });
  // Buttons
  const by = BY + 3.25;
  box(s, CX, by, 2.15, 0.32, GREEN, GREEN, { rectRadius: 0.05 });
  T(s, "Thema bestätigen und freigeben", { x: CX, y: by, w: 2.15, h: 0.32, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, CX + 2.3, by, 1.15, 0.32, WHITE, GREEN, { rectRadius: 0.05 });
  T(s, "Thema anpassen", { x: CX + 2.3, y: by, w: 1.15, h: 0.32, fontSize: 8, bold: true, color: GREEN, align: "center", valign: "middle" });
  T(s, "Nach der Bestätigung ist das Thema gesperrt. Studierende:r und Sekretariat werden informiert.", { x: CX, y: by + 0.38, w: CW, h: 0.2, fontSize: 6.5, italic: true, color: MUTED, valign: "middle" });

  // ---- Smartphone-Mockup ----
  const PX = 5.7, PY = 1.05, PW = 1.5, PH = 4.05;
  box(s, PX, PY, PW, PH, "2B2B2B", "2B2B2B", { rectRadius: 0.16 });
  rect(s, PX + 0.07, PY + 0.22, PW - 0.14, PH - 0.44, WHITE);
  rect(s, PX + 0.07, PY + 0.22, PW - 0.14, 0.28, GREEN);
  T(s, "ThesisFlow", { x: PX + 0.14, y: PY + 0.22, w: 1.2, h: 0.28, fontFace: "Cambria", fontSize: 8.5, bold: true, color: WHITE, valign: "middle" });
  T(s, "Thema bestätigen", { x: PX + 0.14, y: PY + 0.58, w: PW - 0.28, h: 0.22, fontSize: 8.5, bold: true, color: GREEN, valign: "middle" });
  box(s, PX + 0.13, PY + 0.86, PW - 0.26, 1.72, BG, LGREY, { rectRadius: 0.04 });
  T(s, [
    { text: "Studierende:r", options: { color: MUTED, fontSize: 6, breakLine: true } },
    { text: "Moritz Hoffmann", options: { fontSize: 7, breakLine: true } },
    { text: "Thema", options: { color: MUTED, fontSize: 6, breakLine: true } },
    { text: "Digitale Anmeldung von Abschlussarbeiten am FB MND", options: { fontSize: 7, bold: true, breakLine: true } },
    { text: "Entscheidung PA", options: { color: MUTED, fontSize: 6, breakLine: true } },
    { text: "zugelassen am 28.09.2026", options: { fontSize: 7 } },
  ], { x: PX + 0.19, y: PY + 0.9, w: PW - 0.38, h: 1.64, paraSpaceAfter: 2, valign: "top" });
  box(s, PX + 0.13, PY + 2.7, PW - 0.26, 0.3, GREEN, GREEN, { rectRadius: 0.05 });
  T(s, "Bestätigen und freigeben", { x: PX + 0.13, y: PY + 2.7, w: PW - 0.26, h: 0.3, fontSize: 7, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, PX + 0.13, PY + 3.07, PW - 0.26, 0.3, WHITE, GREEN, { rectRadius: 0.05 });
  T(s, "Thema anpassen", { x: PX + 0.13, y: PY + 3.07, w: PW - 0.26, h: 0.3, fontSize: 7, bold: true, color: GREEN, align: "center", valign: "middle" });

  // ---- Dialoglandkarte ----
  const DX = 7.45, DW = 2.1;
  T(s, "Dialoglandkarte", { x: DX, y: 1.05, w: DW, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  const dlg = ["Meine Aufgaben", "Antragsdetail", "Thema bestätigen", "Bearbeitungszeit festlegen"];
  dlg.forEach((d, i) => {
    const y = 1.36 + i * 0.44, act = i === 2;
    box(s, DX, y, DW, 0.3, act ? AMBERPALE : TINT, act ? AMBER : TINT, { rectRadius: 0.05 });
    T(s, d, { x: DX, y, w: DW, h: 0.3, fontSize: 8.5, bold: act, align: "center", valign: "middle" });
    if (i < dlg.length - 1) seg(s, DX + DW / 2, y + 0.3, DX + DW / 2, y + 0.44, true);
  });
  // ---- Schnittstellen ----
  T(s, "Schnittstellen", { x: DX, y: 3.22, w: DW, h: 0.25, fontSize: 10.5, bold: true, color: GREEN, valign: "middle" });
  const ifs = [["THM-Login (SSO)", "Anmeldung, Identität und Rolle der Erstprüfer:in"], ["E-Mail-Dienst THM", "Benachrichtigung mit Link, danach Info an Studierende:n und Sekretariat"]];
  ifs.forEach((f, i) => {
    const y = 3.52 + i * 0.8;
    box(s, DX, y, DW, 0.72, WHITE, GREY, { rectRadius: 0.05, line: { color: GREY, width: 1, dashType: "dash" } });
    T(s, [{ text: f[0], options: { bold: true, fontSize: 8.5, color: MUTED, breakLine: true } }, { text: f[1], options: { fontSize: 7.5 } }], { x: DX + 0.08, y, w: DW - 0.16, h: 0.72, valign: "middle" });
  });
}

// =====================================================================
// Folie 5: Action (Aktivitaetsdiagramm mit Swimlanes)
// =====================================================================
{
  const s = header("Action: Aktivitätsdiagramm", 3, "Ablauf der Story: Antrag erhalten, lesen, Thema prüfen, bestätigen, freigeben");
  const LX = 0.45, LW = 9.1, LAB = 0.95;
  const Y1 = 1.05, H1 = 2.0, Y2 = Y1 + H1, H2 = 1.55;
  rect(s, LX, Y1, LW, H1, WHITE, LGREY); rect(s, LX, Y2, LW, H2, BG, LGREY);
  rect(s, LX, Y1, LAB, H1, TINT, LGREY); rect(s, LX, Y2, LAB, H2, TINT, LGREY);
  T(s, "Erstprüfer:in", { x: LX, y: Y1, w: LAB, h: H1, fontSize: 9.5, bold: true, color: GREEN, align: "center", valign: "middle" });
  T(s, "ThesisFlow (System)", { x: LX + 0.05, y: Y2, w: LAB - 0.1, h: H2, fontSize: 9.5, bold: true, color: GREEN, align: "center", valign: "middle" });

  const AW = 1.15, AH = 0.5;
  const act = (x, cy, text, sys) => {
    box(s, x, cy - AH / 2, AW, AH, sys ? BLUEPALE : TINT, sys ? BLUE : GREEN, { rectRadius: 0.12 });
    T(s, text, { x: x + 0.04, y: cy - AH / 2, w: AW - 0.08, h: AH, fontSize: 8, align: "center", valign: "middle" });
  };
  const cyE = 2.45, cyAlt = 1.5, cyS = 3.85;
  const xS = 1.75, x2 = 3.1, x3 = 4.45, xD = 6.0, x4 = 6.55, x5 = 7.95;
  // Start
  circle(s, 1.5, cyS - 0.09, 0.18, INK);
  seg(s, 1.68, cyS, xS, cyS, true);
  act(xS, cyS, "Erstprüfer:in über Zulassung benachrichtigen", true);
  seg(s, xS + AW / 2, cyS - AH / 2, xS + AW / 2, cyE + AH / 2, true);
  act(xS, cyE, "Antrag erhalten und öffnen");
  seg(s, xS + AW, cyE, x2, cyE, true);
  act(x2, cyE, "Antrag lesen");
  seg(s, x2 + AW, cyE, x3, cyE, true);
  act(x3, cyE, "Thema prüfen");
  seg(s, x3 + AW, cyE, xD - 0.2, cyE, true);
  // Entscheidung
  s.addShape(pres.shapes.DIAMOND, { x: xD - 0.2, y: cyE - 0.2, w: 0.4, h: 0.4, fill: { color: AMBERPALE }, line: { color: AMBER, width: 1.25 } });
  T(s, "Thema entspricht der Absprache?", { x: xD - 0.75, y: cyE + 0.24, w: 1.5, h: 0.3, fontSize: 7, italic: true, color: MUTED, align: "center" });
  seg(s, xD + 0.2, cyE, x4, cyE, true);
  T(s, "[ja]", { x: xD + 0.2, y: cyE - 0.2, w: 0.35, h: 0.18, fontSize: 7, color: MUTED, align: "center" });
  // nein-Zweig
  seg(s, xD, cyE - 0.2, xD, cyAlt, false);
  seg(s, xD, cyAlt, x4, cyAlt, true);
  T(s, "[nein]", { x: xD + 0.04, y: cyAlt + 0.03, w: 0.45, h: 0.18, fontSize: 7, color: MUTED });
  act(x4, cyAlt, "Thema anpassen");
  seg(s, x4 + AW / 2, cyAlt + AH / 2, x4 + AW / 2, cyE - AH / 2, true);
  act(x4, cyE, "Thema bestätigen");
  seg(s, x4 + AW / 2, cyE + AH / 2, x4 + AW / 2, cyS - AH / 2, true);
  act(x4, cyS, "Bestätigung mit Zeitstempel speichern", true);
  seg(s, x4 + AW, cyS, x5, cyS, true);
  act(x5, cyS, "Arbeit freigeben, Beteiligte informieren", true);
  seg(s, x5 + AW, cyS, 9.27, cyS, true);
  // Ende
  circle(s, 9.27, cyS - 0.11, 0.22, WHITE, INK); circle(s, 9.32, cyS - 0.06, 0.12, INK);

  // Erlaeuterung
  box(s, 0.45, 4.7, 9.1, 0.48, AMBERPALE, AMBER);
  T(s, [
    { text: "Auslöser: ", options: { bold: true, color: "7A5C00" } }, { text: "Zulassung durch den Prüfungsausschuss-Vorsitz (US 3.2.1).   " },
    { text: "Ergebnis: ", options: { bold: true, color: "7A5C00" } }, { text: "Status „Thema bestätigt“.   " },
    { text: "Folgeschritt: ", options: { bold: true, color: "7A5C00" } }, { text: "Bearbeitungszeit festlegen (User Task 4.2)." },
  ], { x: 0.57, y: 4.7, w: 8.86, h: 0.48, fontSize: 8.5, valign: "middle" });
}

// =====================================================================
// Folie 6: Data
// =====================================================================
{
  const s = header("Data: Datenausschnitt und Zugriffe", 4, "Für die Story relevanter Ausschnitt des Datenmodells und was mit den Daten geschieht");
  const EW = 1.55, EH = 0.78;
  const ent = (x, y, name, attrs, isNew) => {
    box(s, x, y, EW, EH, isNew ? AMBERPALE : TINT, isNew ? AMBER : GREEN, { rectRadius: 0.05 });
    T(s, [{ text: name, options: { bold: true, fontSize: 9, color: isNew ? "7A5C00" : GREEN, breakLine: true } }, { text: attrs, options: { fontSize: 7, color: INK } }], { x: x + 0.06, y, w: EW - 0.12, h: EH, align: "center", valign: "middle" });
  };
  const ax = 2.35, ay = 2.2;
  // Linien zuerst
  seg(s, ax + EW / 2, 1.83, ax + EW / 2, ay, false);
  seg(s, 2.0, ay + EH / 2, ax, ay + EH / 2, false);
  seg(s, ax + EW, ay + EH / 2, 4.25, ay + EH / 2, false);
  seg(s, 2.7, ay + EH, 2.7, 3.4, false);
  seg(s, 3.6, ay + EH, 3.6, 3.4, false);
  ent(ax, 1.05, "Erstprüfer:in", "Name, THM-Kennung, Rolle");
  ent(0.45, ay, "Studierende:r", "Name, Matrikelnummer, Studiengang");
  ent(ax, ay, "Antrag", "Antragsnummer, Status, Wunschstarttermin");
  ent(4.25, ay, "Thema", "Wortlaut, Stand (vorgeschlagen, bestätigt)");
  ent(1.4, 3.4, "Zulassungsentscheidung", "Ergebnis, Datum, PA-Vorsitz");
  ent(3.3, 3.4, "Themenbestätigung", "bestätigt durch, Zeitstempel", true);
  T(s, "prüft", { x: ax + EW / 2 + 0.05, y: 1.9, w: 0.6, h: 0.18, fontSize: 7, italic: true, color: MUTED });
  T(s, "stellt", { x: 2.0, y: ay + EH / 2 - 0.2, w: 0.35, h: 0.18, fontSize: 7, italic: true, color: MUTED, align: "center" });
  T(s, "hat", { x: ax + EW, y: ay + EH / 2 - 0.2, w: 0.35, h: 0.18, fontSize: 7, italic: true, color: MUTED, align: "center" });
  T(s, "Gelb: wird in dieser Story neu angelegt.", { x: 0.45, y: 4.3, w: 5.3, h: 0.2, fontSize: 8, italic: true, color: MUTED });
  T(s, "Alle Daten liegen im Portal. Nur Identität und Rolle der Erstprüfer:in kommen vom THM-Login. Gelöscht wird in dieser Story nichts.", { x: 0.45, y: 4.55, w: 5.3, h: 0.5, fontSize: 9, color: INK });

  const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: GREEN }, fontSize: 8.5 } });
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 8, color: INK }, o || {}) });
  const rows = [
    [hdr("Daten"), hdr("Zugriff"), hdr("Quelle")],
    [cell("Antrag, Studierende:r, Prüfende"), cell("anzeigen"), cell("Portal")],
    [cell("Zulassungsentscheidung"), cell("anzeigen"), cell("Portal")],
    [cell("Thema (Wortlaut)"), cell("anzeigen, ggf. ändern"), cell("Portal")],
    [cell("Themenbestätigung"), cell("anlegen", { bold: true }), cell("Portal")],
    [cell("Antragsstatus"), cell("ändern"), cell("Portal")],
    [cell("Identität, Rolle"), cell("beziehen"), cell("THM-Login")],
    [cell("Benachrichtigung"), cell("übergeben"), cell("E-Mail-Dienst")],
  ];
  s.addTable(rows, { x: 6.05, y: 1.05, w: 3.5, colW: [1.55, 1.1, 0.85], rowH: 0.33, fontFace: "Calibri", border: { type: "solid", color: LGREY, pt: 0.75 }, valign: "middle", margin: [2, 5, 2, 5] });
}

// =====================================================================
// Folie 7: Control
// =====================================================================
{
  const s = header("Control: Geschäftsregeln", 5, "Regeln und Bedingungen, die das Bestätigen des Themas einschränken");
  const rules = [
    ["Erst nach der Zulassung", "Das Thema kann erst bestätigt werden, wenn der Prüfungsausschuss-Vorsitz den Antrag angenommen hat.", "Zulassungsantrag, Schritte 4 und 5"],
    ["Nur die zuständige Erstprüfer:in", "Bestätigen darf nur die im Antrag eingetragene Erstprüfer:in der THM.", "Zulassungsantrag, Feld „1. Prüfer(in) (THM)“"],
    ["Aktenkundig machen", "Thema und Zeitpunkt der Ausgabe werden mit Person und Zeitstempel dokumentiert.", "Allg. PO der THM § 17 (3), laut Zulassungsantrag"],
    ["Abweichung dokumentieren", "Weicht das Thema von der Absprache ab, wird es vor der Bestätigung angepasst und die Änderung festgehalten.", "Annahme, abgeleitet aus dem Formularfeld „Bestätigung der Vorschläge: Nein“"],
    ["Thema danach gesperrt", "Nach der Bestätigung ist der Wortlaut nicht mehr frei änderbar.", "Annahme, folgt aus der Aktenkundigkeit"],
    ["Rückgabe nur einmal", "Das Thema kann nur einmal und nur innerhalb von vier Wochen nach Ausgabe zurückgegeben werden.", "Allg. PO der THM § 17 (3), laut Zulassungsantrag"],
  ];
  const cw = 2.95, ch = 1.92, gap = 0.125;
  rules.forEach((r, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.45 + col * (cw + gap), y = 1.05 + row * (ch + gap);
    box(s, x, y, cw, ch, TINT);
    num(s, x + 0.12, y + 0.12, "R" + (i + 1), 0.34);
    T(s, r[0], { x: x + 0.55, y: y + 0.1, w: cw - 0.65, h: 0.38, fontSize: 11, bold: true, color: GREEN, valign: "middle" });
    T(s, r[1], { x: x + 0.14, y: y + 0.58, w: cw - 0.28, h: 0.8, fontSize: 9.5, lineSpacingMultiple: 1.08 });
    T(s, "Quelle: " + r[2], { x: x + 0.14, y: y + 1.42, w: cw - 0.28, h: 0.4, fontSize: 7.5, italic: true, color: MUTED, valign: "bottom" });
  });
}

// =====================================================================
// Folie 8: Environment
// =====================================================================
{
  const s = header("Environment: Einsatzumgebung", 6, "Wo und unter welchen Rahmenbedingungen die Story genutzt wird");
  const cards = [
    ["Geräte und Technik", ["Webanwendung im Browser, keine Installation", "Komfortabel nutzbar am Desktop und am Smartphone", "Internetverbindung erforderlich"], "Projektbriefing"],
    ["Zugang", ["Anmeldung mit dem THM-Konto", "Kein eigener Account für Erstprüfende", "Benachrichtigung per THM-E-Mail mit Link"], "THM-Seite „Abschlussarbeit“, Kontextdiagramm"],
    ["Organisation", ["Nutzung im Büro und zwischen Lehrveranstaltungen", "Bestätigung unabhängig von Sprechzeiten des Sekretariats", "Fachbereich MND, Studiengänge B.Sc. und M.Sc. Wirtschaftsinformatik"], "Projektbriefing, Persona"],
  ];
  const cw = 2.95, ch = 2.75, gap = 0.125;
  cards.forEach((c, i) => {
    const x = 0.45 + i * (cw + gap), y = 1.05;
    box(s, x, y, cw, ch, TINT);
    num(s, x + 0.12, y + 0.12, i + 1);
    T(s, c[0], { x: x + 0.5, y: y + 0.1, w: cw - 0.6, h: 0.32, fontSize: 11.5, bold: true, color: GREEN, valign: "middle" });
    bullets(s, c[1], { x: x + 0.16, y: y + 0.6, w: cw - 0.3, h: 1.6, fontSize: 9.5, paraSpaceAfter: 5 });
    T(s, "Quelle: " + c[2], { x: x + 0.14, y: y + ch - 0.4, w: cw - 0.28, h: 0.3, fontSize: 7.5, italic: true, color: MUTED, valign: "bottom" });
  });
  box(s, 0.45, 4.0, 9.1, 1.1, AMBERPALE, AMBER);
  T(s, [
    { text: "Einschätzung: ", options: { bold: true, color: "7A5C00" } },
    { text: "Diese Dimension liefert für die einzelne Story wenig Eigenes. Die Umgebung gilt für das gesamte Portal und ergibt sich aus dem Projektbriefing. Story-spezifisch ist nur, dass die Bestätigung auch mobil und außerhalb der Sprechzeiten möglich sein muss." },
  ], { x: 0.6, y: 4.0, w: 8.8, h: 1.1, fontSize: 10, valign: "middle", lineSpacingMultiple: 1.1 });
}

// =====================================================================
// Folie 9: Quality Attribute
// =====================================================================
{
  const s = header("Quality Attribute: Qualitätsziele", 7, "Messbare Qualitätsanforderungen für das Bestätigen des Themas");
  const q = [
    ["Nachvollziehbarkeit", "Jede Bestätigung ist mit Person und Zeitstempel protokolliert und nachträglich nicht änderbar.", "100 % der Bestätigungen im Protokoll"],
    ["Sicherheit und Datenschutz", "Nur die zugeordnete Erstprüfer:in kann bestätigen. Personenbezogene Daten sind nur für Beteiligte sichtbar.", "0 Bestätigungen durch nicht berechtigte Rollen"],
    ["Bedienbarkeit", "Die Bestätigung gelingt ohne Schulung, am Desktop und am Smartphone.", "höchstens 3 Klicks ab Aufgabenliste (Annahme)"],
    ["Aktualität", "Statusänderung und Benachrichtigung erreichen Studierende:n und Sekretariat ohne Verzögerung.", "Status sofort sichtbar, E-Mail innerhalb von 1 Minute (Annahme)"],
  ];
  const cw = 4.49, ch = 1.92, gap = 0.125;
  q.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.45 + col * (cw + gap), y = 1.05 + row * (ch + gap);
    box(s, x, y, cw, ch, TINT);
    num(s, x + 0.12, y + 0.12, i + 1);
    T(s, r[0], { x: x + 0.5, y: y + 0.1, w: cw - 0.6, h: 0.32, fontSize: 11.5, bold: true, color: GREEN, valign: "middle" });
    T(s, r[1], { x: x + 0.16, y: y + 0.55, w: cw - 0.32, h: 0.65, fontSize: 9.5, lineSpacingMultiple: 1.08 });
    box(s, x + 0.14, y + 1.32, cw - 0.28, 0.45, WHITE, MOSS, { rectRadius: 0.05 });
    T(s, [{ text: "Messgröße: ", options: { bold: true, color: GREEN } }, { text: r[2] }], { x: x + 0.24, y: y + 1.32, w: cw - 0.48, h: 0.45, fontSize: 9, valign: "middle" });
  });
}

// =====================================================================
// Folie 10: Annahmen und offene Punkte
// =====================================================================
{
  const s = header("Annahmen und offene Punkte", 0, "Zur Klärung im Team und im Interview mit dem Auftraggeber");
  box(s, 0.45, 1.05, 4.49, 3.95, TINT);
  T(s, "Getroffene Annahmen", { x: 0.6, y: 1.13, w: 4.2, h: 0.32, fontSize: 12, bold: true, color: GREEN, valign: "middle" });
  bullets(s, [
    "Die Erstprüfer:in darf den Wortlaut des Themas vor der Bestätigung im Portal anpassen.",
    "Nach der Bestätigung ist das Thema gesperrt.",
    "Die Story umfasst nur die Themenbestätigung. Beginn und Abgabe gehören zu User Task 4.2.",
    "Zielwerte für Klicks und Benachrichtigungszeit sind Vorschläge.",
    "Die Persona Erstprüfer:in ist fiktiv, ihre Werte sind nicht erhoben.",
  ], { x: 0.62, y: 1.52, w: 4.15, h: 3.4, fontSize: 10, paraSpaceAfter: 6 });
  box(s, 5.065, 1.05, 4.49, 3.95, AMBERPALE, AMBER);
  T(s, "Offene Fragen", { x: 5.215, y: 1.13, w: 4.2, h: 0.32, fontSize: 12, bold: true, color: "7A5C00", valign: "middle" });
  bullets(s, [
    "Geht ein angepasstes Thema zur Kenntnis oder zur Zustimmung an die Studierenden zurück?",
    "Gilt die Themenbestätigung schon als Ausgabe, oder erst zusammen mit Beginn und Abgabe?",
    "Muss der Prüfungsausschuss ein geändertes Thema erneut sehen?",
    "Welche Zielwerte für Bedienbarkeit und Aktualität akzeptiert der Auftraggeber?",
    "Reicht die digitale Bestätigung rechtlich als Ersatz für die Unterschrift?",
  ], { x: 5.235, y: 1.52, w: 4.15, h: 3.4, fontSize: 10, paraSpaceAfter: 6 });
}

const out = process.env.OUT || "D2D_US_4.1.1_Thema_bestaetigen.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
