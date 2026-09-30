const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "Personas ThesisFlow";
const GREEN = "2C5F2D", MOSS = "97BC62", TINT = "EEF4E8", INK = "1F2A1F", MUTED = "5C6B5C", WHITE = "FFFFFF", PALE = "D9E6CF", GREY = "8A968A";

const personas = [
  {
    n: 1, title: "Persona: Studierende (Bachelor)", name: "Moritz Hoffmann",
    role: "Student B.Sc. Wirtschaftsinformatik, 6. Semester", photo: "photos/bachelor.png", full: true,
    quote: "„Ich will einfach wissen, ob mein Antrag angekommen ist und was noch fehlt.“",
    facts: ["Alter: 23 Jahre", "Wohnort: Gießen, WG", "Familienstand: ledig", "Nebenjob: Werkstudent, 12 h/Woche", "Geräte: Smartphone, Laptop zu Hause"],
    sections: [
      ["Ziele", ["Antrag in einem Durchgang korrekt einreichen", "Bearbeitungsstand jederzeit am Handy sehen", "Prüfende ohne Unterschriftenlauf bestätigen lassen", "Bachelor in Regelstudienzeit abschließen"]],
      ["Eigenschaften und Verhalten", ["Erste Abschlussarbeit, unsicher bei Formalien", "Digital-native, erwartet App-Komfort", "Plant knapp, reagiert auf Erinnerungen", "Volleyball, Serien, wenig Geduld mit Papier"]],
      ["Frustrationen", ["PDF ausdrucken, unterschreiben, scannen, mailen", "Wochenlang keine Rückmeldung zum Antrag", "Unsicher, ob alles vollständig ist", "Sekretariat nur zu Öffnungszeiten erreichbar"]],
      ["Nutzungsüberwindung", ["Schritt-für-Schritt-Führung mit Pflichtfeldprüfung", "Login mit THM-Konto, keine neue Registrierung", "Push- oder Mail-Benachrichtigung bei Statuswechsel", "Mobil genauso komfortabel wie am Desktop"]],
    ],
    scales: [["Technikaffinität", 4], ["Prozesskenntnis", 1], ["Zeitbudget", 3]],
  },
  {
    n: 2, title: "Persona: Sekretariat / Prüfungsamt", name: "Michael Krüger",
    role: "Verwaltungsangestellter, Fachbereich MND", photo: "photos/sekretariat.png",
    quote: "„Jeden unvollständigen Antrag muss ich dreimal in die Hand nehmen, mkay?“",
    facts: ["Alter: 52 Jahre", "Wohnort: Wetzlar", "Familienstand: verheiratet, zwei Kinder", "Am FB MND seit: 14 Jahren", "Geräte: Desktop-PC, Outlook, Excel, Papierakten"],
    sections: [
      ["Ziele", ["Vollständigkeit eines Antrags auf einen Blick prüfen", "Fristen automatisch berechnen und überwachen", "Status ohne Mail- und Telefonrunden weitergeben", "Rechtssichere Ablage ohne Papierstapel"]],
      ["Eigenschaften und Verhalten", ["Sorgfältig, kennt Prüfungsordnungen auswendig", "Erste Anlaufstelle für Studierende und Prüfende", "Skeptisch bei Neuerungen, offen wenn sie entlasten", "Garten, Nordic Walking, Enkelkind"]],
      ["Frustrationen", ["Unleserliche Scans und fehlende Unterschriften", "Fristen aus E-Mails heraussuchen", "Telefonrückfragen binden täglich Stunden", "Sonderfälle nur im eigenen Kopf dokumentiert"]],
      ["Nutzungsüberwindung", ["Läuft zuverlässig und bildet seine Prüfschritte ab", "Exporte und Listen für das Prüfungsamt", "Regeln und Fristen selbst pflegbar", "Kurze Einführung, keine wochenlange Schulung"]],
    ],
    scales: [["Technikaffinität", 2], ["Prozesskenntnis", 5], ["Zeitbudget", 2]],
  },
  {
    n: 3, title: "Persona: Studierende (Master)", name: "Tarek Yilmaz",
    role: "Student M.Sc. Wirtschaftsinformatik, berufsbegleitend", photo: "photos/master.png",
    quote: "„Zwischen Kundenprojekt und Thesis habe ich keine Zeit für Formularpost.“",
    facts: ["Alter: 29 Jahre", "Wohnort: Frankfurt am Main", "Familienstand: verheiratet, ein Kind", "Beruf: IT-Consultant, Vollzeit", "Geräte: Laptop, Firmenhandy, viel unterwegs"],
    sections: [
      ["Ziele", ["Externen Korreferenten aus der Firma direkt im Antrag benennen", "Freigaben asynchron und ohne Präsenztermin", "Verlängerung bei Bedarf online beantragen", "Master neben dem Job abschließen"]],
      ["Eigenschaften und Verhalten", ["Effizient, technikaffin, kennt den Prozess vom Bachelor", "Schreibt die Masterarbeit im eigenen Unternehmen", "Pragmatisch, wenig Geduld für Bürokratie", "Laufen, Fußball mit dem Sohn, Podcasts beim Pendeln"]],
      ["Frustrationen", ["Zusatzformular und Rückfragen für den externen Korreferenten", "Termine im Prüfungsamt kollidieren mit Kundenterminen", "Fristen und Verlängerungsregeln unklar", "Firma hat keinen THM-Zugang"]],
      ["Nutzungsüberwindung", ["Funktioniert asynchron, rund um die Uhr", "Externe können per Einladungslink mitwirken", "Fristen und Verlängerungsoptionen transparent", "Alles in einem Vorgang, keine Doppelerfassung"]],
    ],
    scales: [["Technikaffinität", 5], ["Prozesskenntnis", 3], ["Zeitbudget", 1]],
  },
];

for (const p of personas) {
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText(p.title, { x: 0.45, y: 0.28, w: 6.5, h: 0.5, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 24, bold: true, color: INK, valign: "middle" });
  s.addText("PERSONA " + p.n + " VON 3  ·  THESISFLOW", { x: 6.0, y: 0.35, w: 3.55, h: 0.35, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8, bold: true, color: MOSS, charSpacing: 1.5, align: "right", valign: "middle" });

  // ---- linke Karte ----
  const LX = 0.45, LY = 0.95, LW = 2.75, LH = 4.3;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LX, y: LY, w: LW, h: LH, fill: { color: TINT }, line: { color: TINT }, rectRadius: 0.12 });
  // Foto
  if (p.full) {
    const h = 1.6, w = h * 360 / 836;
    s.addImage({ path: p.photo, x: LX + (LW - w) / 2, y: LY + 0.15, w, h });
  } else {
    const d = 1.35;
    s.addShape(pres.shapes.OVAL, { x: LX + (LW - d) / 2 - 0.05, y: LY + 0.2 - 0.05, w: d + 0.1, h: d + 0.1, fill: { color: GREEN }, line: { color: GREEN } });
    s.addImage({ path: p.photo, x: LX + (LW - d) / 2, y: LY + 0.2, w: d, h: d, rounding: true });
  }
  s.addText(p.name, { x: LX + 0.15, y: LY + 1.85, w: LW - 0.3, h: 0.35, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 16, bold: true, color: INK, align: "center", valign: "middle" });
  s.addText(p.role, { x: LX + 0.15, y: LY + 2.2, w: LW - 0.3, h: 0.35, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9, color: MUTED, align: "center", valign: "top" });
  s.addText(p.quote, { x: LX + 0.2, y: LY + 2.6, w: LW - 0.4, h: 0.6, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 10, italic: true, color: GREEN, align: "center", valign: "top" });
  s.addText(p.facts.map((f, i) => ({ text: f, options: { breakLine: i < p.facts.length - 1 } })), {
    x: LX + 0.2, y: LY + 3.25, w: LW - 0.4, h: 0.95, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 8.5, color: INK, valign: "top", paraSpaceAfter: 2 });

  // ---- rechte Sektionen 2x2 ----
  const RX = 3.5, RW = 2.95, GAPX = 0.15, RY = 0.95, RH = 1.75, GAPY = 0.1;
  p.sections.forEach((sec, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = RX + col * (RW + GAPX), y = RY + row * (RH + GAPY);
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.03, w: 0.26, h: 0.26, fill: { color: GREEN }, line: { color: GREEN } });
    s.addText(String(i + 1), { x, y: y + 0.03, w: 0.26, h: 0.26, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9, bold: true, color: WHITE, align: "center", valign: "middle" });
    s.addText(sec[0], { x: x + 0.36, y, w: RW - 0.36, h: 0.32, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 12.5, bold: true, color: GREEN, valign: "middle" });
    s.addText(sec[1].map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < sec[1].length - 1 } })), {
      x: x + 0.05, y: y + 0.38, w: RW - 0.05, h: RH - 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9.5, color: INK, valign: "top", paraSpaceAfter: 3 });
  });

  // ---- Skalen ----
  const SY = 4.72;
  p.scales.forEach((sc, i) => {
    const x = RX + i * 2.05;
    s.addText(sc[0], { x, y: SY, w: 1.1, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9, bold: true, color: MUTED, valign: "middle" });
    for (let k = 0; k < 5; k++) {
      const on = k < sc[1];
      s.addShape(pres.shapes.OVAL, { x: x + 1.12 + k * 0.19, y: SY + 0.05, w: 0.15, h: 0.15, fill: { color: on ? GREEN : PALE }, line: { color: on ? GREEN : PALE } });
    }
  });
  s.addText("Requirements Engineering – Praxisteil  ·  Personas  ·  Skala 1 = niedrig, 5 = hoch", { x: 3.5, y: 5.15, w: 6.05, h: 0.25, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 7.5, color: GREY, valign: "bottom" });

  s.addNotes(p.name + " – " + p.role + "\n" + p.quote + "\nAufbau nach Vorgabe: Ziele (spezifisch und allgemein), Eigenschaften, Frustrationen, Nutzungsüberwindung.");
}

const out = process.env.OUT || "Personas_Abschlussarbeiten-Portal.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("written", f));
