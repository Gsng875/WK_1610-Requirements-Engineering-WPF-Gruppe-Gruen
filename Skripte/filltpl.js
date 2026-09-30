const JSZip = require("jszip"), fs = require("fs");
const esc = s => s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
(async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync("C:/Users/Ewa/Downloads/OnePager - Template.pptx"));
  let xml = await zip.file("ppt/slides/slide1.xml").async("string");

  const shape = id => {
    const m = xml.match(new RegExp('<p:sp><p:nvSpPr><p:cNvPr id="' + id + '"[^]*?</p:sp>'));
    if (!m) throw new Error("shape " + id + " not found");
    return m[0];
  };
  // Replace all paragraphs of a shape with N paragraphs cloned from its first paragraph (pPr + first run's rPr).
  const setParas = (id, texts, sz) => {
    const sp = shape(id);
    const body = sp.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)[0];
    const firstP = body.match(/<a:p>[\s\S]*?<\/a:p>/)[0];
    const pPr = (firstP.match(/<a:pPr[^>]*\/>|<a:pPr[\s\S]*?<\/a:pPr>/) || [""])[0];
    let rPr = firstP.match(/<a:rPr[^>]*\/>|<a:rPr[\s\S]*?<\/a:rPr>/)[0].replace(/ err="1"/g, "");
    if (sz) rPr = rPr.replace(/sz="\d+"/, 'sz="' + sz + '"');
    const paras = texts.map(t => "<a:p>" + pPr + "<a:r>" + rPr + "<a:t>" + esc(t) + "</a:t></a:r></a:p>").join("");
    const head = body.match(/<p:txBody>[\s\S]*?<a:lstStyle\/>/)[0];
    const newBody = head + paras + "</p:txBody>";
    xml = xml.replace(sp, sp.replace(body, newBody));
  };

  // Titel / Tagline / Vision (id 358): keep the three-run structure
  {
    const sp = shape(358);
    let n = sp
      .replace("<a:t>Projekttitel</a:t>", "<a:t>ThesisFlow – Abschlussarbeiten-Portal MND</a:t>")
      .replace("<a:t>Tagline</a:t>", "<a:t>Anmelden. Betreuen. Abschließen.</a:t>")
      .replace("<a:t> / Vision des Projekts</a:t>", "<a:t> – Jede Abschlussarbeit im Fachbereich MND wird vollständig digital angemeldet, freigegeben und verwaltet: ohne Papier, ohne Medienbrüche, von jedem Gerät aus.</a:t>")
      .replace(/ err="1"/g, "");
    xml = xml.replace(sp, n);
  }
  // Status (id 370)
  {
    const sp = shape(370);
    let n = sp.replace("<a:t>26</a:t>", "<a:t>30</a:t>").replace("<a:t>.06.202</a:t>", "<a:t>.09.202</a:t>").replace("<a:t>5</a:t>", "<a:t>6</a:t>");
    xml = xml.replace(sp, n);
  }
  // Kopfzeile: Labels und Werte
  setParas(359, ["Auftraggeber"]);
  setParas(362, ["Prof. Dr. Carsten Lucke, FB MND"]);
  setParas(360, ["Auftragnehmer"]);
  setParas(363, ["Projektteam RE (Softwarefirma)"]);
  setParas(361, ["Beginn (geschätzt)"]);
  setParas(364, ["Oktober 2026"]);
  setParas(365, ["Ende (geschätzt)"]);
  setParas(366, ["März 2027 (Ausbaustufe 1)"]);
  setParas(367, ["Budget"]);
  setParas(368, ["noch festzulegen"]);

  // Begründung der Projekt-Notwendigkeit (id 372)
  setParas(372, [
    "Anmeldung heute per PDF-Antrag mit Ausdruck, Unterschriften und E-Mail-Umläufen: langsam und fehleranfällig",
    "Kein gemeinsamer Statusüberblick für Studierende, Prüfende und Prüfungsamt",
    "Externe Korreferenten und Verlängerungen laufen über weitere Papierformulare",
    "Der Fachbereich will den Prozess für alle Beteiligten effizienter abwickeln",
  ]);
  // Projektziel (id 356)
  setParas(356, [
    "Webportal für Anmeldung, Freigabe und Verwaltung von Abschlussarbeiten (B.Sc./M.Sc. Wirtschaftsinformatik)",
    "Komfortabel nutzbar im Browser am Desktop und auf dem Smartphone",
    "Externe Korreferenten direkt im Anmeldeprozess einbinden",
    "Regeln der Prüfungsordnungen werden im Antrag geprüft",
    "Ausbaustufe 2: Verlängerungsanträge digital",
  ]);
  // Projektergebnis & Meilensteine (id 374, nummeriert)
  setParas(374, [
    "Anforderungsspezifikation, mit dem FB MND abgestimmt",
    "Klick-Prototyp des Anmeldeprozesses (Desktop und mobil)",
    "Stufe 1 produktiv: digitale Anmeldung inkl. externer Korreferenten",
    "Stufe 2: Verlängerungsanträge digital",
    "Ergebnis: ein Portal statt drei PDF-Formulare",
  ]);
  // Risiken (id 377)
  setParas(377, [
    "Unklare oder sich ändernde Vorgaben aus den Prüfungsordnungen",
    "Akzeptanz bei Prüfenden und Prüfungsamt; Datenschutz bei externen Beteiligten",
  ]);
  // Abhängigkeiten und Kommentare (id 378)
  setParas(378, [
    "THM-Login, Rollen und Datenschutz (personenbezogene Daten)",
    "Klärung offener Fragen im Interview mit dem Auftraggeber",
  ]);
  // Projektteam (id 380)
  setParas(380, [
    "Auftraggeber: Fachbereich MND (Prof. Dr. Lucke)",
    "Auftragnehmer: Projektteam RE – N.N.",
    "Stakeholder: Studierende, Prüfende, Prüfungsamt",
  ]);

  zip.file("ppt/slides/slide1.xml", xml);
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  const dest = "C:/Users/Ewa/OneDrive - Aurora Engineering GmbH/ZXY/Desktop/Studium/Requirement Eng/One-Pager_ThesisFlow_Template.pptx";
  fs.writeFileSync(dest, out);
  console.log("written", dest, out.length);
  // content check
  console.log((xml.match(/<a:t>[^<]*<\/a:t>/g) || []).map(t => t.slice(5, -6)).join(" | "));
})().catch(e => { console.error(e); process.exit(1); });
