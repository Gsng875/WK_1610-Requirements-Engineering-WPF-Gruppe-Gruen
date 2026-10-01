// Story Mapping ThesisFlow – Version 2
// Quelle: ausschliesslich "Story Mapping/1..7-*.md" (Version 1) im Repository.
// Erzeugt: Story_Mapping_V2.pptx (Baumdiagramm, 3 Ebenen, Zeitstrahl 1..7) und
//          Story_Mapping_V2.md (vollstaendige User Stories mit Zuordnung zu den V1-Nummern).
// Aufruf:  OUT_DIR="../Story Mapping" node storymap_v2.js
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const S = "Studierende:r", EP = "Erstprüfer:in", ZPX = "Externe:r Zweitprüfer:in", PR = "Prüfer:in",
  DEK = "Dekanat", SEK = "Sekretariat", PAV = "PA-Vorsitz", PA = "Prüfungsausschuss", PAMT = "Prüfungsamt",
  ST = "Zuständige Stelle", NA = "Rolle in V1 nicht angegeben";

// ------------------------------------------------------------------ Daten (aus V1 uebernommen)
const A = [
  { name: "Vorbereitung der Anmeldung", v1: "1-Vorbereitung-der-Anmeldung.md (Activity: Anmeldung der Abschlussarbeit vorbereiten)", tasks: [
    { t: "Thema abstimmen", v1: "User Task 1: Abschlussarbeitsthema abstimmen", s: [
      { v1: "US01", r: S, k: "Thema mit Erstprüfer:in abstimmen", f: "Als Studierende:r möchte ich das Thema meiner Abschlussarbeit mit meiner:meinem Erstprüfer:in abstimmen, damit Inhalt und Umfang meiner Arbeit vor der Anmeldung klar feststehen.", ak: "Der abgestimmte Themenvorschlag und die zuständige Erstprüfung sind für den Antrag dokumentiert." } ] },
    { t: "Zweitprüfung vorschlagen", v1: "User Task 2: Geeignete Zweitprüfung vorschlagen", s: [
      { v1: "US02", r: S, k: "Interne oder externe Zweitprüfung vorschlagen", f: "Als Studierende:r möchte ich eine interne oder externe Person für die Zweitprüfung vorschlagen, damit die für meine Anmeldung erforderliche zweite Begutachtung frühzeitig vorbereitet ist.", ak: "Die vorgeschlagene Person ist eindeutig benannt und als interne oder externe Zweitprüfung gekennzeichnet." } ] },
    { t: "Externes Korreferat klären", v1: "User Task 3: Externes Korreferat klären", s: [
      { v1: "US03", r: ZPX, k: "Einverständnis zum Korreferat bestätigen", f: "Als externe:r Zweitprüfer:in möchte ich mein Einverständnis zur Übernahme des Korreferats bestätigen, damit die Studierenden mich mit meiner Zustimmung im Antrag vorschlagen können.", ak: "Die Zustimmung ist eindeutig der prüfenden Person und der betreffenden Abschlussarbeit zugeordnet." },
      { v1: "US04", r: ZPX, k: "Fachliche Qualifikation und Berufserfahrung nachweisen", f: "Als externe:r Zweitprüfer:in möchte ich meine fachliche Qualifikation und einschlägige Berufserfahrung nachweisen, damit der Fachbereich meine Eignung für das Korreferat prüfen kann.", ak: "Die erforderlichen Angaben zum Hochschulabschluss und zur Berufspraxis sowie eine Urkundenkopie liegen zur Prüfung vor." } ] },
    { t: "Starttermin abstimmen", v1: "User Task 4: Gewünschten Starttermin abstimmen", s: [
      { v1: "US05", r: S, k: "Wunschstarttermin mit Erstprüfer:in abstimmen", f: "Als Studierende:r möchte ich den gewünschten Beginn meiner Abschlussarbeit mit meiner:meinem Erstprüfer:in abstimmen, damit ich die Bearbeitungszeit mit meinen weiteren Verpflichtungen planen kann.", ak: "Der abgestimmte Wunschstarttermin ist für den Antrag dokumentiert und vom später verbindlich festgelegten Beginn unterscheidbar." } ] },
  ] },
  { name: "Antragstellung", v1: "2-Antragstellung.md", tasks: [
    { t: "Antragsdaten erfassen", v1: "User Task 2.1 – Antragsdaten erfassen", s: [
      { v1: "US 2.1.1", r: S, k: "Persönliche Angaben im Antrag eintragen", f: "Als Studierende:r möchte ich meine persönlichen Angaben im Antrag eintragen, damit die für die Anmeldung erforderlichen Daten vollständig erfasst sind." },
      { v1: "US 2.1.2", r: S, k: "Prüfer:in auswählen", f: "Als Studierende:r möchte ich meine:n Prüfer:in auswählen, damit die betreuende beziehungsweise prüfende Person dem Antrag zugeordnet werden kann." },
      { v1: "US 2.1.3", r: S, k: "Gewünschten Starttermin angeben", f: "Als Studierende:r möchte ich den gewünschten Starttermin meiner Abschlussarbeit angeben, damit der geplante Bearbeitungsbeginn berücksichtigt werden kann." } ] },
    { t: "Unterlagen ergänzen", v1: "User Task 2.2 – Unterlagen ergänzen", s: [
      { v1: "US 2.2.1", r: S, k: "Erforderliche Nachweise im Portal hochladen", f: "Als Studierende:r möchte ich die erforderlichen Nachweise direkt im Portal hochladen, damit alle Unterlagen gemeinsam mit meinem Antrag eingereicht werden können." },
      { v1: "US 2.2.2", r: S, k: "Nachweise zu externen Prüfungsleistungen hinzufügen", f: "Als Studierende:r möchte ich gegebenenfalls Nachweise zu externen Prüfungsleistungen hinzufügen, damit diese bei der Prüfung meines Antrags berücksichtigt werden können." },
      { v1: "US 2.2.3", r: S, k: "Hochgeladene und fehlende Dokumente sehen", f: "Als Studierende:r möchte ich sehen, welche Dokumente bereits hochgeladen wurden und welche noch fehlen, damit ich meinen Antrag vollständig einreichen kann." } ] },
    { t: "Digital bestätigen", v1: "User Task 2.3 – Unterschrift digital bestätigen", s: [
      { v1: "US 2.3.1", r: S, k: "Antrag digital bestätigen bzw. unterschreiben", f: "Als Studierende:r möchte ich meinen Antrag digital bestätigen beziehungsweise unterschreiben, damit keine zusätzliche Unterschrift auf Papier notwendig ist." },
      { v1: "US 2.3.2", r: PR, k: "Antrag der Studierenden digital bestätigen", f: "Als Prüfer:in möchte ich den Antrag der studierenden Person digital bestätigen, damit meine Zustimmung elektronisch dokumentiert werden kann." },
      { v1: "US 2.3.3", r: S, k: "Sehen, ob alle digitalen Bestätigungen vorliegen", f: "Als Studierende:r möchte ich sehen, ob alle erforderlichen digitalen Bestätigungen vorliegen, damit ich weiß, ob der Antrag weiterbearbeitet werden kann." } ] },
    { t: "Antrag absenden", v1: "User Task 2.4 – Antrag final absenden", s: [
      { v1: "US 2.4.1", r: S, k: "Daten und Unterlagen vor dem Absenden prüfen", f: "Als Studierende:r möchte ich vor dem Absenden alle eingegebenen Daten und hochgeladenen Unterlagen noch einmal überprüfen, damit ich fehlerhafte oder unvollständige Angaben korrigieren kann." },
      { v1: "US 2.4.2", r: S, k: "Antrag verbindlich über das Portal einreichen", f: "Als Studierende:r möchte ich meinen Antrag verbindlich über das Portal einreichen, damit die Anmeldung meiner Abschlussarbeit offiziell gestartet wird." },
      { v1: "US 2.4.3", r: S, k: "Bestätigung der Einreichung erhalten", f: "Als Studierende:r möchte ich nach dem Absenden eine Bestätigung über die erfolgreiche Einreichung erhalten, damit ich sicher weiß, dass mein Antrag übermittelt wurde." },
      { v1: "US 2.4.4", r: ST, k: "Eingereichten Antrag automatisch zur Prüfung erhalten", f: "Als zuständige Stelle möchte ich den eingereichten Antrag automatisch zur weiteren Prüfung erhalten, damit der Bearbeitungsprozess ohne manuelle Weiterleitung fortgesetzt werden kann." } ] },
  ] },
  { name: "Prüfung des Antrags", v1: "3-Prüfung-des-Antrags.md", tasks: [
    { t: "Formalien prüfen", v1: "User Task 3.1 – Dekanat prüft Formalien", s: [
      { v1: "US 3.1.1", r: DEK, k: "Voraussetzungen der Zulassung der Studierenden prüfen", f: "Als Mitarbeiter:in des Dekanats möchte ich die Zulassungsvoraussetzungen der Studierenden prüfen, damit nur Studierende zur Abschlussarbeit zugelassen werden, die alle Voraussetzungen der Prüfungsordnung erfüllen." },
      { v1: "US 3.1.2", r: DEK, k: "Zulassung der externen Prüfer:in prüfen", f: "Als Mitarbeiter:in des Dekanats möchte ich die Zulassung der externen Prüferin bzw. des externen Prüfers prüfen, damit sichergestellt ist, dass die Qualifikation den Vorgaben der Prüfungsordnung entspricht." } ] },
    { t: "Über Antrag entscheiden", v1: "User Task 3.2 – Entscheidung des Prüfungsausschusses", s: [
      { v1: "US 3.2.1", r: PAV, k: "Antrag annehmen", f: "Als Vorsitzende:r des Prüfungsausschusses möchte ich einen Antrag annehmen, damit die:der Studierende offiziell zur Abschlussarbeit zugelassen wird." },
      { v1: "US 3.2.2", r: PAV, k: "Antrag ablehnen und Ablehnung begründen", f: "Als Vorsitzende:r des Prüfungsausschusses möchte ich einen Antrag ablehnen und die Ablehnung begründen, damit die Entscheidung für alle Beteiligten nachvollziehbar ist." },
      { v1: "US 3.2.3", r: PR, k: "Entscheidung des Prüfungsausschusses mitgeteilt bekommen", f: "Als Prüfer:in möchte ich die Entscheidung des Prüfungsausschusses mitgeteilt bekommen, damit ich weiß, ob ich die Abschlussarbeit ausgeben kann." } ] },
  ] },
  { name: "Ausgabe der Arbeit", v1: "4-Ausgabe-der-Arbeit.md", tasks: [
    { t: "Thema bestätigen und freigeben", v1: "User Task 4.1 – Thema bestätigen und freigeben", s: [
      { v1: "US 4.1.1", r: EP, k: "Thema der Abschlussarbeit bestätigen", f: "Als Erstprüfer:in möchte ich das Thema der Abschlussarbeit bestätigen, damit die Arbeit mit dem abgestimmten Thema offiziell freigegeben wird." } ] },
    { t: "Bearbeitungszeit festlegen", v1: "User Task 4.2 – Bearbeitungszeit festlegen", s: [
      { v1: "US 4.2.1", r: EP, k: "Beginn der Bearbeitungszeit festlegen", f: "Als Erstprüfer:in möchte ich den Beginn der Bearbeitungszeit festlegen, damit der Start der Abschlussarbeit eindeutig dokumentiert ist." },
      { v1: "US 4.2.2", r: EP, k: "Abgabedatum festlegen", f: "Als Erstprüfer:in möchte ich das Abgabedatum festlegen, damit die Studierenden wissen, bis wann sie ihre Arbeit abgeben müssen." },
      { v1: "US 4.2.3", r: DEK, k: "Bearbeitungsfrist gegen die Prüfungsordnung prüfen", f: "Als Mitarbeiter:in des Dekanats möchte ich die Bearbeitungsfrist prüfen, damit sichergestellt ist, dass sie den Vorgaben der Prüfungsordnung entspricht." } ] },
    { t: "Thema ausgeben", v1: "User Task 4.3 – Ausgabe des Themas", s: [
      { v1: "US 4.3.1", r: SEK, k: "Ausgabebestätigung über das Portal versenden", f: "Als Mitarbeiter:in des Sekretariats möchte ich die Ausgabebestätigung über das Portal versenden, damit die Ausgabe der Abschlussarbeit offiziell dokumentiert ist." },
      { v1: "US 4.3.2", r: S, k: "Über die Ausgabe des Themas informiert werden", f: "Als Studierende:r möchte ich über die Ausgabe meines Themas informiert werden, damit ich Thema und Bearbeitungsfrist kenne und mit der Arbeit beginnen kann." } ] },
  ] },
  { name: "Durchführung der Arbeit", v1: "5-Durchführung-der-Arbeit.md", tasks: [
    { t: "Zeitraum und Status einsehen", v1: "User Task 5.1 – Bearbeitungszeitraum und Status einsehen", s: [
      { v1: "US 5.1.1", r: S, k: "Beginn der Bearbeitungszeit sehen", f: "Als Studierende:r möchte ich den Beginn meiner Bearbeitungszeit sehen, damit ich weiß, ab wann meine Abschlussarbeit offiziell läuft." },
      { v1: "US 5.1.2", r: S, k: "Aktuellen Abgabetermin sehen", f: "Als Studierende:r möchte ich meinen aktuellen Abgabetermin sehen, damit ich meine Bearbeitung entsprechend planen kann." },
      { v1: "US 5.1.3", r: S, k: "Aktuellen Status der Abschlussarbeit sehen", f: "Als Studierende:r möchte ich den aktuellen Status meiner Abschlussarbeit sehen, damit ich jederzeit den Stand meines Vorgangs kenne." } ] },
    { t: "Fristen verfolgen", v1: "User Task 5.2 – Fristen verfolgen und Erinnerungen erhalten", s: [
      { v1: "US 5.2.1", r: S, k: "Verbleibende Bearbeitungszeit sehen", f: "Als Studierende:r möchte ich sehen, wie viel Bearbeitungszeit noch verbleibt, damit ich meinen Fortschritt besser planen kann." },
      { v1: "US 5.2.2", r: S, k: "Vor wichtigen Fristen automatisch erinnert werden", f: "Als Studierende:r möchte ich vor wichtigen Fristen automatisch erinnert werden, damit ich keine Termine verpasse." },
      { v1: "US 5.2.3", r: S, k: "Über geänderten Abgabetermin informiert werden", f: "Als Studierende:r möchte ich bei einer Änderung des Abgabetermins über die neue Frist informiert werden, damit ich immer mit dem aktuellen Termin arbeite." } ] },
    { t: "Verlängerung beantragen", v1: "User Task 5.3 – Verlängerungsantrag digital stellen", s: [
      { v1: "US 5.3.1", r: S, k: "Verlängerungsantrag im Portal starten", f: "Als Studierende:r möchte ich einen Verlängerungsantrag direkt im Portal starten, damit kein separater Papierantrag notwendig ist." },
      { v1: "US 5.3.2", r: S, k: "Grund für die Verlängerung angeben", f: "Als Studierende:r möchte ich den Grund für meine Verlängerung angeben können, damit mein Antrag nachvollziehbar geprüft werden kann." },
      { v1: "US 5.3.3", r: S, k: "Antrag digital absenden und Bearbeitungsstatus sehen", f: "Als Studierende:r möchte ich meinen Verlängerungsantrag digital absenden und den Bearbeitungsstatus sehen, damit ich weiß, ob mein Antrag bereits geprüft wird." } ] },
    { t: "Stellungnahme abgeben", v1: "User Task 5.4 – Stellungnahme zur Verlängerung abgeben", s: [
      { v1: "US 5.4.1", r: EP, k: "Verlängerungsantrag einsehen", f: "Als Erstprüfer:in möchte ich den Verlängerungsantrag der studierenden Person einsehen, damit ich den Grund beurteilen kann." },
      { v1: "US 5.4.2", r: EP, k: "Stellungnahme zum Verlängerungsantrag abgeben", f: "Als Erstprüfer:in möchte ich eine Stellungnahme zum Verlängerungsantrag abgeben, damit meine Einschätzung in die Entscheidung einfließt." },
      { v1: "US 5.4.3", r: EP, k: "Stellungnahme digital weiterleiten", f: "Als Erstprüfer:in möchte ich die Stellungnahme digital weiterleiten, damit der Antrag anschließend durch die zuständige Stelle geprüft werden kann." } ] },
    { t: "Verlängerung entscheiden", v1: "User Task 5.5 – Verlängerungsantrag prüfen und entscheiden", s: [
      { v1: "US 5.5.1", r: PA, k: "Verlängerungsantrag und Stellungnahme einsehen", f: "Als Prüfungsausschuss möchte ich den Verlängerungsantrag und die Stellungnahme der Erstprüferin bzw. des Erstprüfers einsehen, damit ich alle relevanten Informationen für die Entscheidung habe." },
      { v1: "US 5.5.2", r: PA, k: "Verlängerungsantrag genehmigen oder ablehnen", f: "Als Prüfungsausschuss möchte ich den Verlängerungsantrag genehmigen oder ablehnen, damit der Antrag abschließend entschieden werden kann." },
      { v1: "US 5.5.3", r: PA, k: "Entscheidung und ggf. neuen Abgabetermin hinterlegen", f: "Als Prüfungsausschuss möchte ich die Entscheidung und gegebenenfalls den neuen Abgabetermin im Portal hinterlegen, damit die studierende Person über die gültige Frist informiert wird." } ] },
  ] },
  { name: "Abgabe der Arbeit", v1: "6-Abgabe-der-Arbeit.md", tasks: [
    { t: "Arbeit abgeben", v1: "1. Arbeit abgeben (Studierende:r)", s: [
      { v1: "1.1", r: S, k: "Arbeit als PDF im Portal hochladen", f: "Als Studierende:r möchte ich meine Bachelorarbeit als PDF im Portal hochladen, damit ich sie ohne Ausdruck und ohne Weg zum Sekretariat abgeben kann." },
      { v1: "1.2", r: S, k: "Abgabefrist und Restzeit auf der Abgabeseite sehen", f: "Als Studierende:r möchte ich Abgabefrist und Restzeit auf der Abgabeseite sehen, damit ich weiß, bis wann ich abgeben muss." } ] },
    { t: "Abgabe erfassen", v1: "2. Abgabe erfassen (System)", s: [
      { v1: "2.1", r: DEK, k: "System prüft Dateityp und Größe jeder Abgabe", f: "Als Dekanat möchte ich, dass das System Dateityp und Größe jeder Abgabe prüft, damit nur lesbare, vollständige Dateien eingehen." },
      { v1: "2.2", r: DEK, k: "System hält Abgabezeitpunkt fest und gleicht ihn mit der Frist ab", f: "Als Dekanat möchte ich, dass das System den Abgabezeitpunkt automatisch festhält und mit der Frist abgleicht, damit klar belegt ist, ob rechtzeitig abgegeben wurde." },
      { v1: "2.3", r: DEK, k: "Abgegebene Arbeit unveränderbar sichern", f: "Als Dekanat möchte ich, dass die abgegebene Arbeit unveränderbar gesichert wird, damit sie nach der Abgabe nicht mehr verändert werden kann." } ] },
    { t: "Eingang bestätigen", v1: "3. Eingang bestätigen (Studierende:r)", s: [
      { v1: "3.1", r: S, k: "E-Mail-Bestätigung direkt nach der Abgabe erhalten", f: "Als Studierende:r möchte ich direkt nach der Abgabe eine E-Mail-Bestätigung erhalten, damit ich sicher weiß, dass meine Arbeit angekommen ist." },
      { v1: "3.2", r: S, k: "Abgabequittung mit Zeitstempel und Vorgangsnummer sehen und speichern", f: "Als Studierende:r möchte ich im Portal eine Abgabequittung mit Zeitstempel und Vorgangsnummer sehen und speichern können, damit ich den Eingang bei Rückfragen nachweisen kann." } ] },
    { t: "Beteiligte informieren", v1: "4. Beteiligte informieren (Erstprüfer:in, Zweitprüfer:in, Dekanat)", s: [
      { v1: "4.1", r: EP, k: "Automatisch über die Abgabe benachrichtigt werden", f: "Als Erstprüfer:in möchte ich automatisch benachrichtigt werden, wenn meine Studierende ihre Arbeit abgegeben hat, damit ich ohne Nachfragen mit der Bewertung beginnen kann." },
      { v1: "4.2", r: ZPX, k: "Per E-Mail informiert werden und Arbeit über einen Link einsehen", f: "Als externe:r Zweitprüfer:in möchte ich per E-Mail über die Abgabe informiert werden und die Arbeit über einen Link einsehen können, damit ich ohne THM-Zugang mitwirken kann." },
      { v1: "4.3", r: DEK, k: "Über jede Abgabe mit Name, Matrikelnummer und Thema benachrichtigt werden", f: "Als Dekanat möchte ich über jede Abgabe benachrichtigt werden, mit Name, Matrikelnummer und Thema, damit ich den Eingang sofort der richtigen Akte zuordnen kann." } ] },
  ] },
  { name: "Abschluss der Dokumentation", v1: "7-Abschluss-der-Dokumentation.md (Activity: 7)", tasks: [
    { t: "Antragsstatus einsehen", v1: "1. Antragsstatus einsehen", s: [
      { v1: "1, erster Teil", r: NA, k: "Aktuellen Status sehen, um den Fortschritt zu kennen", f: "Aktuellen Status sehen, um den Fortschritt zu sehen." },
      { v1: "1, zweiter Teil", r: NA, k: "Sehen, ob mein Antrag noch bearbeitet wird", f: "Sehen, ob mein Antrag noch bearbeitet wird, damit ich den aktuellen Stand sehe." } ] },
    { t: "Antragsstatus einsehen (Prüfungsamt)", v1: "2. Antragsstatus einsehen (Prüfungsamt)", s: [
      { v1: "2", r: PAMT, k: "Status einsehen, um den Bearbeitungsstand nachzuvollziehen", f: "Als Prüfungsamt Status einsehen können, um den Bearbeitungsstand nachvollziehen zu können." } ] },
    { t: "Rückmeldung zum Antrag erhalten", v1: "3. Rückmeldung zum Antrag erhalten", s: [
      { v1: "3, erster Teil", r: NA, k: "Rückmeldung erhalten, ob noch Schritte notwendig sind", f: "Rückmeldung erhalten, um zu wissen, ob noch Schritte notwendig sind." },
      { v1: "3, zweiter Teil", r: NA, k: "Über neue Rückmeldungen informiert werden", f: "Über neue Rückmeldung informiert werden, um keine Infos zu verpassen." } ] },
    { t: "Fehlende Unterlagen einreichen", v1: "4. Fehlende Unterlagen einreichen", s: [
      { v1: "4, erster Teil", r: NA, k: "Fehlende Unterlagen einreichen", f: "Fehlende Unterlagen einreichen können, um den Antrag vollständig zu bearbeiten." },
      { v1: "4, zweiter Teil", r: NA, k: "Einsehen, welche Unterlagen fehlen", f: "Einsehen können, welche Unterlagen fehlen, damit man gezielt reagieren kann." } ] },
    { t: "Antragsdokumentation prüfen", v1: "5. Antragsdokumentation prüfen", s: [
      { v1: "5", r: PAMT, k: "Antragsdokumentation einsehen, um Vorgänge nachzuvollziehen", f: "Als Prüfungsamt möchte ich die Antragsdokumentation einsehen können, damit ich Vorgänge nachvollziehen kann." } ] },
    { t: "Abschlussbestätigung erhalten", v1: "6. Abschlussbestätigung erhalten", s: [
      { v1: "6, erster Teil", r: NA, k: "Bestätigung als Nachweis für den Abschluss erhalten", f: "Bestätigung erhalten, um einen Nachweis für den Abschluss zu erhalten." },
      { v1: "6, zweiter Teil", r: NA, k: "Bestätigung im Portal einsehen und später abrufen", f: "Bestätigung im Portal einsehen können, um sie später abrufen zu können." } ] },
    { t: "Antragsdokumentation abrufen", v1: "7. Antragsdokumentation abrufen", s: [
      { v1: "7, erster Teil", r: NA, k: "Antrag und Dokumente nach Abschluss einsehen", f: "Antrag und Dokumente nach Abschluss einsehen können, um darauf zugreifen zu können." },
      { v1: "7, zweiter Teil", r: NA, k: "Relevante Dokumente herunterladen", f: "Relevante Dokumente herunterladen können, um meine Unterlagen speichern zu können." } ] },
  ] },
];

// ------------------------------------------------------------------ Zaehlung / Konsistenzpruefung
const nTasks = A.reduce((n, a) => n + a.tasks.length, 0);
const nStories = A.reduce((n, a) => n + a.tasks.reduce((m, t) => m + t.s.length, 0), 0);
if (A.length !== 7) throw new Error("Es muessen genau 7 Activities sein");
A.forEach((a, ai) => a.tasks.forEach((t, ti) => { if (!t.s.length) throw new Error(`User Task ${ai + 1}.${ti + 1} ohne User Story`); }));
console.log(`Activities: ${A.length}, User Tasks: ${nTasks}, User Stories: ${nStories}`);

// ------------------------------------------------------------------ Folie
const pres = new pptxgen();
const W = 56, H = 12.7;
pres.defineLayout({ name: "STORYMAP", width: W, height: H });
pres.layout = "STORYMAP";
pres.title = "Story Mapping ThesisFlow V2";

const C_ACT = "2C5F2D", C_ACT_TXT = "FFFFFF", C_TASK = "DCE8F5", C_TASK_LINE = "4A7FB5", C_TASK_TXT = "1F3A5C",
  C_STORY = "FFF7E0", C_STORY_LINE = "C9A227", INK = "1F2A1F", MUTED = "5C6B5C", MOSS = "97BC62", LINE = "7A867A",
  BAND1 = "EEF4E8", BAND2 = "F1F6FB", BAND3 = "FDFBF3";
const s = pres.addSlide();
s.background = { color: "FFFFFF" };

// Geometrie
const X0 = 2.3, COLW = 1.68, TG = 0.1, AG = 0.36;
const Y_TL = 2.3, Y_ACT = 2.9, H_ACT = 0.9, Y_BUS = 4.1, Y_TASK = 4.4, H_TASK = 0.85, Y_ST = 5.7, H_ST = 1.15, P_ST = 1.27;
const maxStories = Math.max(...A.flatMap(a => a.tasks.map(t => t.s.length)));
const Y_END = Y_ST + maxStories * P_ST - (P_ST - H_ST);

let x = X0;
A.forEach(a => {
  a.x = x;
  a.tasks.forEach((t, i) => { t.x = x; x += COLW + (i < a.tasks.length - 1 ? TG : 0); });
  a.w = x - a.x;
  x += AG;
});
const XR = x - AG;

// Kopf
s.addText("Story Mapping ThesisFlow – Version 2", { x: 0.5, y: 0.35, w: 30, h: 0.85, margin: 0, isTextBox: true, fontFace: "Cambria", fontSize: 36, bold: true, color: INK, valign: "middle" });
s.addText(`Digitale Anmeldung und Verwaltung von Abschlussarbeiten im FB MND  ·  ${A.length} Activities  ·  ${nTasks} User Tasks  ·  ${nStories} User Stories  ·  Leserichtung von links nach rechts (Schritt 1 bis 7)`, { x: 0.5, y: 1.2, w: 40, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 14, color: MUTED });
s.addText("REQUIREMENTS ENGINEERING  ·  GRUPPE GRÜN  ·  STAND 01.10.2026", { x: W - 12.5, y: 0.5, w: 12, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 11, bold: true, color: MOSS, charSpacing: 2, align: "right", valign: "middle" });

// Ebenen-Baender
const band = (y, h, color) => s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: XR + 0.2 - 0.5, h, fill: { color }, line: { color } });
band(Y_ACT - 0.15, H_ACT + 0.3, BAND1);
band(Y_TASK - 0.15, H_TASK + 0.3, BAND2);
band(Y_ST - 0.15, Y_END - Y_ST + 0.3, BAND3);
const lvl = (y, h, n, name, fill, lineC, txt) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.65, y: y + h / 2 - 0.3, w: 1.4, h: 0.6, fill: { color: fill }, line: { color: lineC, width: 1.25 }, rectRadius: 0.06 });
  s.addText([{ text: "EBENE " + n, options: { fontSize: 8, bold: true, breakLine: true } }, { text: name, options: { fontSize: 11, bold: true } }], { x: 0.65, y: y + h / 2 - 0.3, w: 1.4, h: 0.6, margin: 2, isTextBox: true, fontFace: "Calibri", color: txt, align: "center", valign: "middle" });
};
lvl(Y_ACT, H_ACT, 1, "Activity", C_ACT, C_ACT, C_ACT_TXT);
lvl(Y_TASK, H_TASK, 2, "User Task", C_TASK, C_TASK_LINE, C_TASK_TXT);
lvl(Y_ST, H_ST, 3, "User Story", C_STORY, C_STORY_LINE, INK);

// Zeitstrahl
s.addText("ZEITLICHER ABLAUF", { x: 0.5, y: Y_TL - 0.2, w: 1.7, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9, bold: true, color: MUTED, charSpacing: 1, align: "center", valign: "middle" });
s.addShape(pres.shapes.LINE, { x: X0, y: Y_TL, w: XR - X0 + 0.15, h: 0.001, line: { color: C_ACT, width: 2.5, endArrowType: "triangle" } });
A.forEach((a, i) => {
  const cx = a.x + a.w / 2, d = 0.62;
  s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: Y_TL - d / 2, w: d, h: d, fill: { color: C_ACT }, line: { color: "FFFFFF", width: 2 } });
  s.addText(String(i + 1), { x: cx - d / 2, y: Y_TL - d / 2, w: d, h: d, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
});

const vline = (xx, y1, y2) => s.addShape(pres.shapes.LINE, { x: xx, y: y1, w: 0.001, h: y2 - y1, line: { color: LINE, width: 1.5 } });
const hline = (x1, x2, yy) => s.addShape(pres.shapes.LINE, { x: x1, y: yy, w: x2 - x1, h: 0.001, line: { color: LINE, width: 1.5 } });

A.forEach((a, ai) => {
  const an = ai + 1, acx = a.x + a.w / 2;
  // Verbindungen Zeitstrahl -> Activity -> User Tasks
  vline(acx, Y_TL + 0.31, Y_ACT);
  vline(acx, Y_ACT + H_ACT, Y_BUS);
  hline(a.tasks[0].x + COLW / 2, a.tasks[a.tasks.length - 1].x + COLW / 2, Y_BUS);
  a.tasks.forEach(t => {
    const tcx = t.x + COLW / 2;
    vline(tcx, Y_BUS, Y_TASK);
    // Stamm von User Task zu den User Stories (Karten liegen darueber)
    vline(tcx, Y_TASK + H_TASK, Y_ST + (t.s.length - 1) * P_ST + 0.05);
  });
  // Activity
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: a.x, y: Y_ACT, w: a.w, h: H_ACT, fill: { color: C_ACT }, line: { color: C_ACT }, rectRadius: 0.1 });
  s.addText([
    { text: "ACTIVITY " + an, options: { fontSize: 9, bold: true, color: MOSS, charSpacing: 2, breakLine: true } },
    { text: a.name, options: { fontFace: "Cambria", fontSize: 17, bold: true, color: C_ACT_TXT } },
  ], { x: a.x, y: Y_ACT, w: a.w, h: H_ACT, margin: 4, isTextBox: true, fontFace: "Calibri", align: "center", valign: "middle" });
  // User Tasks und User Stories
  a.tasks.forEach((t, ti) => {
    const tn = `${an}.${ti + 1}`;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: t.x, y: Y_TASK, w: COLW, h: H_TASK, fill: { color: C_TASK }, line: { color: C_TASK_LINE, width: 1.25 }, rectRadius: 0.07 });
    s.addText([
      { text: "USER TASK " + tn, options: { fontSize: 7.5, bold: true, color: C_TASK_LINE, charSpacing: 1, breakLine: true } },
      { text: t.t, options: { fontSize: 10.5, bold: true, color: C_TASK_TXT } },
    ], { x: t.x, y: Y_TASK, w: COLW, h: H_TASK, margin: 5, isTextBox: true, fontFace: "Calibri", align: "center", valign: "middle" });
    t.s.forEach((st, si) => {
      const y = Y_ST + si * P_ST;
      st.id = `${tn}.${si + 1}`;
      s.addShape(pres.shapes.RECTANGLE, { x: t.x, y, w: COLW, h: H_ST, fill: { color: C_STORY }, line: { color: C_STORY_LINE, width: 1 } });
      s.addText([
        { text: "US " + st.id, options: { fontSize: 7.5, bold: true, color: "7A5C00", breakLine: true } },
        { text: st.k, options: { fontSize: 10, color: INK, breakLine: true } },
        { text: st.r === NA ? "Rolle offen" : st.r, options: { fontSize: 7.5, italic: true, color: MUTED } },
      ], { x: t.x, y, w: COLW, h: H_ST, margin: 6, isTextBox: true, fontFace: "Calibri", align: "left", valign: "middle" });
    });
  });
});

// Legende
const LY = Y_END + 0.55;
s.addText("Legende", { x: 0.5, y: LY, w: 2, h: 0.3, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 11, bold: true, color: INK });
const leg = (xx, fill, lineC, text, round) => {
  s.addShape(round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, Object.assign({ x: xx, y: LY + 0.42, w: 0.5, h: 0.3, fill: { color: fill }, line: { color: lineC, width: 1 } }, round ? { rectRadius: 0.05 } : {}));
  s.addText(text, { x: xx + 0.6, y: LY + 0.37, w: 4.2, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED, valign: "middle" });
};
leg(0.5, C_ACT, C_ACT, `Ebene 1: Activity (${A.length}), Prozessschritt im Zeitstrahl`, true);
leg(5.6, C_TASK, C_TASK_LINE, `Ebene 2: User Task (${nTasks}), gehört zur Activity darüber`, true);
leg(10.9, C_STORY, C_STORY_LINE, `Ebene 3: User Story (${nStories}), gehört zum User Task darüber`, false);
s.addShape(pres.shapes.LINE, { x: 16.6, y: LY + 0.57, w: 0.5, h: 0.001, line: { color: LINE, width: 1.5 } });
s.addText("Zuordnung (Baumstruktur)", { x: 17.2, y: LY + 0.37, w: 2.6, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED, valign: "middle" });
s.addShape(pres.shapes.LINE, { x: 20.2, y: LY + 0.57, w: 0.7, h: 0.001, line: { color: C_ACT, width: 2.5, endArrowType: "triangle" } });
s.addText("Zeitlicher Ablauf, Schritt 1 bis 7", { x: 21.0, y: LY + 0.37, w: 3.2, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED, valign: "middle" });
s.addText("User Stories sind als Kurztext dargestellt. Vollständiger Wortlaut, Rollen und die Zuordnung zu den Nummern der Version 1 stehen in Story_Mapping_V2.md. „Rolle offen“: In Version 1 ist für diese User Story keine Rolle angegeben.", { x: 25.0, y: LY + 0.2, w: 22, h: 0.75, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 10.5, color: MUTED, valign: "middle" });
s.addText("Quelle: Story Mapping/1 bis 7 (Version 1) im Repository  ·  Aufgabe: User Story Mapping (Praxisteil, Seite 10)", { x: W - 16.5, y: LY + 0.37, w: 16, h: 0.4, margin: 0, isTextBox: true, fontFace: "Calibri", fontSize: 9.5, color: "8A968A", align: "right", valign: "middle" });

s.addNotes("Story Mapping ThesisFlow, Version 2. Drei Ebenen: Activity, User Task, User Story. Zeitstrahl von links nach rechts, Schritt 1 bis 7. Inhalte ausschliesslich aus Story Mapping/1 bis 7 (Version 1) im Repository; Kurztexte mit einheitlicher Nummerierung Activity.Task.Story.");

// ------------------------------------------------------------------ Markdown mit vollstaendigen Stories und V1-Zuordnung
let md = `# Story Mapping ThesisFlow – Version 2\n\n`;
md += `Stand: 01.10.2026 · ${A.length} Activities · ${nTasks} User Tasks · ${nStories} User Stories\n\n`;
md += `Grafische Fassung: \`Story_Mapping_V2.pptx\` (bearbeitbar) und \`Story_Mapping_V2.pdf\` (Export). Erzeugt mit \`Skripte/storymap_v2.js\`.\n\n`;
md += `Quelle sind ausschließlich die Dateien \`1-…\` bis \`7-…\` in diesem Ordner (Version 1). Diese bleiben unverändert erhalten.\n`;
md += `Die Nummerierung ist in Version 2 einheitlich: **Activity.User Task.User Story**. Die ursprüngliche Nummer aus Version 1 steht jeweils dahinter.\n\n`;
md += `## Zeitlicher Ablauf\n\n`;
md += A.map((a, i) => `${i + 1} ${a.name}`).join(" → ") + `\n\n`;
A.forEach((a, ai) => {
  md += `## Activity ${ai + 1} – ${a.name}\n\nQuelle V1: \`${a.v1}\`\n\n`;
  a.tasks.forEach((t, ti) => {
    md += `### User Task ${ai + 1}.${ti + 1} – ${t.t}\n\nV1: „${t.v1}“\n\n`;
    md += `| Nr. V2 | Nr. V1 | Rolle | Kurztext im Story Mapping | Vollständige User Story |\n|---|---|---|---|---|\n`;
    t.s.forEach(st => {
      md += `| US ${st.id} | ${st.v1} | ${st.r} | ${st.k} | ${st.f}${st.ak ? " **Akzeptanzkriterium:** " + st.ak : ""} |\n`;
    });
    md += `\n`;
  });
});
md += `## Änderungen gegenüber Version 1\n\n`;
md += `- **Einheitliche Struktur:** Alle sieben Dateien nutzten unterschiedliche Gliederungen (Datei 1 mit Epic und „User Task 1“, Dateien 2 bis 5 mit „User Task 2.1“, Datei 6 mit „1. Arbeit abgeben“, Datei 7 als nummerierte Stichpunkte). Version 2 verwendet durchgehend Activity, User Task und User Story.\n`;
md += `- **Einheitliche Nummerierung:** US01 bis US05, US 2.1.1, 1.1 und unnummerierte Stichpunkte wurden auf das Schema Activity.User Task.User Story umgestellt, mit Rückverweis auf die V1-Nummer.\n`;
md += `- **Zuordnung korrigiert:** In Datei 6 waren die User Tasks mit 1 bis 4 ohne Bezug zur Activity 6 nummeriert, in Datei 7 standen mehrere User Stories ohne Nummer in einer Zeile. Jede User Story hängt jetzt eindeutig an einem User Task und jeder User Task an einer Activity.\n`;
md += `- **Einheitliche Begriffe:** Rollen einheitlich als Studierende:r, Erstprüfer:in, externe:r Zweitprüfer:in (statt „Studierende“, „meiner Erstprüfer“, „externe Zweitprüfer“, „Der Erstprüfer:in“). Im Kurztext 6.1.1 steht „Arbeit“ statt „Bachelorarbeit“, weil das Story Mapping Bachelor- und Masterarbeiten abdeckt; der Originalwortlaut steht in der Tabelle.\n`;
md += `- **Rechtschreibung:** Korrigiert wurden unter anderem „Antragsstats“, „Rückmelding“, „Rückmledung“, „Fehlene“, „unterlagn“, „Antragsdokumenation“, „dek Abschluss“, „Abachluss“ (Datei 7), „meinung … Einschätzen“ (Datei 5) und der fehlende Satzpunkt in US 3.2.3.\n`;
md += `- **Unklare Formulierungen:** US 5.4.2 und 5.4.3 sprachen von der „Meinung“ bzw. „Stellungnahme des Studierenden“, obwohl der User Task die Stellungnahme der Erstprüfer:in beschreibt. Der Wortlaut wurde entsprechend dem User Task 5.4 geglättet. US 5.4.1 bis 5.5.3 sind jetzt im Format „Als … möchte ich … damit …“ formuliert.\n`;
md += `- **Kurztexte:** Jede User Story hat einen kurzen Text für die Grafik, die fachliche Aussage bleibt unverändert.\n`;
md += `- **Darstellung:** Baumdiagramm mit drei farblich getrennten Ebenen, Zeitstrahl 1 bis 7 von links nach rechts, Verbindungslinien für jede Zuordnung.\n\n`;
md += `## Offene Punkte aus Version 1 (nicht verändert, zur Klärung im Team)\n\n`;
md += `- **Rollen in Activity 7:** Für die User Stories 7.1.1, 7.1.2, 7.3.1, 7.3.2, 7.4.1, 7.4.2, 7.6.1, 7.6.2, 7.7.1 und 7.7.2 nennt Version 1 keine Rolle. In der Grafik steht deshalb „Rolle offen“.\n`;
md += `- **Inhaltliche Überschneidungen:** Status einsehen kommt in US 5.1.3 und im User Task 7.1 vor, der gewünschte Starttermin in US 1.4.1 und US 2.1.3, fehlende Unterlagen in US 2.2.3 und im User Task 7.4. Die Stories wurden nicht zusammengelegt, weil das eine fachliche Entscheidung ist.\n`;
md += `- **Bezeichnung der Verwaltungsrolle:** Version 1 verwendet „Dekanat“, „Sekretariat“, „Prüfungsamt“ und „zuständige Stelle“. Ob damit dieselbe Rolle gemeint ist, geht aus den Dateien nicht hervor; die Bezeichnungen wurden je Story beibehalten.\n`;
md += `- **Akzeptanzkriterien:** Nur Activity 1 enthält Akzeptanzkriterien. Sie wurden übernommen, für die übrigen Activities wurden keine ergänzt.\n`;
md += `- **Zeitliche Lage von Activity 7:** Die User Tasks 7.1 bis 7.4 (Status, Rückmeldung, fehlende Unterlagen) betreffen den Antrag und liegen zeitlich vor der Abgabe. Die vorgegebene Reihenfolge 1 bis 7 wurde nicht verändert.\n`;

const outDir = process.env.OUT_DIR || ".";
fs.writeFileSync(path.join(outDir, "Story_Mapping_V2.md"), md, "utf8");
pres.writeFile({ fileName: path.join(outDir, "Story_Mapping_V2.pptx") }).then(f => console.log("written", f));
