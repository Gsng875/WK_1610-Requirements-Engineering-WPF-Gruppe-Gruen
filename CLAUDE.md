# Requirements Engineering – Praxisteil (THM, Prof. Dr. Carsten Lucke)

## Sprache und Arbeitsweise
- Immer auf Deutsch antworten und alle Dateien auf Deutsch erstellen.
- Ergebnisse in diesem Ordner ablegen. Textauszüge der Quellen liegen in `Quellen/`.
- Vor jeder neuen Aufgabe die Vorgaben des Professors aus den PDFs in diesem Ordner lesen
  (z. B. `02_Personas.pdf`) und die Ausarbeitung daran ausrichten.
- Zuerst kurz klären, wenn eine Aufgabe mehrere sinnvolle Auslegungen hat, dann bauen.

## Projekt: Abschlussarbeiten-Portal MND
- Webbasiertes Portal zur digitalen Anmeldung und Verwaltung von Abschlussarbeiten im
  Fachbereich MND der THM (Wirtschaftsinformatik B.Sc. und M.Sc.).
- Ersetzt den heutigen PDF-Antrag. Muss im Browser am Desktop und am Smartphone komfortabel nutzbar sein.
- Externe Korreferenten müssen in den Anmeldeprozess eingebunden werden können.
- Ausbaustufe 2: Verlängerungsanträge.
- Rollen: Wir sind der Auftragnehmer (Softwarefirma). Der Dozent Prof. Dr. Carsten Lucke ist
  der Auftraggeber (Fachbereich MND). Offene Fragen dürfen per Interview geklärt werden.
- Projektname: **ThesisFlow**. In künftigen Unterlagen diesen Namen verwenden. Die bisherigen Folien
  tragen noch „ThesisPortal“ bzw. „Abschlussarbeiten-Portal MND“ und werden nur auf Wunsch umbenannt.
- Tagline: „Anmelden. Betreuen. Abschließen.“

## Bisherige Ergebnisse (Stand 30.09.2026)
- `One-Pager_Abschlussarbeiten-Portal.pptx` – freie Gestaltung, Grundlage für das Foliendesign.
- `One-Pager_Abschlussarbeiten-Portal_Template.pptx` – gleiche Inhalte im THM-Template aus `Downloads/OnePager - Template.pptx`.
- `Personas_Abschlussarbeiten-Portal_v2.pptx` – drei Personas: Moritz Hoffmann (Bachelor),
  Michael Krüger (Sekretariat/Prüfungsamt), Tarek Yilmaz (Master, berufsbegleitend, externer Korreferent).
  Die Datei ohne `_v2` ist überholt. Im Repo liegt die Kopie unter `Unterlagen/Praxis/`.
- `Kontextdiagramm_ThesisFlow_v2.pptx` – Kontextdiagramm in Datenfluss-Notation nach Vorlesung Kap. 2
  (Systemgrenze, Kontextgrenze, Grauzone; Personen, Systeme, Dokumente, Ereignisse; grüne Input-,
  blaue Output-Pfeile). Skript `Skripte/kontext2.js`, Format 16:9 breit (13,33 x 7,5 Zoll). Die v1 ist überholt.

## Git
- Repo: https://github.com/Gsng875/WK_1610-Requirements-Engineering-WPF-Gruppe-Gruen (Branch `main`).
- Original-PDFs versioniert das Team unter `Unterlagen/Praxis/` und `Unterlagen/Theorie/`; `Quellen/*.pdf` ist
  per `.gitignore` ausgeschlossen, dort nur Textauszüge und Auswertungen.
- Vor dem Push immer `git pull --rebase` oder fetch + merge, das Team pusht parallel. Rebase scheitert, wenn eine
  PPTX in PowerPoint geöffnet ist (Datei gesperrt): dann merge statt rebase.
- Vorlesungsfolien: `Unterlagen/Theorie/Requirements Engineering - Blockveranstaltung - v4.pdf`
  (Kap. 2 Systemkontext, Kap. 6 Use-Case-Diagramme). Vor Modellierungsaufgaben das passende Kapitel lesen.

## Design-Regeln für PowerPoint
- Format 16:9 (10 x 5,625 Zoll), erzeugt mit pptxgenjs. Die Generator-Skripte liegen in `Skripte/`
  (`onepager.js`, `filltpl.js` für das THM-Template, `personas2.js` mit Bildern in `Skripte/photos/`).
  Zum Ausführen in einem Scratchpad-Ordner `npm install pptxgenjs jszip sharp`, dann
  `OUT="<Zielpfad>.pptx" node personas2.js` (Bildpfade relativ zum Skriptordner).
- Farben: Dunkelgrün `2C5F2D`, Moosgrün `97BC62`, heller Kartenhintergrund `EEF4E8`,
  Text `1F2A1F`, gedämpft `5C6B5C`, Weiß.
- Schriften: Cambria für Titel und Namen, Calibri für Fließtext.
- Motiv: nummerierte grüne Kreise vor Abschnittstiteln, abgerundete helle Karten, Fünf-Punkte-Skalen.
- Personas-Layout (v2): Titel „Persona: …“, links helle Karte mit Bild, Name, Rolle, Zitat, Steckbrief;
  rechts 2x2 Blöcke Ziele, Eigenschaften und Verhalten, Frustrationen, Nutzungsüberwindung.
- Keine Farbbalken oder Akzentstreifen, keine Linien unter Titeln.

## Umgebung
- Kein Python, kein LibreOffice, kein pdftoppm auf diesem Rechner. Folien können nicht gerendert werden,
  deshalb den Nutzer bitten, Textüberlauf in PowerPoint zu prüfen.
- Verfügbar: Node 24, pptxgenjs, jszip, sharp (im Scratchpad), pdftotext aus Git for Windows.
- Wenn eine PPTX in PowerPoint geöffnet ist, schlägt das Überschreiben fehl: neue Version daneben speichern.

## Quellen aus dem Projektbriefing
Details und Auszüge: siehe `Quellen/README.md`.
1. Antrag auf Zulassung und Ausgabe der Abschlussarbeit (ab WS23):
   https://www.thm.de/mnd/studium/service/edokumente/antrag-auf-zulassung-und-ausgabe-der-abschlussarbeit-ab-ws23/viewdocument/1140
   – vom Team geliefert und eingelesen: `Quellen/Antrag_Zulassung_Abschlussarbeit_WS2324.pdf`,
   Auswertung mit Feldern, 7-Schritte-Ablauf und Regeln in `Quellen/Antrag_Zulassung_Auswertung.md`.
2. Anlage Abschlussarbeit WK, externer Prüfer:
   https://www.thm.de/mnd/studium/service/edokumente/anlage-abschlussarbeit-wk-externer-pruefer-1/viewdocument/41
   – eingelesen, Text in `Quellen/Anlage_Externer_Pruefer_Text.txt`.
3. Antrag Verlängerung der Abschlussarbeit:
   https://www.thm.de/mnd/studium/service/edokumente/verlaengerung-der-abschlussarbeit-antrag/viewdocument/42
   – vom Team geliefert und eingelesen: `Quellen/Antrag_Verlaengerung_Bachelorarbeit.pdf`
   (Adressat Studienausschuss Wirtschaftsinformatik, Befürwortung durch 1. Referent*in).
4. Prüfungsordnung B.Sc. Wirtschaftsinformatik (Übersichtsseite, aktuelle Fassung 2024 Version 2):
   https://www.thm.de/site/thm-dokumente/studium/modulhandbuecher-studien-und-pruefungsordnungen-studienganginfos/fb-13-mnd-mathematik-naturwissenschaften-und-datenverarbeitung/pruefungsordnungen/wirtschaftsinformatik-bachelor.html
   – eingelesen, Text in `Quellen/PO_Bachelor_WI_2024_v2_Text.txt`.
5. Prüfungsordnung M.Sc. Wirtschaftsinformatik 2024 Version 1:
   https://www.thm.de/site/thm-dokumente/1674-pruefungsordnung-wirtschaftsinformatik-master-2024-version-1/download.html
   – eingelesen, Text in `Quellen/PO_Master_WI_2024_v1_Text.txt`.
6. Öffentliche Seite „Abschlussarbeit“ des FB MND mit Ist-Prozess und Voraussetzungen:
   https://www.thm.de/mnd/studium/service/abschlussarbeit
   – eingelesen, Zusammenfassung in `Quellen/Abschlussarbeit_Prozessseite_MND.md`.
Alle Formulare liegen jetzt als PDF in `Quellen/`. Noch offen: Allgemeine Bestimmungen der THM (Teil I, § 17 und § 18).

## Ist-Prozess der Anmeldung (aus dem Zulassungsantrag, Kurzfassung)
Studierende(r) füllt Antrag aus → 2. Prüfer(in) bestätigt Korreferat → MND-Sekretariat prüft Zulassung und
Qualifikation externer Prüfender → PA-Vorsitz entscheidet (Zulassung, externe(r) Prüfer(in), Begründung bei
Ablehnung) → 1. Prüfer(in) gibt Arbeit aus (Beginn, Abgabe) → Sekretariat prüft Frist, legt in Akte →
Sekretariat informiert Studierende(n). Thema nur einmal und nur binnen vier Wochen zurückgebbar (Allg. PO § 17 (3)).
Verlängerung: eigener Antrag an den Studienausschuss Wirtschaftsinformatik mit Thema, Firma, Grund und
Befürwortung der/des 1. Referent*in.

## Fachliche Regeln, die das Portal abbilden muss (aus den Quellen)
- Externer 2. Referent (Anlage): Name, E-Mail, Telefon, Hochschule, akademischer Grad, Abschlussjahr,
  Fachgebiet; nach § 18 Abs. 2 HHG gleichwertige Qualifikation und mindestens fünf Jahre einschlägige
  Berufspraxis nach dem Abschluss; Kopie der Urkunde ist Pflichtanlage; Unterschrift und Datum.
- Bachelorarbeit (PO B.Sc. § 6): Zulassung erst, wenn bis einschließlich 6. Semester alle bis auf maximal
  drei Module bestanden sind; Berufspraktische Phase (Projektphase und Projektphasenseminar) muss vor
  Beginn abgeschlossen sein; 12 CrP, Bearbeitungszeit drei Monate (§ 18 Abs. 1 Teil I); Kolloquium 3 CrP,
  dafür müssen alle Module bestanden sein.
- Masterarbeit (PO M.Sc.): Modul „Masterarbeit mit Kolloquium“ WK_2115, 30 CrP, 3. Semester; der
  Schwerpunkt wird mit der Zulassung zur Masterarbeit über die erbrachten Module festgelegt; Auflagen
  (bis 30 CrP) müssen spätestens zur Zulassung zur Masterarbeit nachgewiesen sein.
- Bearbeitungszeit, Verlängerung und Referentenregeln im Detail stehen in Teil I (Allgemeine Bestimmungen
  der THM, AMB 39/2014 bzw. AMB 01/2015), die nicht in diesem Ordner liegen.
