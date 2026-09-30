# ThesisFlow – Requirements Engineering (WK_1610), Gruppe Grün

Praxisprojekt im Modul Requirements Engineering (Prof. Dr. Carsten Lucke, THM, Fachbereich MND).
Wir entwickeln als Auftragnehmer die Anforderungen für **ThesisFlow**, ein webbasiertes Portal zur digitalen
Anmeldung und Verwaltung von Abschlussarbeiten im Fachbereich MND (Wirtschaftsinformatik B.Sc. und M.Sc.).

## Ergebnisse
| Aufgabe | Datei |
|---|---|
| One-Pager / Produktvision (freies Design) | `One-Pager_ThesisFlow.pptx` |
| One-Pager im THM-Template | `One-Pager_ThesisFlow_Template.pptx` |
| Personas (Bachelor, Sekretariat, Master) | `Personas_ThesisFlow.pptx` |
| Kontextdiagramm: Folie 1 Soll (ThesisFlow, mit Grauzone), Folie 2 Ist (Laufweg des Formulars, Schwachstellen) | `Kontextdiagramm_ThesisFlow_v2.pptx` |

Die Personas liegen im Repo unter `Unterlagen/Praxis/`. Weitere Folien werden bei Bedarf aus `Skripte/` erzeugt.

## Ordner
- `Aufgabe 1-3/` – abgegebene Aufgaben.
- `Unterlagen/Praxis/` – Aufgabenstellung, Formulare der THM (Zulassungsantrag, Verlängerung, externe Prüfende),
  Prüfungsordnungen B.Sc. und M.Sc. als PDF.
- `Unterlagen/Theorie/` – Vorlesungsfolien, User-Stories-Übung, Lehrbuch Pohl/Rupp.
- `Quellen/` – Textauszüge der Formulare und Prüfungsordnungen sowie Auswertungen
  (`README.md`, `Antrag_Zulassung_Auswertung.md` mit Feldliste und 7-Schritte-Ablauf, `Abschlussarbeit_Prozessseite_MND.md`).
- `Skripte/` – Node-Skripte (pptxgenjs), mit denen die Folien erzeugt werden, inkl. Persona-Bilder.
- `CLAUDE.md` – Arbeitsanweisungen und Projektkontext für Claude Code.

## Folien neu erzeugen
```bash
cd Skripte
npm install pptxgenjs jszip sharp
OUT="../Personas_ThesisFlow.pptx" node personas2.js
OUT="../Kontextdiagramm_ThesisFlow_v2.pptx" node kontext2.js
```

## Aufgabenstellung (Kurzfassung)
- PDF-Antrag durch ein Webportal ersetzen, nutzbar am Desktop und am Smartphone.
- Externe Korreferenten in den Anmeldeprozess einbinden.
- Ausbaustufe 2: Verlängerungsanträge.
- Anforderungen aus Antragsformularen, Prüfungsordnungen und Interviews mit dem Auftraggeber.
