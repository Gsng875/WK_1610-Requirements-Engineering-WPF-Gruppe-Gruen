# ThesisFlow – eigenständige Prüfer-Anwendung

Kompakter, anklickbarer Prototyp für **„Thema durch Erstprüfer bestätigen und freigeben“**.
HTML/CSS/JavaScript, Node.js 24 und SQLite; keine npm-Pakete, keine Installation von Node auf dem eigenen Rechner nötig.
Fiktive Beispieldaten orientieren sich an eurem ThesisFlow-Projekt: Fachbereich MND, Rollen und vorgelagerte Zulassung.

## In VS Code starten

1. ZIP entpacken und den Ordner **thesisflow-pruefer** in VS Code öffnen.
2. Docker Desktop starten.
3. In VS Code **Terminal → Neues Terminal** öffnen. Im Ordner mit `compose.yaml` ausführen:

   ```bash
   docker compose up --build
   ```

4. Im Browser **http://localhost:8080** öffnen (genau diese Adresse verwenden).
5. Als `weber` mit `PrueferDemo!2026` anmelden.

**Beenden:** im Terminal `Strg+C`, anschließend `docker compose down`.
**Nach Codeänderungen:** erneut `docker compose up --build`.
Die Daten bleiben im Docker-Volume gespeichert. Das Programm ist nur über den eigenen Rechner erreichbar.

## Durchklicken

- **Meine Themen → Thema prüfen**: Vorschlag von Moritz Hoffmann lesen.
- Einschätzung wählen, Rückmeldung mit mindestens zehn Zeichen eingeben; optional zunächst speichern.
- Für die Freigabe **„Geeignet“** wählen und beide fachlichen Prüfkriterien bestätigen.
- **Thema freigeben** oder **Ablehnen** wählen und im Dialog ausdrücklich bestätigen.
- Entscheidung, Rückmeldung und Verlauf bleiben nach Neuladen gespeichert. Endgültige Entscheidungen sind gesperrt.
- Zweites Thema: Die Zulassung fehlt. Einschätzung und Ablehnung sind möglich, Freigabe ist gesperrt.
- **Zugriffsregeln** zeigt die Regeln im Interface.

„Bewertung“ ist eine fachliche Einschätzung des Themas, keine Abschlussnote. Die offizielle Ausgabe und
Fristberechnung sind weitere Prozessschritte. Geplanter Start und Themenfreigabe setzen noch keine Bearbeitungsfrist in Gang.
Der Prototyp verschickt keine E-Mails und besitzt keine echte Verbindung zum THM-Portal.

## Demo-Zugänge und Sicherheitsregeln

| Benutzer | Passwort | Berechtigung |
|---|---|---|
| `weber` | `PrueferDemo!2026` | Zwei zugeordnete Themen |
| `klein` | `AndereDemo!2026` | Ausschließlich sein eigenes Beispielthema |
| `student` | `StudentDemo!2026` | Zugang wird verweigert |

- Die Rolle und Themenzuordnung sind im Backend hinterlegt. Es gibt keine freie Registrierung oder Rollenauswahl.
- Passwörter werden mit scrypt und individuellem Salt gehasht; Sitzungen verwenden zufällige HttpOnly-/SameSite-Cookies.
- Sitzungen laufen nach 30 Minuten ohne API-Aktivität ab; nach fünf falschen Anmeldungen gilt eine Sperre bis zum Ende des 15-Minuten-Fensters.
- Änderungsanfragen benötigen die richtige Origin, angemeldete Anfragen zusätzlich ein CSRF-Token.
- Bewertungen werden serverseitig validiert. Unvollständige Angaben verhindern eine endgültige Entscheidung;
  Freigaben benötigen die vorherige Zulassung, „Geeignet“ und beide bestätigten Prüfkriterien.
- Entscheidung und Protokoll werden in einer gemeinsamen Datenbanktransaktion gespeichert.
- Fremde Themen, Datenbank und Serverdateien können nicht über HTTP abgerufen werden.
- Inhalte werden als Text dargestellt; SQLite-Abfragen verwenden Parameter; die Content Security Policy sperrt fremde Skripte.

Die Konten sind **Demo-Konten**, keine verifizierten THM-Identitäten. Für einen echten Hochschulbetrieb sind
THM-SSO mit verifizierter Prüferrolle und Zuordnung, HTTPS (inkl. `COOKIE_SECURE=true`), verwaltete Konten sowie
Datenschutz-, Backup- und Betriebsregeln erforderlich. Das hier ist ein lokaler Lehrprototyp.

Optional vor dem ersten Start `.env.example` als `.env` speichern und Passwörter ändern. Bei einer bestehenden
Datenbank werden Passwörter durch Änderungen der Umgebungsvariablen nicht überschrieben. Der Login-Hinweis im UI
zeigt das Standard-Demopasswort; bei Anpassungen diesen Hinweis in `public/index.html` ändern.

## Dateien

| Datei | Aufgabe |
|---|---|
| `server.js` | Anmeldung, Rechteprüfung, API und SQLite-Speicherung |
| `public/index.html` | Anklickbare Oberfläche und Beispielvorlage |
| `public/app.js` | Aktionen, Dialog und Datenanzeige |
| `public/style.css` | Grün-weißes, responsives Layout |
| `Dockerfile` | Container mit Node.js 24, läuft ohne Root-Rechte |
| `compose.yaml` | Start, lokaler Port und dauerhaftes Daten-Volume |
| `tests/security.test.js` | API-Tests mit separaten, temporären Testdaten |

## Tests

Im laufenden Container:

```bash
docker compose exec pruefer node --test tests/security.test.js
```

Die Tests starten einen separaten Server auf Port 8091. Eure Demo-Daten bleiben unberührt.
Alternativ mit lokalem Node.js 24: `node --test tests/security.test.js`.

## Demo vollständig zurücksetzen

**Achtung: Dieser Befehl löscht alle gespeicherten Demo-Bewertungen und Entscheidungen:**

```bash
docker compose down -v
docker compose up --build
```

## Bezug zur User Story

Als Erstprüfer:in möchte ich das vorgeschlagene Thema einsehen, fachlich bewerten und bestätigen oder ablehnen,
damit die Themenentscheidung nachvollziehbar dokumentiert ist.

User: Erstprüfer:in · Interface: Themenübersicht und Bewertungsformular · Action: lesen, bewerten, entscheiden ·
Data: Themenvorschlag, Bewertung, Status und Verlauf · Control: Rolle, Zuordnung und Freigaberegeln ·
Environment: eigenständige Browseranwendung im lokalen Docker-Container · Quality: nachvollziehbar, persistent und zugriffsgeschützt.

Technische Referenzen: [Docker Compose](https://docs.docker.com/reference/cli/docker/compose/up/),
[Node.js 24 SQLite](https://nodejs.org/download/release/latest-v24.x/docs/api/sqlite.html).
