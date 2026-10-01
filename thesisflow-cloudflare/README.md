# ThesisFlow für Cloudflare Workers + D1

Eigenständige, anklickbare Prüfer-Anwendung. Oberfläche, API und Datenbank laufen bei Cloudflare;
Docker wird für diese Variante nicht benötigt. Keine echte THM-Anbindung, alle Beispieldaten sind fiktiv.

## Voraussetzungen

- Node.js 22 oder neuer mit npm
- Cloudflare-Konto
- ZIP entpacken, den Ordner **thesisflow-cloudflare** in VS Code öffnen.
- Alle folgenden Befehle im VS-Code-Terminal **in diesem Ordner** ausführen.

## 1. Installieren und anmelden

```bash
npm install
npx wrangler login
```

Die Anmeldung öffnet den Browser. Mit eurem eigenen Cloudflare-Konto anmelden.

## 2. D1-Datenbank anlegen

```bash
npx wrangler d1 create thesisflow
```

Cloudflare gibt eine **database_id** zurück. In `wrangler.jsonc` diese Zeile ersetzen:

```json
"database_id": "00000000-0000-0000-0000-000000000000"
```

Die echte ID eintragen, die Anführungszeichen beibehalten. `database_name` bleibt `thesisflow`;
`binding` bleibt `DB`. Wenn die Datenbank in eurem Konto bereits existiert, ihre ID aus dem Dashboard verwenden.
Die ID ist eine Zuordnung, kein Passwort. In der Datei stehen keine Account-Zugangsdaten.

## 3. Tabellen und Beispieldaten online einrichten

```bash
npm run db:remote
npm run seed:remote
```

Das Seed-Skript erzeugt zufällige Passwörter und speichert nur Passwort-Hashes in D1.
Die Zugangsdaten stehen anschließend in **credentials.local.txt** auf eurem Rechner.
Benutzer `weber` sieht zwei Themen, `klein` nur sein eigenes Thema; `student` erhält keinen Zugriff.
Bestehende Benutzer, Themen und Entscheidungen werden beim erneuten Ausführen nicht überschrieben.
Die private Seed-Datei wird wiederverwendet. Löscht sie nicht, wenn ihr eure erzeugten Zugangsdaten behalten möchtet.

## 4. Veröffentlichen

```bash
npm run deploy
```

Die Ausgabe zeigt eure Adresse, z. B. **https://thesisflow-pruefer.EUER-SUBDOMAIN.workers.dev**.
Diese öffnen und mit dem Zugang aus `credentials.local.txt` anmelden.

**Bei Cloudflare ist das ein Worker unter „Workers & Pages“.** Nicht nur `public/` als Pages-Upload
hochladen: Damit würden API und D1-Anmeldung fehlen. Für die vorhandene Variante gibt es keinen Build-Schritt
für die Oberfläche. `npm run deploy` lädt Worker und Static Assets gemeinsam hoch.

## Lokal testen

```bash
npm run db:local
npm run seed:local
npm run dev
```

Dann **http://localhost:8787** öffnen. Lokale und entfernte Datenbank sind getrennt.
Mit derselben privaten Seed-Datei werden dieselben Demo-Zugangsdaten in beiden Datenbanken eingerichtet.
`Strg+C` beendet den lokalen Server. Browser-Anfragen verwenden automatisch die aktuelle Origin.

## Weitere Änderungen deployen

Dateien bearbeiten, dann erneut **`npm run deploy`** ausführen. Erneutes Deployen löscht keine D1-Daten.
Eine benutzerdefinierte Domain kann später im Cloudflare-Dashboard mit dem Worker verbunden werden;
die Anmeldung nutzt die Domain der jeweiligen Anfrage.

## Dateien

| Datei | Aufgabe |
|---|---|
| `public/` | Bestehendes anklickbares Interface |
| `src/worker.js` | API, Rollen, Zuordnung, Bewertung und Entscheidung |
| `src/crypto.js` | Zufallstoken und Web-Crypto-Passwort-Hashes |
| `schema.sql` | Tabellen, Indizes und Datenbankregeln |
| `wrangler.jsonc` | Worker, Static Assets und D1-Binding |
| `scripts/seed.mjs` | Erzeugt Demo-Konten und Beispielthemen |
| `scripts/check-config.mjs` | Verhindert Deployment mit der Platzhalter-Datenbank-ID |
| `tests/worker.test.mjs` | API- und Sicherheitstests mit Miniflare/echtem lokalen D1 |

## Sicherheits- und Prozessregeln

- Rollen und Themenzuordnung werden serverseitig geprüft; keine frei wählbare Prüferrolle.
- PBKDF2-SHA-256 mit individuellen Salts und 100.000 Iterationen; keine Standardpasswörter in der Oberfläche.
- Zufällige Sitzungs-Cookies: HttpOnly, SameSite=Strict, über HTTPS zusätzlich Secure.
- In D1 liegt nur der Hash des Sitzungs-Tokens; Sitzungen laufen nach 30 Minuten Inaktivität ab.
- D1 speichert die Login-Sperre über alle Worker-Instanzen hinweg. Maximal fünf Login-Versuche pro
  IP innerhalb eines 15-Minuten-Fensters; erfolgreiche Prüfer-Anmeldung setzt den Zähler zurück.
- Änderungsanfragen benötigen dieselbe Origin und nach Anmeldung ein CSRF-Token.
- Freigabe benötigt vollständige Angaben, hinterlegte Zulassung, „Geeignet“ und beide Prüfkriterien.
- Entscheidung und Verlauf werden atomar gespeichert; parallele Anfragen überschreiben keine endgültige Entscheidung.
- Themen werden nicht als öffentliche HTML-Dateien ausgeliefert. Bewertungen bleiben nach Neuladen und Deployment gespeichert.
- Die Themenfreigabe ist keine offizielle Ausgabe, Abschlussnote oder elektronische Hochschulunterschrift.

Die Demo prüft registrierte Konten, keine echten THM-Identitäten. Für echten Hochschulbetrieb müssen
THM-SSO, verifizierte Rollen/Zuweisungen und die institutionellen Betriebs- und Datenschutzregeln ergänzt werden.
`credentials.local.txt` und `seed.local.sql` sind in `.gitignore`; beide liegen außerhalb von `public/`
und werden nicht als Static Assets veröffentlicht. Verwendet für die Demonstration nur die fiktiven Daten.

## Prüfen

```bash
npm test
npm run check
```

Tests erzeugen ausschließlich temporäre lokale Datenbanken, keine Änderungen an eurem Cloudflare-Konto.
`check` erstellt einen Deployment-Build ohne Veröffentlichung.

## Häufige Probleme

- **database_id fehlt:** ID aus Schritt 2 in `wrangler.jsonc` eintragen.
- **no such table:** `npm run db:remote` bzw. `npm run db:local` ausführen.
- **Benutzer unbekannt:** `npm run seed:remote` bzw. `npm run seed:local` ausführen und die erzeugten Zugangsdaten verwenden.
- **Login-Sperre:** 15-Minuten-Fenster abwarten; nicht wiederholt Passwörter ausprobieren.
- **Port 8787 belegt:** `npx wrangler dev --port 8788` verwenden und die angezeigte Adresse öffnen.

Referenzen: [Static Assets](https://developers.cloudflare.com/workers/static-assets/binding/),
[D1-Batch-Transaktionen](https://developers.cloudflare.com/d1/worker-api/d1-database/),
[Web Crypto](https://developers.cloudflare.com/workers/runtime-apis/web-crypto/).

Prüfstand: Deployment-Dry-Run, lokale Schema-/Seed-Befehle und API-Tests in der Worker-Laufzeit erfolgreich.
Die Browserabläufe Anmeldung, Themenanzeige, Speichern, Freigabe, Neuladen, Abmeldung und Studentensperre
wurden am Desktop sowie die freigegebene Ansicht mobil geprüft. Das Deployment in eurem Cloudflare-Konto erfolgt erst durch euch.
