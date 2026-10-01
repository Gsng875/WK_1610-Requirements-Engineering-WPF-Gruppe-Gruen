# ThesisFlow Prototypen (Docker)

Statische Webseiten der Gruppe Grün, ausgeliefert über nginx.

## Starten

    docker compose up -d --build

Danach im Browser öffnen: http://localhost:8080

## Stoppen

    docker compose down

## Inhalt (`site/`)

| Datei | Inhalt |
|---|---|
| `index.html` | Startseite mit Links |
| `praesentation.html` | D2D-Präsentation Verlängerungsantrag (11 Folien, mit Smartphone-Prototyp) |
| `verlaengerung-smartphone.html` | Verlängerungsantrag, Smartphone-Prototyp |
| `verlaengerung-pc.html` | Verlängerungsantrag, PC-Version |
| `themenfreigabe-smartphone.html` | Themenfreigabe, Smartphone-Prototyp |
| `themenfreigabe-pc.html` | Themenfreigabe, PC-Version |

Alle Angaben sind Beispieldaten, es wird nichts gespeichert. Die Seiten laden Schriften von Google Fonts; ohne Internet greifen Ersatzschriften.

## Port ändern

In `docker-compose.yml` bei `ports` den linken Wert ändern, zum Beispiel `"9090:80"`.
