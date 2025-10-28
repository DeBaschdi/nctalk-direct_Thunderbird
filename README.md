# Nextcloud Talk Direkt - Entwicklerleitfaden (Version 2.1.0)

![Screenshot: Nextcloud Talk Direkt Optionen](screenshot.jpg)

Dieses Thunderbird-Add-on ergaenzt den Termin-Dialog um einen Talk-Button, der in wenigen Schritten einen Nextcloud-Talk-Raum erzeugt und den Teilnahmelink direkt in das Ereignis uebernimmt.

Die Erweiterung kann Raeume als klassische Konversation oder als Termin-Unterhaltung anlegen. 
Letzteres koppelt den Raum an die Start- und Endzeit. 

Falls ein Termin verworfen oder während des Dialoges im Terminplaner geloescht wird, wird der verwaiste Raum automatisch entfernt. 

Optional lassen sich Lobby, Passwortschutz und Moderatoren komfortabel konfigurieren.

Bei der Moderatoren Auswahl wird man selbst nach Uebergabe aus dem Raum entfernt.
Avatare kommen aus dem Nextcloud-Systemadressbuch.

## Verzeichnisstruktur

```text
nc-talk-2.1.0-xpi/
|-- modules/
|   |-- background.js      # Einstiegspunkt, Runtime-Messaging, Avatar-Dekodierung
|   `-- talkcore.js        # Zentrale Nextcloud-API-Helfer (REST, Adressbuch, Cleanup)
|-- experiments/
|   |-- calToolbar/
|   |   |-- parent.js      # Frontend/Experiment, Dialog & Button
|   |   `-- schema.json    # Experiment-API-Beschreibung
|-- icons/                 # Add-on-Icons (verschiedene Groessen)
|-- _locales/              # Uebersetzungen (de/en/fr via browser.i18n)
|   |-- de/messages.json
|   |-- en/messages.json
|   `-- fr/messages.json
|-- manifest.json          # Manifest v2 + browser_specific_settings
|-- options.html           # Einstellungsdialog: Markup fuer die Optionen
|-- options.js             # Einstellungsdialog: Logik & Validierung
|-- LICENSE.txt            # Lizenztext (GNU AGPL v3)
`-- README.md              # Diese Datei
```

## Wichtige Ablaeufe

1. **Toolbar-Injektion** (`parent.js::inject`): fuegt den Button in Termin- und Aufgaben-Dialoge ein. Alle Event-Listener werden zentral ueber `handle(...)` registriert.
2. **Erstellen-Dialog** (`openCreateDialog`): steuert die UI fuer Titel, Passwort, Lobby-Optionen sowie Moderator-Auswahl und delegiert alle Netzwerk-Operationen an das Hintergrundskript.
3. **Hintergrundskript** (`modules/background.js` + `modules/talkcore.js`):
   - Talk-REST-Aufrufe (`/ocs/v2.php/apps/spreed/...`)
   - CardDAV-Lookups (`remote.php/dav/...`) des Nextcloud-Systemadressbuchs
   - Avatar-Verarbeitung (Pixel-Extraktion + Serialisierung fuer das Frontend)
4. **Lobby-Watcher** (`setupLobbyWatcher`): beobachtet Termin-Aenderungen und synchronisiert die Lobby-Startzeit nach dem Speichern. Berechtigungsfehler werden abgefangen und sichtbar gemacht.

## Berechtigungen

- `storage`: persistiert Basis-URL, Benutzername und App-Passwort.
- Host-Permissions `*://*/ocs/*` und `*://*/remote.php/*`: notwendig fuer Talk- und CardDAV-Aufrufe. Weitere Domains werden nicht angesprochen.
- Keine Content-Skripte; alle Requests laufen ueber `fetch` im Hintergrundskript.

## Sicherheit & Robustheit

- Jeder Netzwerk-Request wird in `try/catch` gefasst und mit aussagekraeftigen Fehlermeldungen quittiert. 403-Antworten koennen z.B. die Lobby-Option deaktivieren.
- Systemadressbuch-Ergebnisse werden fuer 5 Minuten gecacht (`SYSTEM_ADDRESSBOOK_TTL`). Danach wird neu beim Server angefragt.
- Avatar-Daten werden nur als Plain-Array weitergereicht (kein ArrayBuffer erforderlich), womit Structured-Clone-Probleme in Thunderbird vermieden werden.
- Entfernte Termine hinterlassen keine Talk-Reste: Wird der Dialog ohne Speichern geschlossen oder ein Termin geloescht, entfernt das Add-on den zuvor erzeugten Talk-Raum automatisch.

## Aenderungen in Version 2.1.0

- **Modularisierung**: Hintergrundskript und Talk-Core liegen jetzt im Ordner `modules/` und werden direkt ueber das Manifest geladen.
- **Event-Unterhaltungen**: Event-Raeume sind nicht mehr auf Moderatoren beschraenkt; alle eingeladenen Teilnehmer koennen der Unterhaltung beitreten.
- **Dokumentation**: Funktionen erhielten aussagekraeftige Block-Kommentare.

## Aenderungen in Version 2.0.1

- **Event-Konversationen**: Optionaler Modus koppelt den Talk-Raum an den Termin (`objectType:event`, `objectId`). Bei fehlender Server-Unterstuetzung faellt das Add-on automatisch auf den bisherigen Standard-Raum zurueck.
- **Avatar-Pipeline**: Hintergrund liefert Pixelarrays ohne TypedArray-Abhaengigkeiten; Frontend zeichnet Avatare robust inklusive Fallback auf Initialen.
- **Moderator-Vorschau**: Ausgewaehlter Moderator erscheint unterhalb des Eingabefelds mit Avatar (oder Initialen) und Name/E-Mail.
- **Dialog-Refit & Debug-Schalter**: Aufgeraeumte UI, erweiterte Statusmeldungen sowie zuschaltbares Debug-Logging in den Add-on-Optionen.
- **Debug-Initialisierung**: Experiment referenziert browser.storage erst nach API-Kontext; verhindert Startfehler "browser is not defined" im Experiment.
- **Lobby-Setup**: Lobby-Endpoint nutzt wieder `PUT` (Nextcloud Webinar API); Standard-Raeume mit Lobby erzeugen keinen 405-Fallback mehr.
- **Beschreibung**: Event-Konversationen blockieren `PUT /description` (HTTP 400); Update erfolgt nur fuer Standard-Raeume.
- **Capabilities**: Fallback auf `/ocs/v2.php/cloud/capabilities` inklusive Versionspruefung (<32 -> Event-Modus deaktiviert). Debug-Logs zeigen, welche Quelle (Talk/Cloud) den Support liefert.
- **Mehrsprachige UI**: Dialoge und Optionen in Deutsch/Englisch/Franzoesisch via browser.i18n.
- **Erweitertes Debug-Logging**: Hintergrund (`[NCBG]`) protokolliert alle Kernaktionen inklusive gekuerzter Token/IDs; das Experiment-Frontend (`[NCExp]`) loggt Dialog-Workflows, Listenabrufe und Pending-Delegationen.
- **Verbindungs-Test in den Optionen**: Button "Verbindung testen" prueft Basis-URL, Benutzername und App-Passwort direkt per `/ocs/v2.php/cloud/capabilities` und meldet Erfolg bzw. Fehler mit klarer Statusanzeige.
- **Event-IDs & Cleanup**: Event-Konversationen erhalten `start#stop` als `objectId`; verworfene Termin-Dialoge loeschen den erstellten Raum automatisch.










