# Nextcloud Talk Direkt – Entwicklerdokumentation

Diese Variante integriert einen Direkt-Button in den Thunderbird-Termin-Dialog,
der ohne Zwischenschritt eine neue öffentliche Nextcloud-Talk-Unterhaltung anlegt.

## Verzeichnisstruktur

```
Nextcloud Talk 2.0.0/
├── background.js         # Hintergrundlogik: API-Aufrufe & Utilities
├── experiments/
│   └── calToolbar/
│       ├── parent.js     # Frontendcode, der den Button und den Dialog steuert
│       └── schema.json   # Experiment-API-Definition
├── icons/                # Add-on- und Spenden-Icons
├── manifest.json         # Add-on-Metadaten (Manifest v2)
├── options.html/.js      # Einstellungsseite für URL/Benutzer/App-Passwort
└── README.md             # Diese Datei
```

## Abläufe

1. **Toolbar-Injektion** (`parent.js::inject`): Der Button wird in den Termin-Dialog eingefügt.
   Ein Klick ruft `openCreateDialog` auf.
2. **Dialog** (`openCreateDialog`): Fragt Titel/Passwort/Lobby/Moderator ab, führt API-Aufrufe
   über das Hintergrundskript aus und fügt den Talk-Link in den Termin ein.
3. **Hintergrundskript** (`background.js`): Übernimmt sämtliche REST-Aufrufe
   (Talk API sowie CardDAV-Systemadressbuch) und kapselt sie hinter sicheren Helfern.
4. **Lobby-Watcher** (`setupLobbyWatcher`): Beobachtet Terminänderungen und synchronisiert die
   Lobby-Startzeit, sobald der Termin gespeichert wurde.

## Berechtigungen

- `storage`: Speichert URL, Benutzername und App-Passwort in der Add-on-Konfiguration.
- `*://*/ocs/*`: Notwendig für sämtliche Talk-REST-Endpunkte (`/ocs/v2.php/apps/spreed/...`).
- `*://*/remote.php/*`: Wird ausschließlich verwendet, um das vom Server bereitgestellte
  Systemadressbuch (`remote.php/dav/...`) auszulesen. Der entsprechende Code ist in
  `background.js` dokumentiert.

Weitere Host-Permissions sind nicht erforderlich; WebDAV/CalDAV-Anfragen laufen über die
Thunderbird-eigene Infrastruktur.

## Sicherheit / Hardening

- Alle Netzwerkfehler werden abgefangen und führen zu klaren UI-Meldungen.
- 403-Antworten bei Lobby-Änderungen resultieren in einer deaktivierten Lobby-Checkbox,
  sodass keine wiederholten Fehler auftreten.
- Das Systemadressbuch wird gecacht (`SYSTEM_ADDRESSBOOK_TTL`), um API-Traffic zu minimieren.
- Für Textverarbeitung (Beschreibung, vCard) werden Escapes entfernt und Eingaben normalisiert.

## Tests und Validierung

- Add-on lässt sich in Thunderbird 140.* bis 144.* testen (siehe `strict_min_version`/`max_version`).
- ATN-Hinweise zu Manifest/Icons wurden beseitigt (`browser_specific_settings`, korrekte Icon-Größen).
- Bitte nach jeder Codeänderung das XPI neu packen (`tar -a -cf …`) und per `about:debugging` laden.
