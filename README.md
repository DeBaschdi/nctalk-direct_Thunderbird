[Deutsch](README.md) | [English](README.en.md)

# Nextcloud Enterprise for Thunderbird 

Nextcloud Enterprise for Thunderbird verbindet Ihr Thunderbird direkt mit Nextcloud Talk und der sicheren Nextcloud-Freigabe. Ein einziger Klick öffnet einen modernen Wizard, erstellt automatisch Talk-Räume inklusive Lobby und Moderatoren Delegation und fügt den Meeting-Link mitsamt Passwort sauber in den Termin ein. Aus dem Verfassen-Fenster heraus erzeugen Sie auf Wunsch sofort eine Nextcloud-Freigabe inklusive Upload-Ordner, Ablaufdatum, Passwort und personalisierter Nachricht. Keine Copy-&-Paste-Orgien mehr, keine offenen Links in Mails: alles läuft in Thunderbird, alles wird sauber in Ihrer Nextcloud abgelegt.

## Highlights

- **Ein Klick zu Nextcloud Talk** 
Termin öffnen, Nextcloud Talk wählen, Raum konfigurieren, Moderator definieren. Der Wizard schreibt Titel/Ort/Beschreibung inklusive Hilfe-Link automatisch in den Termin.
- **Sharing deluxe** 
Compose-Button Nextcloud Freigabe hinzufügen startet den Freigabe-Assistenten mit Upload-Queue, Passwortgenerator, Ablaufdatum und Notizfeld. Die fertige Freigabe landet als formatiertes HTML direkt in der E-Mail.
- **Enterprise-Sicherheit** 
Lobby bis Startzeit, Moderator-Delegation, automatisches Aufräumen nicht gespeicherter Termine, Pflicht-Passwörter und Ablauffristen schützen sensible Meetings und Dateien.
- **Nahtlose Nextcloud-Integration** 
Login-Flow V2, automatische Raumverfolgung sowie Debug-Logs in [NCBG], [NCUI], [NCSHARE] helfen beim Troubleshooting.
- **ESR-ready** 
Optimiert und getestet für Thunderbird ESR 140.X mit minimalem Experiment-Anteil.

## Was ist neu in 2.2.3?

- Neues Branding Nextcloud Enterprise for Thunderbird.
- Freigabe-Wizard (Umbenennung und klare Abgrenzung zu cloudFile).
- Optionale Host-Permissions pro Instanz + i18n Fix im Talk-Dialog.
- Neue Default-Optionen fuer Freigabe und Talk (Share-Name, Rechte, Passwort/Ablauf-Tage, Talk-Titel/Lobby/Listbarkeit/Raumtyp).
- Erweiterte Spracheinstellungen fuer Freigabe-HTML und Talk-Textblock (default/en/de/fr).
- Hinweis: Die Add-on-ID wurde geaendert. Bitte Versionen <=2.2.1 deinstallieren, bevor 2.2.3 installiert wird.

## Funktionsüberblick

### Nextcloud Talk direkt aus dem Termin
- Talk-Popup mit Lobby, Passwort, Listbarkeit, Raumtyp und Moderatorensuche.
- Automatische Einträge von Titel, Ort, Beschreibung (inkl. Hilfe-Link und Passwort) in das Terminfenster.
- Room-Tracking, Lobby-Updates, Delegations-Workflow und Cleanup, falls der Termin verworfen oder verschoben wird.
- Kalender-Aenderungen (Drag-and-drop oder Dialog-Edit) halten Lobby/Startzeit des Talk-Raums synchron.

### Nextcloud Sharing im Compose-Fenster
- Vier Schritte (Freigabe, Ablaufdatum, Dateien, Notiz) mit passwortgeschütztem Upload-Ordner.
- Upload-Queue mit Duplikatprüfung, Fortschrittsanzeige und optionaler Freigabe.
- Automatische HTML-Bausteine mit Link, Passwort, Ablaufdatum und optionaler Notiz.

### Administration & Compliance
- Login Flow V2 (App-Passwort wird automatisch angelegt) und zentrale Optionen (Basis-URL, Debug-Modus, Freigabe-Pfade, Defaultwerte fuer Freigabe/Talk).
- Vollständige Internationalisierung (DE/EN/FR) und strukturierte Debug-Logs für Support-Fälle.

## Systemvoraussetzungen
- Thunderbird ESR 140.X (Windows/macOS/Linux)
- Nextcloud mit Talk & Freigabe (DAV) aktiviert
- App-Passwort oder Login Flow V2

## Installation
1. Aktuelle XPI 
nextcloud-enterprise-2.2.3.xpi in Thunderbird installieren (Add-ons ? Zahnrad ? Add-on aus Datei installieren).
2. Thunderbird neu starten.
3. In den Add-on-Optionen Basis-URL, Benutzer und App-Passwort hinterlegen oder den Login Flow starten.

## Support & Feedback
- **Fehleranalyse:** Debug-Modus in den Optionen aktivieren; relevante Logs erscheinen als [NCUI][Talk], [NCUI][Sharing], [NCBG], [NCSHARE], [NCExp].

Viel Erfolg beim sicheren, professionellen Arbeiten mit Nextcloud Enterprise for Thunderbird!

## Screenshots

### Settings-Menü
![Settings-Menü](screenshots/Settings.png)

### Talk Wizard
![Talk Wizard](screenshots/talk_wizzard.png)

### Sharing Wizard
![Sharing Wizard Schritt 1](screenshots/filelink_wizzard1.png)
![Sharing Wizard Schritt 2](screenshots/filelink_wizzard2.png)
![Sharing Wizard Schritt 3](screenshots/filelink_wizzard3.png)
![Sharing Wizard Schritt 4](screenshots/filelink_wizzard4.png)
![Sharing Wizard Schritt 5](screenshots/filelink_wizzard5.png)
