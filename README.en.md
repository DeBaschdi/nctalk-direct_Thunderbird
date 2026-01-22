[Deutsch](README.md) | [English](README.en.md)

# Nextcloud Enterprise for Thunderbird

Nextcloud Enterprise for Thunderbird connects Thunderbird directly with Nextcloud Talk and secure Nextcloud sharing. One click opens a modern wizard, creates Talk rooms with lobby and moderator delegation, and inserts the meeting link (including password) into the event. From the compose window, you can generate a Nextcloud share with upload folder, expiration date, password, and personal message. No copy-paste juggling and no open links in emails: everything stays in Thunderbird and is stored cleanly in your Nextcloud.

## Highlights

- **One-click Nextcloud Talk**
  Open an event, choose Nextcloud Talk, configure the room, and define a moderator. The wizard writes title/location/description (including help link) into the event.
- **Sharing deluxe**
  The "Add Nextcloud Share" button starts the sharing assistant with upload queue, password generator, expiration date, and note field. The finished share is inserted as formatted HTML into the email.
- **Enterprise security**
  Lobby until start time, moderator delegation, automatic cleanup of unsaved events, required passwords, and expiration policies protect sensitive meetings and files.
- **Seamless Nextcloud integration**
  Login Flow V2, automatic room tracking, and debug logs in [NCBG], [NCUI], [NCSHARE] help with troubleshooting.
- **ESR-ready**
  Optimized and tested for Thunderbird ESR 140.X with a minimal experiment footprint.

## Whats new in 2.2.2?

- New branding: Nextcloud Enterprise for Thunderbird.
- Sharing wizard (renamed, explicitly separate from cloudFile).
- Optional host permissions per instance + Talk dialog i18n fix.
- New default options for Sharing and Talk (share name, permissions, password/expiry days, Talk title/lobby/listing/room type).
- Advanced language overrides for the sharing HTML block and Talk text block (default/en/de/fr).
- Note: The add-on ID has changed. Please uninstall the previously installed version before installing 2.2.2.

## Feature overview

### Nextcloud Talk directly from the event
- Talk popup with lobby, password, listable option, room type, and moderator search.
- Automatic insertion of title, location, and description (including help link and password) into the event.
- Room tracking, lobby updates, delegation workflow, and cleanup if the event is discarded or moved.
- Calendar changes (drag-and-drop or dialog edits) keep lobby/start time in sync on the server.

### Nextcloud Sharing in the compose window
- Four steps (share, expiration date, files, note) with a password-protected upload folder.
- Upload queue with duplicate checks, progress display, and optional share without upload.
- Automatic HTML blocks with link, password, expiration date, and optional note.

### Administration & compliance
- Login Flow V2 (app password is created automatically) and central options (base URL, debug mode, sharing paths, default values for Sharing/Talk).
- Full internationalization (DE/EN/FR) and structured debug logs for support cases.

## System requirements
- Thunderbird ESR 140.X (Windows/macOS/Linux)
- Nextcloud with Talk & Sharing (DAV) enabled
- App password or Login Flow V2

## Installation
1. Install the current XPI `nextcloud-enterprise-2.2.2.xpi` in Thunderbird (Add-ons > Gear > Install Add-on From File).
2. Restart Thunderbird.
3. In the add-on options, enter base URL, user, and app password or start the login flow.

## Support & feedback
- **Troubleshooting:** Enable debug mode in the options; relevant logs appear as [NCUI][Talk], [NCUI][Sharing], [NCBG], [NCSHARE], [NCExp].

Good luck with secure, professional work using Nextcloud Enterprise for Thunderbird!

## Screenshots

### Settings menu
![Settings menu](screenshots/Settings.png)

### Talk wizard
![Talk wizard](screenshots/talk_wizzard.png)

### Sharing wizard
![Sharing wizard step 1](screenshots/filelink_wizzard1.png)
![Sharing wizard step 2](screenshots/filelink_wizzard2.png)
![Sharing wizard step 3](screenshots/filelink_wizzard3.png)
![Sharing wizard step 4](screenshots/filelink_wizzard4.png)
![Sharing wizard step 5](screenshots/filelink_wizzard5.png)
