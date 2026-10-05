# Server-Scheduling: Outlook verschickt nicht mehr selbst

Arbeitsprotokoll fuer den Branch `server-side-scheduling`. Ziel: Outlook soll
sich verhalten wie Thunderbird - nie selbst iTIP verschicken, immer den
CalDAV-Server (sabre/dav, RFC 6638) verschicken lassen. Anlass: bei
Besprechungen, die Outlook selbst organisiert, ist *Accept Proposal* bei einem
Gegenvorschlag von Google ausgegraut; bei Terminen aus Thunderbird nicht.

Dieser Branch liegt im Original-Repo und wird **nicht gepusht**. Spaeter
eigenes oeffentliches Repo (AGPL-3.0).

## Tests und Ergebnisse

| Datum | Test | Ergebnis |
| --- | --- | --- |
| 2026-10-04 | `Set SCHEDULE-AGENT=CLIENT` aus | Server und Outlook verschicken, Einladung doppelt |
| 2026-10-04 | SOGo-Profil | verworfen, setzt `X-SOGO-SEND-APPOINTMENT-NOTIFICATIONS: NO` |
| 2026-10-05 | *Use GlobalAppointmentID for UID attribute*, neue Besprechung | *Accept Proposal* ausgegraut. UID ist nicht die Ursache |
| 2026-10-05 | Profil deaktiviert, neue Besprechung | *Accept Proposal* ausgegraut. Das Plugin ist nicht die Ursache, es liegt an dem, was Outlook verschickt |

Hinweis: mit *Use GlobalAppointmentID* heissen neue Dateien auf dem Server
`040000008200E000...ics` (Grossbuchstaben). Ein Suchmuster wie
`[0-9a-f-]+\.ics` findet sie nicht; `[^/<>]+\.ics` verwenden.

## Prototyp: Versand abfangen

`CalDavSynchronizer/Implementation/Events/ServerSchedulingSendGuard.cs`,
eingehaengt in `ComponentContainer.cs`.

Faengt `Application.ItemSend` fuer Besprechungsanfragen und Absagen ab, die man
als Organisator verschickt, in Ordnern eines Profils, das aktiv ist, Teilnehmer
abbildet und weder `SCHEDULE-AGENT=CLIENT` noch die SOGo-Eigenschaft setzt.

- Einladung/Aenderung: Versand abgebrochen, Termin als verschickt markiert
  (`PidLidFInvited`), gespeichert, Fenster geschlossen. Das Plugin laedt hoch,
  der Server verschickt.
- Ganze Besprechung abgesagt: Versand abgebrochen, Termin geloescht. Das Plugin
  loescht auf dem Server, der Server schickt die Absage.
- Bei jedem Fehler verschickt Outlook wie bisher.

Profileinstellungen fuer den Test:

- *Set SCHEDULE-AGENT=CLIENT*: aus
- *Use GlobalAppointmentID for UID attribute*: an (sonst ordnet Outlook den
  Gegenvorschlag nicht zu, weil die Einladung die UID des Servers traegt)
- *Synchronize items immediately after change*: an
- Profil aktiv

Log: `%LOCALAPPDATA%\CalDavSynchronizer\log.txt`, Eintraege mit `ItemSend:`.

## Build-Umgebung (Windows-VM unter Proxmox)

VM: 4 vCPUs (CPU-Typ `host`), 8-16 GB RAM, 40 GB frei. Outlook in derselben VM.

Visual Studio 2022 Community mit Office/SharePoint-Workload (Admin-PowerShell):

```powershell
winget install --id Microsoft.VisualStudio.2022.Community -e --override "--add Microsoft.VisualStudio.Workload.Office --includeRecommended --passive --wait"
```

Falls `git --version` nichts liefert: `winget install --id Git.Git -e`.

## Stand von Linux nach Windows bringen

Linux schreibt ein Git-Bundle auf das Share `omv.local/public`:

```bash
cd ~/git/outlookcaldavsynchronizer
git bundle create "/run/user/1000/gvfs/smb-share:server=omv.local,share=public/ocs.bundle" server-side-scheduling
```

Windows, einmalig:

```powershell
git clone -b server-side-scheduling \\omv.local\public\ocs.bundle C:\src\ocs
```

Windows, danach bei jedem neuen Stand:

```powershell
cd C:\src\ocs
git pull
```

Lokal unter `C:\` bauen, nicht direkt vom Share (VSTO-Vertrauensstellung).

## Bauen und starten

1. Das installierte CalDav Synchronizer unter *Apps & Features*
   deinstallieren. Die Profile in `%APPDATA%\CalDavSynchronizer` bleiben.
2. `C:\src\ocs\CalDavSynchronizer.sln` in Visual Studio oeffnen.
3. Projekt *CalDavSynchronizer* -> Eigenschaften -> *Signierung* ->
   *Testzertifikat erstellen* (das Zertifikat des Originalautors fehlt).
   Diese Aenderung nicht committen.
4. Outlook schliessen, *CalDavSynchronizer* als Startprojekt, F5.
