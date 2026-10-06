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
| 2026-10-05 | Prototyp 1, neue Besprechung `zzz` | Outlook verschickt nichts. Danach Rueckfrage *Save changes and send update* (halber Zustand). Ohne Teilsynchronisation noch nicht auf dem Server, keine Einladung bei Gmail |
| 2026-10-05 | Prototyp 1, Absage alter Besprechungen | Versand unterdrueckt, Outlook loescht den Termin trotzdem selbst (`Delete()` danach: *item has been moved or deleted*). Absage erreicht niemanden: Outlook schweigt, und der Server schickt bei `SCHEDULE-AGENT=CLIENT` kein CANCEL |
| 2026-10-05 | Prototyp 2, neue Besprechung `yyy` | **Ziel erreicht.** Teilsynchronisation nach 10 s, Server verschickt, Einladung bei Gmail, Gegenvorschlag zurueck, *Accept Proposal* **aktiv**. Auf dem Server weder `SCHEDULE-AGENT=CLIENT` noch Organisator als `ATTENDEE` - der Organisator als Teilnehmer ist also nicht noetig, entscheidend ist, wer verschickt |
| 2026-10-05 | Prototyp 2, Absage `zzz` aus dem geoeffneten Termin | Versand dreimal unterdrueckt, Termin blieb stehen: aus dem Fenster heraus loescht Outlook nicht selbst (anders als aus der Kalenderansicht). Prototyp 3 loescht dann selbst |
| 2026-10-05 | Prototyp 3, Absage `yyy` aus dem geoeffneten Termin | Plugin: Outlook loescht selbst, `DELETE` auf dem Server ok. Absage kam bei Google nicht an - Fehler lag im sabre/dav-Dienst (email-amazon), dort behoben |

| 2026-10-06 | *Accept Proposal* mit *Send Update* | ok, ein Update bei Gmail |
| 2026-10-06 | Absage aus der Kalenderansicht | ok, eine Absage bei Gmail |
| 2026-10-06 | Zwei Teilnehmer, einer entfernt | ok, der Entfernte bekommt die Absage vom Server |
| 2026-10-06 | Neue Besprechung senden | Rueckfrage *Save changes and send update* erscheint weiterhin. Offen: kommt sie von Outlook nach dem abgebrochenen Versand oder von `Inspector.Close`? Logzeilen `closing inspector` / `inspector closed` eingebaut |
| 2026-10-06 | Dasselbe mit Logzeilen | `Close` kehrt nach 0,5 s zurueck, die Rueckfrage kommt danach. Ursache vermutlich `PidLidFInvited` vor dem Schliessen: Outlook haelt die Besprechung fuer verschickt und den Fensterinhalt fuer eine Aenderung. Jetzt erst nach dem Schliessen gesetzt |
| 2026-10-06 | Neue Besprechung senden | Fenster geht ohne Rueckfrage zu, Einladung bei Gmail, Antwort kommt zurueck |
| 2026-10-06 | Zeit aendern, *Send Update* | Rueckfrage erscheint (gleiche Ursache, stoert nicht, kein doppelter Versand). Update und Annahme ok |
| 2026-10-06 | Einladung an interne Adresse | Keine Mail: sabre/dav liefert lokal in den Kalender des Empfaengers. Empfehlung fuer den Server: interne Empfaenger wie externe nur per iMIP einladen (sonst droht in Outlook ein doppelter Termin, wenn Mail und Synchronisation konkurrieren) |

Hinweis: mit *Use GlobalAppointmentID* heissen neue Dateien auf dem Server
`040000008200E000...ics` (Grossbuchstaben). Ein Suchmuster wie
`[0-9a-f-]+\.ics` findet sie nicht; `[^/<>]+\.ics` verwenden.

## Prototyp: Versand abfangen

`CalDavSynchronizer/Implementation/Events/ServerSchedulingSendGuard.cs`,
eingehaengt in `ComponentContainer.cs`.

Faengt `Application.ItemSend` fuer Besprechungsanfragen und Absagen ab, die man
als Organisator verschickt, in Ordnern eines Profils, das aktiv ist, Teilnehmer
abbildet und weder `SCHEDULE-AGENT=CLIENT` noch die SOGo-Eigenschaft setzt.

Uebernommen werden nur **neue** Besprechungen (`PidLidFInvited` nicht gesetzt)
und solche, die schon dem Server uebergeben wurden (Markierung
`CalDavSynchronizerServerScheduled`). Was Outlook frueher selbst verschickt hat,
bleibt bei Outlook, und der Mapper setzt dafuer weiter `SCHEDULE-AGENT=CLIENT`
- unabhaengig vom Haken im Profil. Sonst kaemen Aenderungen doppelt an.
Dasselbe gilt fuer empfangene Einladungen: Antworten verschickt Outlook.

- Einladung/Aenderung: noch innerhalb von `ItemSend` als verschickt
  (`PidLidFInvited`) und als Server-Termin markiert und gespeichert, dann
  Versand abgebrochen und Fenster ohne Rueckfrage geschlossen. Das Plugin laedt
  hoch, der Server verschickt.
- Ganze Besprechung abgesagt: Versand abgebrochen, Fenster verworfen, Termin
  geloescht (falls Outlook das nicht schon selbst getan hat). Das Plugin loescht
  ihn auf dem Server, der Server schickt die Absage.
- Schlaegt das Speichern fehl, verschickt Outlook wie bisher.

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
2. `C:\src\ocs\CalDavSynchronizer.sln` in Visual Studio oeffnen. Die Meldung,
   `CalDavSynchronizer.Setup.vdproj` werde nicht unterstuetzt, mit *OK*
   bestaetigen. Das ist nur das MSI-Installer-Projekt (braucht die Erweiterung
   *Microsoft Visual Studio Installer Projects*), fuer F5 nicht noetig.
3. Projekt *CalDavSynchronizer* -> Eigenschaften -> *Signierung* ->
   *Testzertifikat erstellen* (das Zertifikat des Originalautors fehlt).
   Diese Aenderung nicht committen.
4. Outlook schliessen, *CalDavSynchronizer* als Startprojekt, F5.

## Automatische Einrichtung

`CalDavSynchronizer/AutoSetup/AccountAutoSetup.cs`, laeuft auf Knopfdruck:
Menueband *CalDav Synchronizer* -> **Set up calendars**
(`ComponentContainer.SetUpAccountsAsync`). Waehrenddessen ein Wartefenster,
danach eine Zusammenfassung (gefunden / hinzugefuegt / reaktiviert /
deaktiviert). Kein Abgleich danach (verlangsamt die Einrichtung); die Profile
gleichen nach ihrem Intervall oder per *Synchronize now* ab.

Zuerst lief sie bei jedem Start. Am 2026-10-06 hat Outlook das Add-in
abgeschaltet (*caused Outlook to start slowly*, 1,3 s). Die Suche lief zwar
erst nach dem Laden in einem Timer, vermutlich lag es am Debug-Build; ein
Knopf vermeidet aber unerwartete Netzzugriffe beim Start. Wieder einschalten:
*Slow and Disabled Add-ins* -> *Options* -> *Always enable this add-in*.

- Fuer jedes IMAP-Konto in Outlook: Server `https://dav.{domain}/`,
  Benutzer = Mailadresse, Passwort = gespeichertes IMAP-Passwort.
- Sucht ueber `/.well-known/caldav` und `/.well-known/carddav` alle Kalender
  und Adressbuecher (eigene und geteilte).
- Eigener Kalender `default` -> Standard-Kalenderordner des IMAP-Kontos
  (Outlook sucht Besprechungsantworten nur dort). Eigenes Adressbuch `default`
  -> Standard-Kontaktordner. Alle anderen -> Unterordner davon, geteilte mit
  Besitzer im Namen, z.B. `Buchhaltung (accounting@...)`.
- Schreibrecht -> Abgleich in beide Richtungen; nur Leserecht -> nur
  Server nach Outlook. Sofortabgleich nach Aenderung an.
- Kalender: `SCHEDULE-AGENT=CLIENT` aus, *Use GlobalAppointmentID* an.
- Aufgaben (VTODO) werden nicht eingerichtet.
- Profile erkennt sie an der URL wieder. Von Hand angelegte Profile mit
  derselben URL bleiben unangetastet (keine Dubletten). Verschwindet eine
  Sammlung vom Server, wird ihr Profil deaktiviert, der Ordner bleibt.
  Findet sie gar nichts (Server nicht erreichbar, Anmeldung abgelehnt), aendert
  sie nichts.
- Eigene Markierung `IsAutoConfigured` in den Profilen.

Der Versandschutz greift nur noch bei Profilen, die Outlook-Aenderungen
hochladen; aus einem nur-lesend geteilten Kalender verschickt also weiter
Outlook.

Geteilte Kalender: sabre/dav meldet als `DAV:owner` den Empfaenger der
Freigabe, nicht den Besitzer (Test 2026-10-06: Ordner hiess nur `Calendar`).
Der Besitzer kommt jetzt aus `DAV:invite` (bzw. `CS:invite`), dort steht er als
*organizer* (bzw. in `DAV:invite` als `sharee` mit `shared-owner`), per curl am 2026-10-06 bestaetigt. Ordnername = Teil der Besitzer-Adresse vor dem `@`, bei eigenem
Kalendernamen `besitzer - Name`. Punkte im Namen werden zu Leerzeichen: mit
`accounting@bayarea-cc.com` meldete Outlook *Cannot create the folder*
(vermutlich der IMAP-Hierarchietrenner `.`). Scheitert ein Ordner, wird nur
diese Sammlung uebersprungen und gemeldet.

Einschraenkung: Einladungen funktionieren sauber nur aus dem
Standardkalender.

Test 2026-10-06: zwei Kalender und ein Adressbuch gefunden, eigener Kalender
und Adressbuch auf die Standardordner, geteilter Kalender als Unterordner,
nur lesend. Add-in trotz Start ohne Debugger erneut wegen langsamen Starts
abgeschaltet - VSTO-Ladezeit, nicht unser Code (der laeuft erst danach).
Abhilfe auf der VM:
`reg add HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList /v CalDavSynchronizer /t REG_DWORD /d 1 /f`

## Weitere Anpassungen

- **Update-Pruefung abgeschaltet** (`ComponentContainer`): sie fragte das
  Originalprojekt und haette Kunden dessen Version ohne Server-Scheduling
  angeboten. Spaeter gegen den eigenen Server neu bauen.
- **About ausgeblendet** (Menueband und Tray-Menue), Links und Update-Knopf
  zeigen aufs Original. Beim Neubau an AGPL denken: Hinweis auf die
  Originalautoren behalten und Link auf unseren Quelltext anbieten.
- **Fortschritt bei *Synchronize now***: das vorhandene Fortschrittsfenster
  erscheint bei manuellem Abgleich immer (Schwelle 0), beim automatischen
  Abgleich weiter erst ab der Schwelle aus *General Options* (333).

## Naechste Schritte

1. **Abspecken auf sabre/dav.** Automatische Einrichtung (siehe oben) ist
   gebaut. Danach: alle anderen Anbieterprofile, OAuth (Google, Swisscom),
   Aufgaben und die alten Einrichtungsdialoge entfernen.
2. **Installer** mit dem vorhandenen Setup-Projekt
   (`CalDavSynchronizer.Setup.vdproj`, braucht die VS-Erweiterung *Microsoft
   Visual Studio Installer Projects*). Zertifikat vorerst selbst erstellt; der
   Installer traegt das Add-in in die VSTO-Vertrauensliste ein, damit Outlook
   nicht nachfragt. Gekauftes Code-Signing-Zertifikat spaeter (Kandidat:
   Certum Open-Source-Zertifikat; Let's Encrypt stellt keine aus). Er soll das
   Add-in in Outlooks Liste *DoNotDisableAddinList* eintragen
   (`HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList`,
   Wert `CalDavSynchronizer` = DWORD 1), damit Outlook es nicht wegen
   langsamen Starts abschaltet.
