# Development notes

How the Bay Area Edition differs from the original CalDavSynchronizer, why,
and how to build and release it. See the [README](../README.md) for an overview.

## 1. Background

Most users work with Outlook. With the original add-in, a meeting that Outlook
organizes has *Accept Proposal* greyed out when an attendee (for example on
Google Calendar) proposes a new time. The same meeting created in Thunderbird
works. Side by side, same attendees, same counter proposal:

| Meeting created in | Who sends the iTIP mail | *Accept Proposal* |
| --- | --- | --- |
| Thunderbird | the CalDAV server (sabre/dav, RFC 6638) | works |
| Outlook | Outlook itself | greyed out |

The calendar data differed accordingly:

    # Thunderbird
    ORGANIZER:mailto:organizer@example.com
    ATTENDEE:mailto:organizer@example.com
    ATTENDEE;SCHEDULE-STATUS=1.1:mailto:attendee@gmail.com

    # Outlook (original add-in)
    ORGANIZER;SCHEDULE-AGENT=CLIENT:mailto:organizer@example.com
    ATTENDEE;SCHEDULE-AGENT=CLIENT:mailto:attendee@gmail.com

**Goal: Outlook behaves like Thunderbird.** It never sends iTIP itself; the
server always does. Tests showed that the organizer as `ATTENDEE` is not
needed. What matters is who sends.

## 2. How server-side scheduling works

### Send guard

`CalDavSynchronizer/Implementation/Events/ServerSchedulingSendGuard.cs`,
created in `ComponentContainer`.

It handles `Application.ItemSend` for meeting requests and cancellations that
the user sends as organizer. It acts only if the meeting's folder belongs to a
profile that

- is active,
- uploads Outlook changes (two-way, or Outlook to server),
- maps attendees,
- sets neither `SCHEDULE-AGENT=CLIENT` nor `X-SOGO-SEND-APPOINTMENT-NOTIFICATIONS`.

It takes over only **new** meetings (`PidLidFInvited` not set) and meetings it
already handed to the server, which carry the named property
`CalDavSynchronizerServerScheduled`. Meetings that Outlook sent itself earlier
stay with Outlook, and so do responses to received invitations. Otherwise
attendees would get updates from both Outlook and the server.

**Invitation or update:**

1. Still inside `ItemSend`, the guard sets the marker and saves the meeting.
   If saving fails, Outlook sends as usual.
2. It cancels the send.
3. Afterwards (posted to the synchronization context, because closing an
   inspector inside `ItemSend` is not allowed) it closes the inspector with
   `olDiscard`, then sets `PidLidFInvited` and saves again.

Setting `PidLidFInvited` **before** closing makes Outlook think the meeting
was already sent, so it sees the inspector's content as a change and asks
*"save changes and send update"*. The synchronization (change-triggered,
about 10 s later) uploads the meeting and the server sends the invitation.

**Cancellation of the whole meeting:** the guard cancels the send. Outlook
deletes the appointment itself when you delete it from the calendar view, but
keeps it when you cancel from the open meeting. The guard therefore deletes it
too and ignores "item has been moved or deleted". The synchronization deletes
the event on the server, and the server sends the `CANCEL`.

**Removing a single attendee:** Outlook's cancellation to that attendee is
suppressed; the server sends it after the upload.

Log entries start with `ItemSend:` in `%LOCALAPPDATA%\CalDavSynchronizer\log.txt`.

### Mapper

`EventEntityMapper.Map1To2` computes per meeting:

    scheduleAgentClient = profile.ScheduleAgentClient || !IsServerScheduled(appointment)

`SCHEDULE-AGENT=CLIENT` is therefore written for every meeting the guard did
not take over: meetings Outlook sent itself, and received invitations. This
holds even when the profile setting is off, so the server never sends a second
copy.

### Required profile settings

The automatic setup applies these:

- *Set SCHEDULE-AGENT=CLIENT*: off
- *Use GlobalAppointmentID for UID attribute*: on. The invitation now carries
  the server's UID. Outlook matches replies and counter proposals through the
  Global Object ID, so that UID has to be Outlook's GlobalAppointmentID.
- *Synchronize items immediately after change*: on

New events then get resource names like `040000008200E000...ics`
(upper-case hex). A pattern like `[0-9a-f-]+\.ics` misses them.

### Approaches tried and rejected

- **Turning off `SCHEDULE-AGENT=CLIENT` alone.** The server sends, and Outlook
  keeps sending too. Invitations arrive twice.
- **The SOGo profile.** It also sets `X-SOGO-SEND-APPOINTMENT-NOTIFICATIONS: NO`,
  which tells the server to stay silent because Outlook sent. sabre/dav
  ignores it, so invitations arrive twice.
- **Adding the organizer as attendee on the server.** No effect while Outlook
  sends: the invitation is already out before the server sees the event.
- ***Use GlobalAppointmentID for UID attribute* alone.** *Accept Proposal*
  stays greyed out.
- **Disabling synchronization.** *Accept Proposal* stays greyed out. This
  proved that the cause is what Outlook sends, not what the add-in changes
  afterwards.

## 3. Known limitations

- **Invitations work reliably only from the account's default calendar.**
  Outlook looks for meeting responses only there.
- **Changing a meeting and pressing *Send Update*** still shows Outlook's
  *"save changes and send update"* prompt. It is harmless: there is no second
  send. Removing it would require repurposing the inspector's Send button
  (ribbon XML) instead of cancelling `ItemSend`.
- **Hidden general options keep their values.** Switches removed from the UI
  (see section 5) are still in the data. A user who enabled one in the
  original keeps it, invisibly.
- **Outlook may disable the add-in for slow startup.** This is VSTO and .NET
  load time, not the add-in's own code. The installer adds the add-in to
  `DoNotDisableAddinList` (section 7).

## 4. Server requirements (sabre/dav)

- CalDAV scheduling (RFC 6638) with iMIP delivery. Deleting an organizer's
  event must send a `CANCEL`.
- **Recipients on the same server have to be invited by iMIP mail as well.**
  sabre/dav delivers locally into the recipient's calendar by default. That
  gives Outlook users no invitation mail and no Accept/Decline. Delivering both
  ways is not recommended: if the mail arrives before the synchronization,
  Outlook may end up with a duplicate.
- Discovery through `/.well-known/caldav` and `/.well-known/carddav` at
  `https://dav.<mail domain>/`, login with mail address and mail password.
- Sharing (`calendarserver-sharing`). sabre/dav reports the **sharee** as
  `DAV:owner` of a shared calendar instance. The owner is found in `DAV:invite`
  (the sharee with `shared-owner` access) or in `CS:invite/CS:organizer`.

## 5. Other changes

### Automatic setup

`CalDavSynchronizer/AutoSetup/AccountAutoSetup.cs`, run from the ribbon button
**Set up calendars** (`ComponentContainer.SetUpAccountsAsync`). A progress
window is shown while it runs, then a summary of what it found, added,
reactivated and deactivated. There is no synchronization afterwards (it would
slow the setup down). The first version ran at every startup; it was changed
to a button to avoid network access at startup.

**Accounts and login.** Each IMAP account in Outlook is set up against
`https://dav.<domain>/`. The login is the mail address with the IMAP password
stored in Outlook.

**Which folder a collection gets:**

- The own calendar `default` maps to the account's default calendar folder.
- The own address book `default` maps to the default contacts folder.
- All other collections become subfolders of these.
  - Shared ones are named after the owner's local part, for example
    `accounting`, with `owner - name` when the collection has a specific name.
  - Dots become spaces: Outlook refused *Cannot create the folder* for
    `accounting@bayarea-cc.com` in the IMAP store, probably because of the IMAP
    hierarchy delimiter.

**Sync mode.** Writable collections sync both ways with change-triggered sync.
Read-only ones sync server to Outlook only.

**Matching and updates.** Profiles are matched by URL, and profiles created
by hand are never changed.

- A collection that disappeared from the server gets its profile deactivated.
  The folder is kept.
- If nothing is found (server unreachable, login rejected), nothing changes.
- If one folder can't be created, only that collection is skipped and
  reported.

**Not covered.** Tasks (VTODO) are not set up. Profiles created here are
marked `IsAutoConfigured`.

### Stripped down to generic CalDAV/CardDAV

- **Removed projects and code:**
  - the projects `CalDavSynchronizer.OAuth.Google` and `CalDavSynchronizer.OAuth.Swisscom`
  - Google contact and task synchronization (Google APIs)
  - all profile types except `GenericProfile`
  - bulk profile setup (*Add multiple profiles*) and the Open-Xchange dialog
  - the provider logos and their tests
- **NuGet packages.** The Google and Zlib packages are removed.
  `Newtonsoft.Json` is now referenced directly; it used to come in only
  through the Google packages.
- **Kept for compatibility:**
  - The Google values of `ServerAdapterType` stay in the enum, so options files
    that contain them still deserialize. Otherwise the add-in would fail at
    startup. Such profiles fail when they synchronize instead.
  - Profiles of removed provider types fall back to *Generic*.
- **UI:**
  - With one profile type left, *Add new profile* no longer asks for the type.
  - General options no longer show the startup update check, *Accept invalid
    chars*, *useUnsafeHeaderParsing*, *Disable Certificate Validation*,
    *Client Certificates* and *SSL3*.
  - *Synchronize now* always shows the progress window. Automatic runs keep
    the threshold from the general options.

### About and update check

About shows:

- the logo (`Resources/BayAreaLogo.png`, embedded)
- the version
- the source code link required by the AGPL
- the original authors and project
- the support address

The license text starts with the modification notice. The application icon
was replaced as well, including the copies embedded in the WinForms dialogs.

The startup update check is disabled; it queried the upstream project and
would have offered its release. *Check for Updates* in About reads
`https://download.bayarea-cc.com/caldavsynchronizer/latest.json`:

```json
{ "version": "5.0.1", "url": "https://download.bayarea-cc.com/caldavsynchronizer/setup.exe", "notes": "optional" }
```

If that version is newer than the installed one, it offers to open the
download URL. It doesn't install anything. A template is in
`release/latest.json`.

## 6. Development environment

**Machine.** A Windows VM (for example 4 vCPUs, 8–16 GB RAM, 40 GB free) with
Outlook installed.

**Visual Studio.** Visual Studio 2022 with the Office workload, from an
administrator PowerShell:

```powershell
winget install --id Microsoft.VisualStudio.2022.Community -e --override "--add Microsoft.VisualStudio.Workload.Office --includeRecommended --passive --wait"
```

Then install the extension *Microsoft Visual Studio Installer Projects 2022*
(*Extensions → Manage Extensions*; it installs when Visual Studio is closed,
confirm with *Modify*). Without it, `CalDavSynchronizer.Setup` shows as
incompatible.

**Signing certificate.** Manifests are signed with a self-signed certificate,
`CN=Bay Area Affiliates`, valid 5 years. The thumbprint is set in
`CalDavSynchronizer.csproj`. The private key is only in the backup `.pfx`,
never in the repository. To build on another machine, import the `.pfx` into
*Personal* certificates.

**Building.**

- Work from a local clone (for example `C:\src\...`). VSTO trust doesn't work
  from network shares.
- Develop and debug in *Debug* (F5 registers the add-in as `CalDavSynchronizer`).
  Before installing a setup on the same machine, run *Build → Clean Solution*;
  otherwise Outlook loads both.
- All projects treat warnings as errors.

## 7. Installer

`CalDavSynchronizer.Setup/CalDavSynchronizer.Setup.vdproj`.

- **Product and version.** Product *CalDavSynchronizer (Bay Area Edition)*,
  manufacturer *Bay Area Affiliates*, version 5.0.0 (assembly 5.0.0.0).
- **Replacing the original.** The upgrade code is the original's, so an
  installed original (up to 4.7.1) is replaced. For that the version must be
  higher, because a Windows Installer only replaces older versions. The
  original's signing step was removed.
- **Build in *Release*.** `CalDavSynchronizer.vsto` and `.dll.manifest` are
  referenced from `bin\Release`, while the DLL comes from the selected
  configuration. A Debug setup would mix Release manifests with a Debug DLL,
  and VSTO refuses to load that because of the hashes. The result is in
  `CalDavSynchronizer.Setup\Release\`.
- **Registry entries.** The installed add-in is registered as
  `CalDavSynchronizer.1`. The installer writes, for the installing user:
  - `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList`,
    value `CalDavSynchronizer.1` = 1, so Outlook doesn't disable the add-in for
    slow startup.
  - the VSTO inclusion list
    (`HKCU\Software\Microsoft\VSTO\Security\Inclusion\{GUID}`: manifest URL and
    public key), so Outlook doesn't ask whether to trust the publisher on first
    start.
- **When a purchased certificate would help.** A purchased code-signing
  certificate would remove Windows' "unknown publisher" warning on
  `setup.exe` and the trust prompt for other Windows users on the same PC.
  Let's Encrypt doesn't issue code-signing certificates; Certum offers one for
  open-source developers.

## 8. Release process

1. Raise the version:
   - the setup's `ProductVersion` (answer *Yes* when Visual Studio asks to
     change the ProductCode)
   - `CalDavSynchronizer/Properties/AssemblyInfoVersion.cs`
2. Build in *Release* and build the setup.
3. Uninstalling first is only needed for a setup with the same version.
4. Upload `setup.exe` (and `.msi`) to
   `https://download.bayarea-cc.com/caldavsynchronizer/`.
5. Update `latest.json` there.
6. Commit, tag (`v5.0.1`), push. The AGPL requires the source of every
   released version to be available.

## 9. Test log

| Date | Test | Result |
| --- | --- | --- |
| 2026-10-04 | `SCHEDULE-AGENT=CLIENT` off | server and Outlook both send, invitation twice |
| 2026-10-05 | *Use GlobalAppointmentID* | *Accept Proposal* greyed out |
| 2026-10-05 | profile disabled | *Accept Proposal* greyed out: the cause is what Outlook sends |
| 2026-10-05 | guard, new meeting | server sends, counter proposal arrives, ***Accept Proposal* works**; neither `SCHEDULE-AGENT=CLIENT` nor organizer as attendee on the server |
| 2026-10-05 | cancel from open meeting | Outlook keeps the appointment, so the guard deletes it |
| 2026-10-05 | cancel, server side | `DELETE` ok; missing `CANCEL` was a server bug, fixed there |
| 2026-10-06 | *Accept Proposal* + *Send Update* | one update at the attendee |
| 2026-10-06 | cancel from calendar view | one cancellation |
| 2026-10-06 | remove one of two attendees | removed attendee gets the cancellation from the server |
| 2026-10-06 | "send update" prompt on first send | caused by setting `PidLidFInvited` before closing; fixed |
| 2026-10-06 | recipient on the same server | no mail (local delivery), server change needed (section 4) |
| 2026-10-06 | automatic setup | own calendar and address book on default folders, shared calendar as read-only subfolder named after its owner |
| 2026-10-06 | installer | no trust prompt, add-in not disabled after several restarts |
