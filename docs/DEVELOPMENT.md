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
stay with Outlook. Otherwise attendees would get updates from both Outlook
and the server.

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

**Responses (accept, tentative, decline)** to received meetings, with the same
profile conditions:

1. Still inside `ItemSend`, the guard sets the marker and `PidLidResponseStatus`
   on the appointment (in case Outlook sets it only after sending) and saves.
2. It cancels the send. A response inspector (*Edit the response before
   sending*) is closed afterwards.
3. The guard starts a synchronization right away: the change trigger
   (`FolderChangeWatcher`) deliberately ignores received meetings, so
   otherwise the response would wait for the next interval.
4. The synchronization uploads the attendee's `PARTSTAT` without
   `SCHEDULE-AGENT=CLIENT`, and sabre/dav's iTIP broker sends the `REPLY`, to
   an internal organizer directly into the organizer's event (also in a shared
   calendar), to an external one by iMIP.

*Don't send a response* triggers no `ItemSend`, so the appointment is not
marked and the server stays silent too.

**Responses in the organizer's Outlook:** Outlook applies a response mail only
to a meeting in the default calendar; for one in a shared calendar it shows
*"We couldn't find this meeting in the calendar"*. The server enters the
`REPLY` into the organizer's event, but Outlook's *Tracking* page doesn't
show it: `PR_RECIPIENT_TRACKSTATUS` can't be set through the object model
(`0x80040102`, tried 2026-10-19); only Extended MAPI (`IMessage::ModifyRecipients`)
could. The organizer sees the response in the server's mail. So that the next
upload doesn't reset it, an attendee Outlook shows without response keeps the
`PARTSTAT` of the event on the server, as long as start, end and recurrence
are unchanged. When the time changes, the responses are reset (RFC 5546):
otherwise the update tells the attendee he already accepted, his new
acceptance changes nothing, and the server sends no `REPLY` (seen
2026-10-19).

**Response before the first synchronization:** an invitation from a user on
the same server exists twice before the add-in knows either: as the copy the
server delivered (`sabredav-<uuid>.ics`) and as the appointment Outlook made
from the mail. The synchronization pairs them by subject and time (initial
entity matching), which normally counts both as unchanged, so the response
was never uploaded (seen 2026-10-19). For an appointment with the marker,
`InitialEventEntityMatcher` hands back a stale Outlook version, and the
synchronizer then treats the Outlook side as changed and uploads it in the
same run. The pairing itself stays: it prevents a second copy in Outlook and
on the server.

**Counter proposals stay with Outlook.** *Propose New Time* is sent as a
tentative or negative response with `PidLidAppointmentCounterProposal` set.
The uploaded appointment keeps the original time, so the proposal would be
lost, and sabre/dav's broker doesn't process `COUNTER` (`SCHEDULE-STATUS:
5.0`). The organizer's *Accept Proposal* works as before for meetings in the
default calendar.

**Responding in a calendar shared by someone else:** there the attendee sees
the organizer's event itself. Declining (or deleting) it, Outlook removes it
before `ItemSend`, so the guard lets Outlook send the response. Without
protection the synchronization would then delete the organizer's event on the
server, for everybody. Profiles of such calendars therefore carry
`SharedCollectionOwner` (set by *Set up calendars*), and `CalDavRepository`
doesn't delete a meeting organized by someone else from them. It only drops
the relation, so the next synchronization brings the event back to Outlook
(report warning *"Not deleted on the server ..."*). Events without organizer
and the user's own meetings are deleted as usual.

**Deleting the own copy of someone else's meeting:** declining from the
invitation mail, Outlook removes the appointment *before* `ItemSend`; the
guard can't even read it (`MAPI_E_OBJECT_DELETED`), so Outlook sends the
decline. If that copy was answered through the server before, it has no
`SCHEDULE-AGENT=CLIENT`, and sabre/dav would send a second `REPLY DECLINED`
when the synchronization deletes it (seen 2026-10-19). `CalDavRepository`
therefore sets `SCHEDULE-AGENT=CLIENT` on such a copy before deleting it, so
deleting never makes the server reply. Outlook's choice applies: its decline,
or nothing with *Don't send a response*.

**Decline:** if Outlook removes the appointment despite the cancelled send,
the synchronization would only delete it on the server, where it still has
`SCHEDULE-AGENT=CLIENT`, so no `REPLY`. The guard checks afterwards and in that
case lets Outlook send the response after all (log: *letting Outlook send
it*).

Log entries start with `ItemSend:` in `%LOCALAPPDATA%\CalDavSynchronizer\log.txt`.

### Mapper

`EventEntityMapper.Map1To2` computes per meeting:

    scheduleAgentClient = profile.ScheduleAgentClient || !IsServerScheduled(appointment)

`SCHEDULE-AGENT=CLIENT` is therefore written for every meeting the guard did
not take over: meetings Outlook sent itself, and received invitations not yet
answered through the guard. This holds even when the profile setting is off,
so the server never sends a second copy.

**Reminders** (`ReminderMapping.JustUpcoming`, the default): in a shared
calendar there is one event with one `VALARM` for everybody; sabre/dav keeps
no per-user alarms. The rule is about the **event**, not the reminder time:

- A reminder is mapped in both directions while the event hasn't ended, even
  if it is due already (set deliberately late, it pops up for everybody at
  once). Reminders of past events are left out, otherwise the first
  synchronization of a calendar with history would show hundreds of overdue
  reminders.
- Dismissing a reminder clears `ReminderSet` in Outlook. On upload, a due
  `VALARM` is kept on the server until the event has ended, so one member's
  dismiss doesn't remove it for the others. Removing a reminder that is not
  due yet removes it for everybody.
- From the server, a due reminder is not set again on an Outlook item where it
  was dismissed, so it doesn't pop up again on every change of the event.

### Required profile settings

The automatic setup applies these:

- *Set SCHEDULE-AGENT=CLIENT*: off
- *Use GlobalAppointmentID for UID attribute*: on. The invitation now carries
  the server's UID. Outlook matches replies and counter proposals through the
  Global Object ID, so that UID has to be Outlook's GlobalAppointmentID.
- *Synchronize items immediately after change*: on

New events then get resource names like `040000008200E000...ics`
(upper-case hex). A pattern like `[0-9a-f-]+\.ics` misses them.

sabre/dav's PostgreSQL schema stores the UID and the resource name in
`varchar(200)` columns; a longer one fails with `500`, *"value too long for
type character varying(200)"*. Outlook's own GlobalAppointmentIDs are 112
characters, but for an invitation from another calendar system (Google,
Teams, booking tools) Outlook embeds the organizer's UID, so the
GlobalAppointmentID is 106 + 2 × its length. Above 196 characters the add-in
uses that embedded UID instead (it is the organizer's UID anyway), and if
that is too long as well, a new GUID.

The UID is also the resource name, but only if it is a plain path segment
(letters, digits, `-_.@`, at most 196 characters); otherwise the name is a
GUID. A UID can be anything: Adobe Connect uses a URL, and `new Uri(server,
"https://…")` drops the server, so 5.2.1 sent the PUT, with the Basic
credentials of the profile, to adobeconnect.com (answer `401` without text).
`WebDavDataAccess.GetUrlInCollection` now refuses every new resource URL
outside the profile's collection.

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
- **Invitations between two accounts of the same Outlook give two copies.**
  Outlook puts meeting requests for any IMAP account into the calendar of the
  default data file, next to the organizer's copy and with the same UID
  (GlobalAppointmentID). The upload of the second copy fails with `412
  Precondition Failed` (`If-None-Match: *`, the resource `<UID>.ics` exists).
  The add-in then fetches the existing resource; with the same UID it doesn't
  overwrite it and reports a warning instead of an error (no popup with the
  default report settings). The warning repeats on every run until one copy is
  deleted in Outlook. With a different UID it stays an error.

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
stored in Outlook. The host `dav.<domain>` is a CNAME on the server; DNS SRV
records (RFC 6764) are not used, since they would be one more entry per domain
to forget on a move. If the host doesn't resolve, the summary says so instead
of "nothing found". Accounts that are not IMAP are listed as skipped.

**Which folder a collection gets:**

- The own calendar `default` maps to the account's default calendar folder.
- The own address book `default` maps to the default contacts folder.
- All other collections become subfolders of these.
  - Shared ones are named after the owner's local part, for example
    `accounting`, with `owner - name` when the collection has a specific name.
  - Global contact lists of the company (`gal-<name>` on the server) have no
    owner and are named after their display name, for example `customers`.
  - Dots become spaces: Outlook refused *Cannot create the folder* for
    `accounting@bayarea-cc.com` in the IMAP store, probably because of the IMAP
    hierarchy delimiter.

**Sync mode.** Taken from `{DAV:}current-user-privilege-set`, never guessed
from the name. Writable collections sync both ways with change-triggered sync.
Read-only ones (no write privilege) sync server to Outlook only.

**No separate .pst.** Subfolders are created through the Outlook object model
with a fixed folder type, which gives a real calendar or contacts folder even in
an IMAP store. Only Outlook's *New Folder* dialog turns them into mail folders.

**Matching and updates.** Profiles are matched by URL, and profiles created
by hand are never changed.

- A collection that disappeared from the server gets its profile deactivated
  and ` (removed on server)` appended to its name, so nobody re-enables it
  without knowing why. The folder is kept: it may be the last copy of the
  data. If the collection comes back, the profile is reactivated and renamed.
- A profile whose Outlook account is gone (matched by the account's SMTP
  address) is deactivated and gets ` (account removed)` appended; its folder
  was in the account's data file. If the account is added again, the profile
  is replaced by a new one pointing to the new data file.
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
{ "version": "5.0.1", "url": "https://download.bayarea-cc.com/caldavsynchronizer/CalDavSynchronizer-BayArea-5.2.2.zip", "notes": "optional" }
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

Short version with the update check: [UPDATES.md](UPDATES.md).

1. Raise the version:
   - the setup's `ProductVersion` (answer *Yes* when Visual Studio asks to
     change the ProductCode)
   - `CalDavSynchronizer/Properties/AssemblyInfoVersion.cs`
2. Build in *Release* and build the setup.
3. Uninstalling first is only needed for a setup with the same version.
4. Zip both files of `CalDavSynchronizer.Setup\Release\`. `setup.exe` is
   only the bootstrapper: it checks the prerequisites and runs the `.msi`
   next to it.
   ```powershell
   Compress-Archive -Path CalDavSynchronizer.Setup\Release\setup.exe, CalDavSynchronizer.Setup\Release\CalDavSynchronizer.Setup.msi -DestinationPath CalDavSynchronizer-BayArea-5.2.2.zip
   ```
   Upload the zip to `https://download.bayarea-cc.com/caldavsynchronizer/`,
   for example with the OpenSSH client built into Windows:
   ```powershell
   scp CalDavSynchronizer-BayArea-5.2.2.zip release\latest.json <user>@<download server>:<download directory>/caldavsynchronizer/
   ```
   The version in the name keeps browsers and proxies from serving an old
   cached file.
5. Update `latest.json` there (`url` pointing to the new zip). Upload the zip
   first, so `latest.json` never points to a missing file.
6. Commit, tag (`v5.2.2`), push. The AGPL requires the source of every
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
| 2026-10-08 | recipient on the same server, after the server change | exactly one invitation, exactly one reply back |
| 2026-10-06 | automatic setup | own calendar and address book on default folders, shared calendar as read-only subfolder named after its owner |
| 2026-10-06 | installer | no trust prompt, add-in not disabled after several restarts |
| 2026-10-19 | reminder due already, event tomorrow | pops up at once for the creator and, after synchronizing, for the other member of the shared calendar; `VALARM` stays on the server |
| 2026-10-19 | accept, invitation from the own calendar | Outlook's response suppressed, server sends the `REPLY` (sender `caldav@…`), organizer's Outlook applies it |
| 2026-10-19 | accept, invitation from a shared calendar | `REPLY` entered into the organizer's event (`PARTSTAT=ACCEPTED`); organizer gets the mail, Outlook can't apply it (*couldn't find this meeting*), *Tracking* not updated |
| 2026-10-19 | accept before the first synchronization | response was lost (initial matching); fixed by uploading marked meetings after matching |
| 2026-10-19 | accept of an update | response waited for the scheduled synchronization; fixed by synchronizing right after the response |
| 2026-10-19 | decline from the invitation mail after accepting through the server | was delivered twice (Outlook and the server on delete); now exactly one, from Outlook |
| 2026-10-20 | first sync of a customer calendar with existing invitations from other systems | `500`, UID longer than 200 characters; fixed in 5.2.1 |
| 2026-10-20 | Adobe Connect invitations (UID is a URL) | 5.2.1 sent the PUT to adobeconnect.com, `401`; fixed in 5.2.2 |
| 2026-10-15 | 412 on *Synchronize now* | attendee copy of an invitation from `andreas.knuth@` to `info@haiky.app` (same Outlook) in the organizer's calendar, same UID as the organizer copy; now a warning instead of an error |
| 2026-10-13 | automatic setup, global contact list | `info@haiky.app`: calendar, own address book and the company list `customers` (display name, not `gal-customers`) as a real contacts subfolder; contacts synced |
