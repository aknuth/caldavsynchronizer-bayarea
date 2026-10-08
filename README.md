# CalDavSynchronizer (Bay Area Edition)

An Outlook add-in that synchronizes Outlook calendars and contacts with a
CalDAV/CardDAV server. This is a modified version of
[CalDavSynchronizer](https://github.com/aluxnimm/outlookcaldavsynchronizer)
by Gerhard Zehetbauer and Alexander Nimmervoll, adapted to a single kind of
server: [sabre/dav](https://sabre.io/dav/) with server-side scheduling
(RFC 6638), where users log in with their mail address and mail password.

## What is different from the original

- **The server sends meeting invitations, not Outlook.** When you send a
  meeting from a synchronized calendar, the add-in cancels Outlook's own iTIP
  mail, saves the meeting and lets the synchronization upload it. The CalDAV
  server then sends the invitation, update or cancellation exactly once. This
  makes Outlook behave like Thunderbird: replies and counter proposals from
  other calendars (Google, ...) are matched to the meeting, and *Accept
  Proposal* works. Meetings that Outlook sent itself before keep being handled
  by Outlook (`SCHEDULE-AGENT=CLIENT`), so nobody gets anything twice.
- **Automatic setup.** *Set up calendars* in the ribbon finds every calendar
  and address book of each IMAP account in Outlook on
  `https://dav.<mail domain>/` (login: mail address and the IMAP password
  stored in Outlook) and creates a synchronization profile for each:
  - the own default calendar and address book map to the account's default
    Outlook folders (Outlook looks for meeting responses only there),
  - other own and shared collections become subfolders, shared ones named
    after their owner,
  - read-only collections are only synchronized from the server to Outlook.
- **Reduced to generic CalDAV/CardDAV.** Google, Swisscom and the other
  provider profiles, OAuth, and the bulk profile setup are removed.
- The update check in *About* reads `latest.json` from our release server.

How it works, the reasons behind it, known limitations and the server
requirements are described in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Requirements

- Windows with Outlook 2016 or later (Microsoft 365), .NET Framework 4.8 and
  the Visual Studio Tools for Office runtime (the installer adds them if
  missing).
- An IMAP account in Outlook whose mail domain has a sabre/dav server at
  `https://dav.<mail domain>/` with CalDAV scheduling enabled.

## Usage

1. Unzip the release and run `setup.exe`.
2. In Outlook, open the *CalDav Synchronizer* ribbon and click
   **Set up calendars**.
3. Click **Synchronize now** (otherwise the profiles synchronize at their
   interval).

Invitations work reliably only from the default calendar of the account.

A step-by-step checklist for installing at a customer is in
[docs/CUSTOMER-SETUP.md](docs/CUSTOMER-SETUP.md).

## Building

- Visual Studio 2022 with the *Office/SharePoint development* workload and the
  extension *Microsoft Visual Studio Installer Projects 2022*.
- The add-in manifests are signed with the certificate whose thumbprint is set
  in `CalDavSynchronizer/CalDavSynchronizer.csproj`; it has to be in the
  personal certificate store of the user who builds.
- Develop and debug (F5) in the *Debug* configuration.
- Build the installer in *Release*: build `CalDavSynchronizer.Setup`, the
  result is in `CalDavSynchronizer.Setup/Release/`.

## Releasing

1. Raise the version in the setup project (`ProductVersion`, answer *Yes* when
   Visual Studio asks to change the ProductCode) and in
   `CalDavSynchronizer/Properties/AssemblyInfoVersion.cs`.
2. Build the installer in *Release*.
3. Zip `setup.exe` and `CalDavSynchronizer.Setup.msi` (both are needed:
   `setup.exe` checks the prerequisites and runs the `.msi` next to it) as
   `CalDavSynchronizer-BayArea-<version>.zip`, upload it to the release server
   and update `latest.json` there (template in `release/latest.json`).
4. Tag the commit and push it, so the source of every released version is
   available.

## License

GNU Affero General Public License v3.0, see [LICENSE.txt](LICENSE.txt).

Copyright (c) 2015-2019 Gerhard Zehetbauer, Alexander Nimmervoll
(original CalDavSynchronizer), Copyright (c) 2026 Bay Area Affiliates
(modifications).

Support: support@bayarea-cc.com
