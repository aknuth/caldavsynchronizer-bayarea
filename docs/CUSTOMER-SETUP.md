# Customer setup checklist

Setting up CalDavSynchronizer (Bay Area Edition) on a customer's PC.

## Before you start

- [ ] Windows 10 or 11.
- [ ] **Classic Outlook** 2016 or later, or Microsoft 365 for desktop. The *new
      Outlook for Windows* does not support add-ins of this kind. If it is
      active, switch back to classic Outlook (toggle at the top right).
- [ ] You are logged in to Windows **as the customer's user**, not as an
      administrator account. The installer writes two settings for the user
      who installs it (see step 2).
- [ ] The mail address and mail password of the customer.

## 1. Mail account in Outlook

- [ ] Add the mailbox as an **IMAP** account. Check *Remember password*:
      the add-in logs in with the IMAP password stored in Outlook.
- [ ] *File → Account Settings → Data Files*: the data file of this account
      is the **default** (*Set as Default*). Outlook looks for meeting
      responses only in the default calendar. With another default, everything
      works except replies to invitations.
- [ ] Close Outlook.

## 2. Install

- [ ] Download `CalDavSynchronizer-BayArea-<version>.zip` from
      `https://download.bayarea-cc.com/caldavsynchronizer/` and unzip it.
- [ ] Run `setup.exe` (not only the `.msi`). It installs .NET Framework 4.8 and
      the Office runtime first, if missing.
- [ ] Choose **Just me**.
- [ ] An installed original CalDavSynchronizer is replaced automatically. Its
      profiles are kept.

The installer adds two entries for the installing Windows user:

- Outlook's list of add-ins that must never be disabled for slow startup.
- The VSTO trust list, so Outlook doesn't ask whether to trust the publisher.

## 3. Set up

- [ ] Start Outlook. There is no security prompt, and the ribbon has a tab
      **CalDav Synchronizer**.
- [ ] Click **Set up calendars**. The summary lists, per account, the
      calendars and address books found and the profiles added:
  - the own calendar on the account's *Calendar* folder,
  - the own address book on *Contacts*,
  - shared calendars as subfolders named after their owner, read only if
    shared read only.
- [ ] Click **Synchronize now** and wait for the progress window to finish.
- [ ] Shared calendars appear under *My Calendars*. Tick them to show them.

## 4. Test

- [ ] Create a meeting in the **Calendar** folder with an external attendee
      (for example a Gmail address) and send it. The window closes without a
      question.
- [ ] The attendee receives **exactly one** invitation, sent by the server.
- [ ] The attendee accepts. The response arrives in Outlook and the attendee's
      status is updated.
- [ ] Optional: the attendee proposes a new time. In Outlook, **Accept
      Proposal** is enabled.
- [ ] Delete the test meeting and send the cancellation. The attendee receives
      exactly one cancellation.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| Ribbon tab missing, *"add-in disabled"* bar | *File → Options → Add-ins → Manage COM Add-ins*: enable *CalDavSynchronizer*. Or in *Slow and Disabled Add-ins* choose *Always enable this add-in*. |
| *Set up calendars*: "no IMAP password stored" | Open the account settings, enter the password and check *Remember password*. |
| *Set up calendars* finds nothing | Check that `https://dav.<mail domain>/` is reachable and the mail password is correct. |
| Responses to invitations don't update the meeting | The account's data file is not the default (step 1). |
| Invitation arrives twice | The meeting was sent from a folder whose profile still has *Set SCHEDULE-AGENT=CLIENT* on. Run *Set up calendars* or switch it off in the profile's *Event Mapping Configuration*. |
| Anything else | *CalDav Synchronizer → Reports* for the last synchronization runs. The log is `%LOCALAPPDATA%\CalDavSynchronizer\log.txt` (lines starting with `ItemSend:` are from sending meetings). |

Support: support@bayarea-cc.com
