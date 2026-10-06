// This file is Part of CalDavSynchronizer (http://outlookcaldavsynchronizer.sourceforge.net/)
// Copyright (c) 2015 Gerhard Zehetbauer
// Copyright (c) 2015 Alexander Nimmervoll
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as
// published by the Free Software Foundation, either version 3 of the
// License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using CalDavSynchronizer.Contracts;
using CalDavSynchronizer.DataAccess;
using CalDavSynchronizer.Implementation;
using CalDavSynchronizer.Implementation.ComWrappers;
using CalDavSynchronizer.ProfileTypes.ConcreteTypes;
using CalDavSynchronizer.Scheduling;
using CalDavSynchronizer.Ui.ConnectionTests;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace CalDavSynchronizer.AutoSetup
{
    /// <summary>
    /// Configures a synchronization profile for every calendar and address book found on the DAV server
    /// of each IMAP account in Outlook, without any user input.
    /// The server is expected at https://dav.{domain of the account's email address}/, the credentials are
    /// the email address and the IMAP password stored in Outlook.
    /// The own default collections map to the account's default calendar and contacts folders (Outlook looks
    /// for meeting responses only in the default calendar), all others to subfolders of them.
    /// Profiles found again are kept, profiles whose collection is gone are deactivated, folders are never deleted.
    /// Profiles configured by hand (not marked IsAutoConfigured) are left untouched.
    /// </summary>
    public class AccountAutoSetup
    {
        private static readonly ILog s_logger = LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType);

        private const string DefaultCollectionName = "default";

        private readonly NameSpace _session;
        private readonly IOutlookAccountPasswordProvider _passwordProvider;
        private readonly GenericProfile _profileType = new GenericProfile();

        public AccountAutoSetup(NameSpace session, IOutlookAccountPasswordProvider passwordProvider)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _passwordProvider = passwordProvider ?? throw new ArgumentNullException(nameof(passwordProvider));
        }

        /// <param name="report">Receives one line per account and per change, to be shown to the user.</param>
        /// <returns>The new profile list, or null if nothing changed.</returns>
        public async Task<Options[]> GetUpdatedOptionsOrNull(Options[] currentOptions, GeneralOptions generalOptions, List<string> report)
        {
            var options = currentOptions.ToList();
            var changed = false;

            foreach (var account in _session.Accounts.ToSafeEnumerable<Account>())
            {
                if (account.AccountType != OlAccountType.olImap)
                    continue;

                var accountName = account.DisplayName;
                try
                {
                    changed |= await SetUpAccount(account, options, generalOptions, report);
                }
                catch (Exception x)
                {
                    s_logger.Error("AutoSetup: error setting up account, profiles of this account stay as they are.", x);
                    report.Add($"{accountName}: {x.Message}");
                }
            }

            return changed ? options.ToArray() : null;
        }

        private async Task<bool> SetUpAccount(Account account, List<Options> options, GeneralOptions generalOptions, List<string> report)
        {
            var accountName = account.DisplayName;
            var emailAddress = account.SmtpAddress;
            var atIndex = emailAddress?.IndexOf('@') ?? -1;
            if (atIndex < 0)
            {
                s_logger.Warn($"AutoSetup: account '{accountName}' has no email address, skipped.");
                report.Add($"{accountName}: no email address, skipped.");
                return false;
            }

            var password = _passwordProvider.GetPassword(accountName);
            if (password.Length == 0)
            {
                s_logger.Warn($"AutoSetup: no IMAP password stored for account '{accountName}', skipped.");
                report.Add($"{emailAddress}: no IMAP password stored in Outlook, skipped.");
                return false;
            }

            var serverUrl = new Uri($"https://dav.{emailAddress.Substring(atIndex + 1)}/");
            var webDavClient = SynchronizerFactory.CreateWebDavClient(
                emailAddress,
                password,
                serverUrl.ToString(),
                generalOptions.CalDavConnectTimeout,
                ServerAdapterType.WebDavHttpClientBased,
                false,
                true,
                true,
                new ProxyOptions {ProxyUseDefault = true},
                generalOptions.EnableClientCertificate,
                generalOptions.AcceptInvalidCharsInServerResponse);

            var calDavDataAccess = new CalDavDataAccess(serverUrl, webDavClient);
            var calendars = (await calDavDataAccess.GetUserResourcesIncludingCalendarProxies(true)).CalendarResources;
            var addressBooks = await new CardDavDataAccess(serverUrl, webDavClient, string.Empty, contentType => true).GetUserAddressBooksNoThrow(true);
            var summary = $"{emailAddress} ({serverUrl}): {calendars.Count} calendar(s), {addressBooks.Count} address book(s) found.";
            s_logger.Info($"AutoSetup: {summary}");
            report.Add(summary);

            var calendarCollections = new List<Collection>();
            foreach (var calendar in calendars)
            {
                var ownerEmail = await calDavDataAccess.GetSharingOwnerEmailOrNull(calendar.Uri)
                                 ?? (calendar.OwnerProperties != null && calendar.OwnerProperties.IsSharedCalendar ? calendar.OwnerProperties.CalendarOwnerEmail : null);
                if (string.Equals(ownerEmail, emailAddress, StringComparison.OrdinalIgnoreCase))
                    ownerEmail = null;
                calendarCollections.Add(new Collection(calendar.Uri, calendar.Name, ownerEmail, calendar.Privileges));
            }

            var changed = false;
            using (var store = GenericComObjectWrapper.Create(account.DeliveryStore))
            {
                // An account always has at least its default collections. Nothing found means the server was not
                // reachable or rejected the login, so the existing profiles must not be deactivated.
                if (calendars.Count > 0)
                {
                    using (var calendarFolder = GenericComObjectWrapper.Create((Folder) store.Inner.GetDefaultFolder(OlDefaultFolders.olFolderCalendar)))
                    {
                        changed |= Reconcile(calendarCollections, calendarFolder.Inner, OlDefaultFolders.olFolderCalendar, accountName, emailAddress, options, CreateEventMappingConfiguration, report);
                    }
                }

                if (addressBooks.Count > 0)
                {
                    using (var contactsFolder = GenericComObjectWrapper.Create((Folder) store.Inner.GetDefaultFolder(OlDefaultFolders.olFolderContacts)))
                    {
                        var collections = addressBooks.Select(a => new Collection(a.Uri, a.Name, null, a.Privileges));
                        changed |= Reconcile(collections, contactsFolder.Inner, OlDefaultFolders.olFolderContacts, accountName, emailAddress, options, _profileType.CreateContactMappingConfiguration, report);
                    }
                }
            }

            return changed;
        }

        private EventMappingConfiguration CreateEventMappingConfiguration()
        {
            var mapping = _profileType.CreateEventMappingConfiguration();
            // The server sends iTIP (see ServerSchedulingSendGuard). Its UID must be Outlook's GlobalAppointmentID,
            // otherwise Outlook can't match the responses to the meeting.
            mapping.ScheduleAgentClient = false;
            mapping.UseGlobalAppointmentID = true;
            return mapping;
        }

        private bool Reconcile(
            IEnumerable<Collection> collections,
            Folder defaultFolder,
            OlDefaultFolders folderType,
            string accountName,
            string emailAddress,
            List<Options> options,
            Func<MappingConfigurationBase> createMappingConfiguration,
            List<string> report)
        {
            var changed = false;
            var mappingType = createMappingConfiguration().GetType();
            var foundUrls = new HashSet<string>();

            foreach (var collection in collections)
            {
                foundUrls.Add(NormalizeUrl(collection.Uri.ToString()));
                var synchronizationMode = collection.IsWritable ? SynchronizationMode.MergeInBothDirections : SynchronizationMode.ReplicateServerIntoOutlook;

                var existing = options.FirstOrDefault(o => NormalizeUrl(o.CalenderUrl) == NormalizeUrl(collection.Uri.ToString()));
                if (existing != null)
                {
                    if (existing.IsAutoConfigured && (existing.Inactive || existing.SynchronizationMode != synchronizationMode))
                    {
                        s_logger.Info($"AutoSetup: updating profile '{existing.Name}'.");
                        report.Add($"  {(existing.Inactive ? "reactivated" : "updated")}: {existing.Name}");
                        existing.Inactive = false;
                        existing.SynchronizationMode = synchronizationMode;
                        existing.EnableChangeTriggeredSynchronization = collection.IsWritable;
                        changed = true;
                    }

                    continue;
                }

                var isDefault = collection.IsOwn && GetLastSegment(collection.Uri) == DefaultCollectionName;
                var displayName = GetFolderName(collection);

                string folderEntryId;
                string folderStoreId;
                if (isDefault)
                {
                    folderEntryId = defaultFolder.EntryID;
                    folderStoreId = defaultFolder.StoreID;
                }
                else
                {
                    using (var subFolder = GenericComObjectWrapper.Create(GetOrCreateSubFolder(defaultFolder, displayName, folderType)))
                    {
                        folderEntryId = subFolder.Inner.EntryID;
                        folderStoreId = subFolder.Inner.StoreID;
                    }
                }

                var newOptions = _profileType.CreateOptions();
                newOptions.IsAutoConfigured = true;
                newOptions.Name = isDefault ? $"{emailAddress} ({collection.Name})" : $"{emailAddress} - {displayName}";
                newOptions.EmailAddress = emailAddress;
                newOptions.UserName = emailAddress;
                newOptions.UseAccountPassword = true;
                newOptions.CalenderUrl = collection.Uri.ToString();
                newOptions.OutlookFolderEntryId = folderEntryId;
                newOptions.OutlookFolderStoreId = folderStoreId;
                newOptions.OutlookFolderAccountName = accountName;
                newOptions.SynchronizationMode = synchronizationMode;
                newOptions.EnableChangeTriggeredSynchronization = collection.IsWritable;
                newOptions.MappingConfiguration = createMappingConfiguration();

                s_logger.Info($"AutoSetup: adding profile '{newOptions.Name}' for '{newOptions.CalenderUrl}'.");
                report.Add($"  added: {newOptions.Name}{(collection.IsWritable ? "" : " (read only)")}");
                options.Add(newOptions);
                changed = true;
            }

            foreach (var gone in options.Where(o =>
                o.IsAutoConfigured
                && !o.Inactive
                && o.EmailAddress == emailAddress
                && o.MappingConfiguration?.GetType() == mappingType
                && !foundUrls.Contains(NormalizeUrl(o.CalenderUrl))))
            {
                // Keep the folder, the user may have stored something in it.
                s_logger.Info($"AutoSetup: '{gone.CalenderUrl}' no longer on the server, deactivating profile '{gone.Name}'.");
                gone.Inactive = true;
                report.Add($"  deactivated (no longer on the server): {gone.Name}");
                changed = true;
            }

            return changed;
        }

        private static Folder GetOrCreateSubFolder(Folder parent, string name, OlDefaultFolders folderType)
        {
            name = name.Replace('\\', '-').Replace('/', '-');
            using (var folders = GenericComObjectWrapper.Create(parent.Folders))
            {
                try
                {
                    return (Folder) folders.Inner[name];
                }
                catch (COMException)
                {
                    var folder = (Folder) folders.Inner.Add(name, folderType);
                    // Folders in IMAP stores get " (This computer only)" appended.
                    folder.Name = name;
                    return folder;
                }
            }
        }

        private static readonly HashSet<string> s_genericCollectionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Calendar", "Default Calendar", "Kalender", "Contacts", "Default Address Book", "Kontakte", DefaultCollectionName
        };

        /// <summary>
        /// Shared collections are named after their owner, since they are typically just called "Calendar",
        /// which would show up as a second "Calendar" in Outlook. A specific name is appended.
        /// </summary>
        private static string GetFolderName(Collection collection)
        {
            if (collection.OwnerEmailOrNull == null)
                return collection.Name;
            if (string.IsNullOrEmpty(collection.Name) || s_genericCollectionNames.Contains(collection.Name))
                return collection.OwnerEmailOrNull;
            return $"{collection.OwnerEmailOrNull} - {collection.Name}";
        }

        private static string GetLastSegment(Uri uri) => uri.Segments.Last().TrimEnd('/');

        private static string NormalizeUrl(string url) => url == null ? string.Empty : Uri.UnescapeDataString(url).TrimEnd('/').ToLowerInvariant();

        private class Collection
        {
            public Collection(Uri uri, string name, string ownerEmailOrNull, AccessPrivileges privileges)
            {
                Uri = uri;
                Name = name;
                OwnerEmailOrNull = ownerEmailOrNull;
                IsWritable = (privileges & (AccessPrivileges.Modify | AccessPrivileges.Create)) == (AccessPrivileges.Modify | AccessPrivileges.Create);
            }

            public Uri Uri { get; }
            public string Name { get; }
            public string OwnerEmailOrNull { get; }
            public bool IsOwn => OwnerEmailOrNull == null;
            public bool IsWritable { get; }
        }
    }
}
