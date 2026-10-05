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
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using CalDavSynchronizer.Contracts;
using CalDavSynchronizer.Implementation.ComWrappers;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace CalDavSynchronizer.Implementation.Events
{
    /// <summary>
    /// Prevents Outlook from sending iTIP mails (meeting requests and cancellations) for meetings
    /// in folders whose profile leaves scheduling to the CalDAV server (RFC 6638), i.e. profiles
    /// that map attendees and neither set SCHEDULE-AGENT=CLIENT nor X-SOGO-SEND-APPOINTMENT-NOTIFICATIONS.
    /// The send is cancelled, the appointment is saved (or deleted on cancellation) and the
    /// synchronization uploads it, so the server sends the invitation exactly once.
    /// </summary>
    public class ServerSchedulingSendGuard : IDisposable
    {
        private static readonly ILog s_logger = LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType);

        private const string PR_FINVITED = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B";

        private readonly Application _application;
        private readonly NameSpace _session;
        private readonly Func<Options[]> _loadOptions;

        public ServerSchedulingSendGuard(Application application, NameSpace session, Func<Options[]> loadOptions)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _loadOptions = loadOptions ?? throw new ArgumentNullException(nameof(loadOptions));

            ((ApplicationEvents_11_Event) _application).ItemSend += Application_ItemSend;
        }

        public void Dispose()
        {
            ((ApplicationEvents_11_Event) _application).ItemSend -= Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            try
            {
                var meetingItem = item as MeetingItem;
                if (meetingItem == null)
                    return;

                var meetingClass = meetingItem.Class;
                if (meetingClass != OlObjectClass.olMeetingRequest && meetingClass != OlObjectClass.olMeetingCancellation)
                    return;

                // Prefer the appointment of the open inspector, so it can be saved and closed without
                // a conflict, but only if it is the meeting being sent.
                var inspector = _application.ActiveInspector();
                var inspectorAppointment = inspector?.CurrentItem as AppointmentItem;
                var associatedAppointment = meetingItem.GetAssociatedAppointment(false);
                if (inspectorAppointment != null && associatedAppointment != null
                    && inspectorAppointment.GlobalAppointmentID != associatedAppointment.GlobalAppointmentID)
                {
                    inspectorAppointment = null;
                    inspector = null;
                }

                var appointment = inspectorAppointment ?? associatedAppointment;
                if (appointment == null)
                {
                    s_logger.Warn($"ItemSend: no appointment found for '{meetingItem.Subject}', Outlook sends.");
                    return;
                }

                // Only meetings we organize. Forwarding a received invitation is left to Outlook.
                var meetingStatus = appointment.MeetingStatus;
                if (meetingStatus != OlMeetingStatus.olMeeting && meetingStatus != OlMeetingStatus.olMeetingCanceled)
                    return;

                var profile = GetServerSchedulingProfileOrNull(appointment);
                if (profile == null)
                    return;

                s_logger.Info($"ItemSend: suppressing {meetingClass} for '{appointment.Subject}' (profile '{profile.Name}'), the server sends instead.");
                cancel = true;

                // Closing the inspector or deleting the item inside ItemSend is not allowed, so defer it.
                ComponentContainer.EnsureSynchronizationContext();
                var isWholeMeetingCancelled = meetingStatus == OlMeetingStatus.olMeetingCanceled;
                var appointmentInspector = inspectorAppointment != null ? inspector : null;
                SynchronizationContext.Current.Post(_ => CompleteSuppressedSend(appointment, appointmentInspector, isWholeMeetingCancelled), null);
            }
            catch (Exception x)
            {
                // Never break sending because of the guard. If in doubt, Outlook sends.
                s_logger.Error("ItemSend: error in send guard, Outlook sends.", x);
            }
        }

        private void CompleteSuppressedSend(AppointmentItem appointment, Inspector inspectorOrNull, bool isWholeMeetingCancelled)
        {
            try
            {
                if (isWholeMeetingCancelled)
                {
                    // Deleting the meeting makes the synchronization delete it on the server,
                    // which sends the CANCEL to the attendees.
                    inspectorOrNull?.Close(OlInspectorClose.olDiscard);
                    appointment.Delete();
                    s_logger.Info("ItemSend: cancelled meeting deleted, server sends CANCEL.");
                }
                else
                {
                    // Mark the invitations as sent, so Outlook treats the meeting as sent
                    // ("Send Update" on later changes) instead of "invitations have not been sent".
                    using (var propertyAccessor = GenericComObjectWrapper.Create(appointment.PropertyAccessor))
                    {
                        propertyAccessor.Inner.SetProperty(PR_FINVITED, true);
                    }

                    appointment.Save();
                    if (inspectorOrNull != null)
                        inspectorOrNull.Close(OlInspectorClose.olSave);
                    s_logger.Info("ItemSend: meeting saved, synchronization uploads it and the server sends.");
                }
            }
            catch (Exception x)
            {
                s_logger.Error("ItemSend: error completing suppressed send.", x);
            }
            finally
            {
                if (inspectorOrNull != null)
                    Marshal.ReleaseComObject(inspectorOrNull);
                Marshal.ReleaseComObject(appointment);
            }
        }

        private Options GetServerSchedulingProfileOrNull(AppointmentItem appointment)
        {
            string folderEntryId;
            string folderStoreId;
            var folder = (Folder) appointment.Parent;
            try
            {
                folderEntryId = folder.EntryID;
                folderStoreId = folder.StoreID;
            }
            finally
            {
                Marshal.ReleaseComObject(folder);
            }

            return _loadOptions().FirstOrDefault(o =>
                !o.Inactive
                && o.MappingConfiguration is EventMappingConfiguration mapping
                && mapping.MapAttendees
                && !mapping.ScheduleAgentClient
                && !mapping.SendNoAppointmentNotifications
                && o.OutlookFolderStoreId == folderStoreId
                && (o.OutlookFolderEntryId == folderEntryId || _session.CompareEntryIDs(o.OutlookFolderEntryId, folderEntryId)));
        }
    }
}
