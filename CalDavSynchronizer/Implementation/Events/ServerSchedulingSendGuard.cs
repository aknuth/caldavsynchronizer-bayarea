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
    /// The send is cancelled, the appointment is saved and the synchronization uploads it, so the
    /// server sends the invitation exactly once. On cancellation Outlook deletes the appointment
    /// itself and the synchronization deletes it on the server, which sends the CANCEL.
    /// Only new meetings and meetings already handed to the server are taken over; meetings that
    /// Outlook has sent itself stay with Outlook (and keep SCHEDULE-AGENT=CLIENT on the server).
    /// </summary>
    public class ServerSchedulingSendGuard : IDisposable
    {
        private static readonly ILog s_logger = LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType);

        private const string PR_FINVITED = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B";
        private const string PR_SERVER_SCHEDULED = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/CalDavSynchronizerServerScheduled";

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

        public static bool IsServerScheduled(AppointmentItem appointment) => GetBooleanPropertyOrFalse(appointment, PR_SERVER_SCHEDULED);

        private static bool GetBooleanPropertyOrFalse(AppointmentItem appointment, string propertyName)
        {
            try
            {
                using (var propertyAccessor = GenericComObjectWrapper.Create(appointment.PropertyAccessor))
                {
                    return propertyAccessor.Inner.GetProperty(propertyName) is bool value && value;
                }
            }
            catch (COMException)
            {
                // Property not set
                return false;
            }
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

                // Meetings Outlook has already sent itself stay with Outlook, otherwise attendees would get
                // updates from Outlook and the server.
                var isNewMeeting = !GetBooleanPropertyOrFalse(appointment, PR_FINVITED);
                if (!isNewMeeting && !IsServerScheduled(appointment))
                    return;

                var profile = GetServerSchedulingProfileOrNull(appointment);
                if (profile == null)
                    return;

                var isWholeMeetingCancelled = meetingStatus == OlMeetingStatus.olMeetingCanceled;
                if (!isWholeMeetingCancelled)
                {
                    // Save before cancelling, while still inside ItemSend. If saving fails, Outlook sends as usual.
                    // The marker hands the scheduling of the meeting to the server.
                    // FInvited makes Outlook treat the meeting as sent ("Send Update" on later changes). With an
                    // open inspector it is set only after closing it: set before, Outlook sees changes since
                    // the "sent" state when closing and asks "save changes and send update".
                    using (var propertyAccessor = GenericComObjectWrapper.Create(appointment.PropertyAccessor))
                    {
                        propertyAccessor.Inner.SetProperty(PR_SERVER_SCHEDULED, true);
                        if (inspectorAppointment == null)
                            propertyAccessor.Inner.SetProperty(PR_FINVITED, true);
                    }

                    appointment.Save();
                }

                s_logger.Info($"ItemSend: suppressing {meetingClass} for '{appointment.Subject}' (profile '{profile.Name}'), the server sends instead.");
                cancel = true;

                // Closing the inspector or deleting inside ItemSend is not allowed, so defer it.
                ComponentContainer.EnsureSynchronizationContext();
                if (isWholeMeetingCancelled)
                {
                    var cancelledInspector = inspectorAppointment != null ? inspector : null;
                    SynchronizationContext.Current.Post(_ => DeleteCancelledMeeting(appointment, cancelledInspector), null);
                }
                else if (inspectorAppointment != null)
                {
                    SynchronizationContext.Current.Post(_ => CloseInspectorAndMarkInvited(inspector, appointment), null);
                }
            }
            catch (Exception x)
            {
                // Never break sending because of the guard. If in doubt, Outlook sends.
                s_logger.Error("ItemSend: error in send guard, Outlook sends.", x);
            }
        }

        private void DeleteCancelledMeeting(AppointmentItem appointment, Inspector inspectorOrNull)
        {
            // Deleting from the calendar view, Outlook deletes the appointment itself despite the cancelled send.
            // Cancelling from the open meeting, it keeps it. The synchronization then deletes it on the server,
            // which sends the CANCEL.
            try
            {
                inspectorOrNull?.Close(OlInspectorClose.olDiscard);
                appointment.Delete();
                s_logger.Info("ItemSend: cancelled meeting deleted, server sends CANCEL.");
            }
            catch (COMException x) when (x.ErrorCode == unchecked((int) 0x8004010A))
            {
                s_logger.Info("ItemSend: cancelled meeting already deleted by Outlook, server sends CANCEL.");
            }
            catch (Exception x)
            {
                s_logger.Error("ItemSend: could not delete cancelled meeting.", x);
            }
            finally
            {
                if (inspectorOrNull != null)
                    Marshal.ReleaseComObject(inspectorOrNull);
                Marshal.ReleaseComObject(appointment);
            }
        }

        private void CloseInspectorAndMarkInvited(Inspector inspector, AppointmentItem appointment)
        {
            try
            {
                // Already saved in ItemSend.
                inspector.Close(OlInspectorClose.olDiscard);
            }
            catch (Exception x)
            {
                // The user may have closed it already.
                s_logger.Warn("ItemSend: could not close inspector.", x);
            }

            try
            {
                using (var propertyAccessor = GenericComObjectWrapper.Create(appointment.PropertyAccessor))
                {
                    propertyAccessor.Inner.SetProperty(PR_FINVITED, true);
                }

                appointment.Save();
            }
            catch (Exception x)
            {
                s_logger.Warn("ItemSend: could not mark meeting as invited.", x);
            }
            finally
            {
                Marshal.ReleaseComObject(inspector);
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
