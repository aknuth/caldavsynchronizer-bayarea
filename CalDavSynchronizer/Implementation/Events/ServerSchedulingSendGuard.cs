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
    /// Responses (accept, tentative, decline) to received meetings are handled the same way: the
    /// appointment is marked, the synchronization uploads the attendee's PARTSTAT without
    /// SCHEDULE-AGENT=CLIENT, and the server sends the REPLY. Counter proposals stay with Outlook,
    /// since an uploaded appointment can't express them and sabre/dav doesn't process COUNTER.
    /// </summary>
    public class ServerSchedulingSendGuard : IDisposable
    {
        private static readonly ILog s_logger = LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType);

        private const string PR_FINVITED = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B";
        private const string PR_RESPONSE_STATUS = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82180003";
        private const string PR_COUNTER_PROPOSAL = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8257000B";
        private const string PR_SERVER_SCHEDULED = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/CalDavSynchronizerServerScheduled";

        private readonly Application _application;
        private readonly NameSpace _session;
        private readonly Func<Options[]> _loadOptions;

        // Set while the guard itself hands a response back to Outlook, so ItemSend lets it pass.
        private bool _isSendingFallback;

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

        public static bool IsServerScheduled(AppointmentItem appointment) => GetBooleanPropertyOrFalse(appointment.PropertyAccessor, PR_SERVER_SCHEDULED);

        private static bool GetBooleanPropertyOrFalse(AppointmentItem appointment, string propertyName) => GetBooleanPropertyOrFalse(appointment.PropertyAccessor, propertyName);

        private static bool GetBooleanPropertyOrFalse(PropertyAccessor accessor, string propertyName)
        {
            try
            {
                using (var propertyAccessor = GenericComObjectWrapper.Create(accessor))
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
                if (_isSendingFallback)
                    return;

                var meetingItem = item as MeetingItem;
                if (meetingItem == null)
                    return;

                var meetingClass = meetingItem.Class;
                if (meetingClass == OlObjectClass.olMeetingResponsePositive
                    || meetingClass == OlObjectClass.olMeetingResponseTentative
                    || meetingClass == OlObjectClass.olMeetingResponseNegative)
                {
                    HandleResponse(meetingItem, meetingClass, ref cancel);
                    return;
                }

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

        private void HandleResponse(MeetingItem response, OlObjectClass responseClass, ref bool cancel)
        {
            // A counter proposal ("Propose New Time") is sent as a tentative or negative response. The uploaded
            // appointment keeps the original time, so the proposal would be lost, and sabre/dav doesn't know COUNTER.
            if (GetBooleanPropertyOrFalse(response.PropertyAccessor, PR_COUNTER_PROPOSAL))
            {
                s_logger.Info($"ItemSend: counter proposal for '{response.Subject}', Outlook sends.");
                return;
            }

            var appointment = response.GetAssociatedAppointment(false);
            if (appointment == null)
            {
                s_logger.Warn($"ItemSend: no appointment found for response '{response.Subject}', Outlook sends.");
                return;
            }

            OlMeetingStatus meetingStatus;
            try
            {
                meetingStatus = appointment.MeetingStatus;
            }
            catch (COMException x) when (IsMovedOrDeleted(x))
            {
                // Declining or deleting from the calendar, Outlook removes the appointment before sending. Nothing is left
                // to upload, so Outlook sends as before.
                s_logger.Info($"ItemSend: appointment of response '{response.Subject}' already removed by Outlook, Outlook sends.");
                return;
            }

            if (meetingStatus != OlMeetingStatus.olMeetingReceived)
            {
                s_logger.Info($"ItemSend: response for '{appointment.Subject}' with meeting status {meetingStatus}, Outlook sends.");
                return;
            }

            var profile = GetServerSchedulingProfileOrNull(appointment);
            if (profile == null)
                return;

            OlResponseStatus responseStatus;
            switch (responseClass)
            {
                case OlObjectClass.olMeetingResponsePositive:
                    responseStatus = OlResponseStatus.olResponseAccepted;
                    break;
                case OlObjectClass.olMeetingResponseTentative:
                    responseStatus = OlResponseStatus.olResponseTentative;
                    break;
                default:
                    responseStatus = OlResponseStatus.olResponseDeclined;
                    break;
            }

            s_logger.Info($"ItemSend: response {responseStatus} for '{appointment.Subject}', Outlook's status is {appointment.ResponseStatus}.");

            // Save before cancelling, while still inside ItemSend. If saving fails, Outlook sends as usual.
            // The response status is set explicitly, in case Outlook sets it only after sending.
            using (var propertyAccessor = GenericComObjectWrapper.Create(appointment.PropertyAccessor))
            {
                propertyAccessor.Inner.SetProperty(PR_SERVER_SCHEDULED, true);
                propertyAccessor.Inner.SetProperty(PR_RESPONSE_STATUS, (int) responseStatus);
            }

            appointment.Save();

            s_logger.Info($"ItemSend: suppressing {responseClass} for '{appointment.Subject}' (profile '{profile.Name}'), the server sends the REPLY instead.");
            cancel = true;

            // "Edit the response before sending" opens the response in an inspector, which has to be closed.
            var inspector = _application.ActiveInspector();
            var isResponseInspector = inspector?.CurrentItem is MeetingItem inspectorItem && inspectorItem.Class == responseClass;

            ComponentContainer.EnsureSynchronizationContext();
            SynchronizationContext.Current.Post(_ => FinishResponse(response, appointment, responseStatus, isResponseInspector ? inspector : null), null);
        }

        private void FinishResponse(MeetingItem response, AppointmentItem appointment, OlResponseStatus responseStatus, Inspector inspectorOrNull)
        {
            try
            {
                try
                {
                    inspectorOrNull?.Close(OlInspectorClose.olDiscard);
                }
                catch (Exception x)
                {
                    s_logger.Warn("ItemSend: could not close response inspector.", x);
                }

                // When declining, Outlook may delete the appointment despite the cancelled send. The synchronization
                // would then only delete it on the server, where it still has SCHEDULE-AGENT=CLIENT, so no REPLY.
                // In that case Outlook sends the response after all, as before.
                if (IsInCalendar(appointment))
                {
                    s_logger.Info($"ItemSend: appointment kept with response {appointment.ResponseStatus}, the synchronization uploads it.");
                    return;
                }

                s_logger.Warn($"ItemSend: Outlook removed the appointment after the {responseStatus} response, letting Outlook send it.");
                try
                {
                    _isSendingFallback = true;
                    response.Send();
                }
                finally
                {
                    _isSendingFallback = false;
                }
            }
            catch (Exception x)
            {
                s_logger.Error($"ItemSend: could not finish the {responseStatus} response. It may not have been sent.", x);
            }
            finally
            {
                if (inspectorOrNull != null)
                    Marshal.ReleaseComObject(inspectorOrNull);
                Marshal.ReleaseComObject(appointment);
                Marshal.ReleaseComObject(response);
            }
        }

        private bool IsInCalendar(AppointmentItem appointment)
        {
            try
            {
                return GetServerSchedulingProfileOrNull(appointment) != null;
            }
            catch (COMException x) when (IsMovedOrDeleted(x))
            {
                return false;
            }
        }

        // "The item has been moved or deleted", seen with both codes.
        private static bool IsMovedOrDeleted(COMException x) => x.ErrorCode == unchecked((int) 0x8004010A) || x.ErrorCode == unchecked((int) 0x8834010A);

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
            catch (COMException x) when (IsMovedOrDeleted(x))
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
                // Only profiles that upload Outlook's changes, otherwise the server never sees the meeting.
                && (o.SynchronizationMode == SynchronizationMode.MergeInBothDirections
                    || o.SynchronizationMode == SynchronizationMode.MergeOutlookIntoServer
                    || o.SynchronizationMode == SynchronizationMode.ReplicateOutlookIntoServer)
                && o.MappingConfiguration is EventMappingConfiguration mapping
                && mapping.MapAttendees
                && !mapping.ScheduleAgentClient
                && !mapping.SendNoAppointmentNotifications
                && o.OutlookFolderStoreId == folderStoreId
                && (o.OutlookFolderEntryId == folderEntryId || _session.CompareEntryIDs(o.OutlookFolderEntryId, folderEntryId)));
        }
    }
}
