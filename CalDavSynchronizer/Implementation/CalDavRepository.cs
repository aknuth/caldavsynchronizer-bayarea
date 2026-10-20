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
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using CalDavSynchronizer.Contracts;
using CalDavSynchronizer.DataAccess;
using CalDavSynchronizer.DDayICalWorkaround;
using CalDavSynchronizer.Diagnostics;
using CalDavSynchronizer.Implementation.TimeRangeFiltering;
using CalDavSynchronizer.Utilities;
using DDay.iCal;
using DDay.iCal.Serialization;
using DDay.iCal.Serialization.iCalendar;
using GenSync;
using GenSync.EntityRepositories;
using GenSync.Logging;
using GenSync.Utilities;
using log4net;

namespace CalDavSynchronizer.Implementation
{
    public static class CalDavRepository
    {
        public enum EntityType
        {
            Event,
            Todo
        }
    }

    public class CalDavRepository<TContext> : IEntityRepository<WebResourceName, string, IICalendar, TContext>, IStateAwareEntityRepository<WebResourceName, string, TContext, string>
    {
        private static readonly ILog s_logger = LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType);

        private readonly ICalDavDataAccess _calDavDataAccess;
        private readonly IStringSerializer _calendarSerializer;
        private readonly CalDavRepository.EntityType _entityType;
        private readonly IDateTimeRangeProvider _dateTimeRangeProvider;
        private readonly bool _deleteAndCreateOnUpdateError403;
        private readonly IEqualityComparer<string> _versionComparer;
        private readonly string _userAddressOrNull;
        private readonly bool _isSharedCollection;


        public CalDavRepository(
            ICalDavDataAccess calDavDataAccess,
            IStringSerializer calendarSerializer,
            CalDavRepository.EntityType entityType,
            IDateTimeRangeProvider dateTimeRangeProvider,
            bool deleteAndCreateOnUpdateError403,
            IEqualityComparer<string> versionComparer,
            string userAddressOrNull = null,
            bool isSharedCollection = false)
        {
            if (versionComparer == null) throw new ArgumentNullException(nameof(versionComparer));
            _deleteAndCreateOnUpdateError403 = deleteAndCreateOnUpdateError403;
            _versionComparer = versionComparer;
            _calDavDataAccess = calDavDataAccess;
            _calendarSerializer = calendarSerializer;
            _entityType = entityType;
            _dateTimeRangeProvider = dateTimeRangeProvider;
            _userAddressOrNull = userAddressOrNull;
            _isSharedCollection = isSharedCollection;
        }

        public async Task<IEnumerable<EntityVersion<WebResourceName, string>>> GetVersions(IEnumerable<IdWithAwarenessLevel<WebResourceName>> idsOfEntitiesToQuery, TContext context, IGetVersionsLogger logger)
        {
            return await _calDavDataAccess.GetVersions(idsOfEntitiesToQuery.Select(i => i.Id));
        }

        public async Task<IEnumerable<EntityVersion<WebResourceName, string>>> GetAllVersions(IEnumerable<WebResourceName> idsOfknownEntities, TContext context, IGetVersionsLogger logger)
        {
            using (AutomaticStopwatch.StartInfo(s_logger, "CalDavRepository.GetVersions"))
            {
                switch (_entityType)
                {
                    case CalDavRepository.EntityType.Event:

                        return await _calDavDataAccess.GetEventVersions(_dateTimeRangeProvider.GetRange());
                    case CalDavRepository.EntityType.Todo:
                        return await _calDavDataAccess.GetTodoVersions(_dateTimeRangeProvider.GetRange());
                    default:
                        throw new NotImplementedException(string.Format("EntityType '{0}' not implemented.", _entityType));
                }
            }
        }

        public async Task<IEnumerable<EntityWithId<WebResourceName, IICalendar>>> Get(ICollection<WebResourceName> ids, ILoadEntityLogger logger, TContext context)
        {
            if (ids.Count == 0)
                return new EntityWithId<WebResourceName, IICalendar>[] { };

            using (AutomaticStopwatch.StartInfo(s_logger, string.Format("CalDavRepository.Get ({0} entitie(s))", ids.Count)))
            {
                var entities = await _calDavDataAccess.GetEntities(ids);
                return await ParallelDeserialize(entities, logger);
            }
        }

        public Task VerifyUnknownEntities(Dictionary<WebResourceName, string> unknownEntites, TContext context)
        {
            return Task.FromResult(0);
        }

        public void Cleanup(IICalendar entity)
        {
            // nothing to do
        }

        public void Cleanup(IEnumerable<IICalendar> entities)
        {
            // nothing to do
        }

        private Task<IReadOnlyList<EntityWithId<WebResourceName, IICalendar>>> ParallelDeserialize(IReadOnlyList<EntityWithId<WebResourceName, string>> serializedEntities, ILoadEntityLogger logger)
        {
            return Task.Factory.StartNew(() =>
            {
                var result = new List<EntityWithId<WebResourceName, IICalendar>>();

                Parallel.ForEach(
                    serializedEntities,
                    () => Tuple.Create(new iCalendarSerializer(), new List<Tuple<WebResourceName, IICalendar>>()),
                    (serialized, loopState, threadLocal) =>
                    {
                        IICalendar calendar;
                        string normalizedICalData, fixedICalData;

                        // fix some linebreak issues with Open-Xchange
                        if (serialized.Entity.Contains("\r\r\n"))
                        {
                            normalizedICalData = CalendarDataPreprocessor.NormalizeLineBreaks(serialized.Entity);
                        }
                        else
                        {
                            normalizedICalData = serialized.Entity;
                        }

                        // emClient sets DTSTART in VTIMEZONE to year 0001, which causes a 90 sec delay in DDay.iCal to evaluate the recurrence rule.
                        // If we find such a DTSTART we replace it 0001 with 1970 since the historic data is not valid anyway and avoid the performance issue.
                        if (normalizedICalData.Contains("DTSTART:00010101"))
                        {
                            fixedICalData = CalendarDataPreprocessor.FixInvalidDTSTARTInTimeZoneNoThrow(normalizedICalData);
                            s_logger.InfoFormat("Changed DTSTART from year 0001 to 1970 in VTIMEZONE of ICalData '{0}'.", serialized.Id);
                        }
                        else
                        {
                            fixedICalData = normalizedICalData;
                        }

                        if (TryDeserializeCalendar(fixedICalData, out calendar, serialized.Id, threadLocal.Item1, NullLoadEntityLogger.Instance /* use null logger to ignore errors on first deserialization try */))
                        {
                            // Add only if there is atleast one vevent or vtodo to avoid Nullreference Exceptions when processing
                            if ((_entityType == CalDavRepository.EntityType.Event && calendar.Events.Count > 0) ||
                                (_entityType == CalDavRepository.EntityType.Todo && calendar.Todos.Count > 0))
                            {
                                threadLocal.Item2.Add(Tuple.Create(serialized.Id, calendar));
                            }
                        }
                        else
                        {
                            // maybe deserialization failed because of the iCal-TimeZone-Bug =>  try to fix it
                            var fixedICalData2 = CalendarDataPreprocessor.FixTimeZoneComponentOrderNoThrow(fixedICalData);
                            if (TryDeserializeCalendar(fixedICalData2, out calendar, serialized.Id, threadLocal.Item1, logger))
                            {
                                // Add only if there is atleast one vevent or vtodo to avoid Nullreference Exceptions when processing
                                if ((_entityType == CalDavRepository.EntityType.Event && calendar.Events.Count > 0) ||
                                    (_entityType == CalDavRepository.EntityType.Todo && calendar.Todos.Count > 0))
                                {
                                    threadLocal.Item2.Add(Tuple.Create(serialized.Id, calendar));
                                    s_logger.Info(string.Format("Deserialized ICalData with reordering of TimeZone data '{0}'.", serialized.Id));
                                }
                            }
                        }

                        return threadLocal;
                    },
                    threadLocal =>
                    {
                        lock (result)
                        {
                            foreach (var calendar in threadLocal.Item2)
                                result.Add(EntityWithId.Create(calendar.Item1, calendar.Item2));
                        }
                    });

                IReadOnlyList<EntityWithId<WebResourceName, IICalendar>> readOnlyResult = result;
                return readOnlyResult;
            });
        }


        public async Task<bool> TryDelete(WebResourceName entityId, string version, TContext context, IEntitySynchronizationLogger logger)
        {
            using (AutomaticStopwatch.StartDebug(s_logger))
            {
                if (_userAddressOrNull != null && _entityType == CalDavRepository.EntityType.Event)
                {
                    var calendar = await GetCalendarOrNull(entityId);
                    if (calendar != null && IsMeetingOrganizedByOthers(calendar))
                    {
                        if (_isSharedCollection)
                        {
                            // In a calendar shared by someone else, Outlook removes a meeting the user declines or deletes. On the
                            // server that would delete it for everybody. Leave it there: the relation is dropped, so the next
                            // synchronization brings it back to Outlook.
                            s_logger.Info($"Not deleting '{entityId.OriginalAbsolutePath}': a meeting organized by someone else in a shared calendar.");
                            logger.LogWarning("Not deleted on the server: a meeting organized by someone else, in a shared calendar. It comes back in Outlook with the next synchronization.");
                            return true;
                        }

                        // The user's own copy of someone else's meeting. Without SCHEDULE-AGENT=CLIENT, sabre/dav sends a REPLY
                        // DECLINED when it is deleted. But Outlook removes the appointment itself when declining from the
                        // invitation mail and sends the decline (the send guard can't read the removed appointment), or the
                        // user chose not to send a response. So silence the server first.
                        version = await SetScheduleAgentClientOrNull(entityId, version, calendar) ?? version;
                    }
                }

                return await _calDavDataAccess.TryDeleteEntity(entityId, version);
            }
        }

        private async Task<IICalendar> GetCalendarOrNull(WebResourceName entityId)
        {
            var entities = await _calDavDataAccess.GetEntities(new[] {entityId});
            if (entities.Count == 0)
                return null;

            try
            {
                return DeserializeCalendar(entities[0].Entity, _calendarSerializer);
            }
            catch (Exception x)
            {
                s_logger.Error($"Could not deserialize '{entityId.OriginalAbsolutePath}', deleting it as it is.", x);
                return null;
            }
        }

        private bool IsMeetingOrganizedByOthers(IICalendar calendar)
        {
            var organizer = calendar.Events.FirstOrDefault()?.Organizer?.Value?.ToString();
            if (string.IsNullOrEmpty(organizer))
                return false;

            const string mailtoPrefix = "mailto:";
            if (organizer.StartsWith(mailtoPrefix, StringComparison.OrdinalIgnoreCase))
                organizer = organizer.Substring(mailtoPrefix.Length);

            return !string.Equals(organizer, _userAddressOrNull, StringComparison.OrdinalIgnoreCase);
        }

        /// <returns>The new version, or null if nothing was changed or the update failed.</returns>
        private async Task<string> SetScheduleAgentClientOrNull(WebResourceName entityId, string version, IICalendar calendar)
        {
            var changed = false;
            foreach (var evt in calendar.Events.Where(e => e.Organizer != null))
            {
                // Already CLIENT, or another explicit value (SERVER, NONE) that is left as it is.
                if (evt.Organizer.Parameters.ContainsKey("SCHEDULE-AGENT"))
                    continue;

                evt.Organizer.Parameters.Add("SCHEDULE-AGENT", "CLIENT");
                changed = true;
            }

            if (!changed)
                return null;

            try
            {
                var result = await _calDavDataAccess.TryUpdateEntity(entityId, version, SerializeCalendar(calendar));
                s_logger.Info($"Set SCHEDULE-AGENT=CLIENT on '{entityId.OriginalAbsolutePath}' before deleting it, so the server sends no decline.");
                return result?.Version;
            }
            catch (Exception x)
            {
                s_logger.Error($"Could not set SCHEDULE-AGENT=CLIENT on '{entityId.OriginalAbsolutePath}' before deleting it.", x);
                return null;
            }
        }

        public async Task<EntityVersion<WebResourceName, string>> TryUpdate(
            WebResourceName entityId,
            string entityVersion,
            IICalendar entityToUpdate,
            Func<IICalendar, Task<IICalendar>> entityModifier,
            TContext context,
            IEntitySynchronizationLogger logger)
        {
            using (AutomaticStopwatch.StartDebug(s_logger))
            {
                var updatedEntity = await entityModifier(entityToUpdate);
                try
                {
                    return await _calDavDataAccess.TryUpdateEntity(entityId, entityVersion, SerializeCalendar(updatedEntity));
                }
                catch (WebDavClientException ex)
                {
                    if (_deleteAndCreateOnUpdateError403 && ex.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        s_logger.Warn("Server returned '403' ('Forbidden') for update, trying Delete and Recreate instead...");
                        logger.LogWarning("Server returned '403' ('Forbidden') for update, trying Delete and Recreate instead...");

                        await TryDelete(entityId, entityVersion, context, logger);

                        var uid = Guid.NewGuid().ToString();
                        if (updatedEntity.Events.Count > 0)
                            updatedEntity.Events[0].UID = uid;
                        else
                            updatedEntity.Todos[0].UID = uid;

                        return await _calDavDataAccess.CreateEntity(SerializeCalendar(updatedEntity), uid);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        public async Task<EntityVersion<WebResourceName, string>> Create(Func<IICalendar, Task<IICalendar>> entityInitializer, TContext context)
        {
            using (AutomaticStopwatch.StartDebug(s_logger))
            {
                IICalendar newCalendar = new iCalendar();
                newCalendar = await entityInitializer(newCalendar);
                var uid = newCalendar.Events.Count > 0 ? newCalendar.Events[0].UID : newCalendar.Todos[0].UID;
                var name = IsUsableAsResourceName(uid) ? uid : Guid.NewGuid().ToString();
                try
                {
                    return await _calDavDataAccess.CreateEntity(SerializeCalendar(newCalendar), name);
                }
                catch (ResourceAlreadyExistsException x)
                {
                    IICalendar existingCalendar;
                    try
                    {
                        existingCalendar = DeserializeCalendar(x.ExistingContent, _calendarSerializer);
                    }
                    catch (Exception deserializationException)
                    {
                        s_logger.Error("Could not deserialize the existing resource.", deserializationException);
                        throw x;
                    }

                    var existingUid = existingCalendar.Events.Count > 0 ? existingCalendar.Events[0].UID : existingCalendar.Todos.FirstOrDefault()?.UID;
                    if (existingUid != uid)
                        throw;

                    // Typically an invitation between two accounts of the same Outlook profile: Outlook puts the
                    // attendee's copy into the default calendar next to the organizer's copy, both with the same UID.
                    // Uploading it would overwrite the organizer's copy, so leave it in Outlook and just warn.
                    var description = newCalendar.Events.Count > 0
                        ? $"'{newCalendar.Events[0].Summary}' on {newCalendar.Events[0].Start?.Value:g}"
                        : $"'{newCalendar.Todos[0].Summary}'";
                    s_logger.Warn($"Not creating {description}, the server already has an entity with UID '{uid}'.");
                    throw new EntityCreateSkippedException(
                        $"Not uploaded: {description} is already on the server, with the same UID. " +
                        "This Outlook item is a second copy of it, usually an invitation between two accounts in this Outlook " +
                        "that Outlook put into this calendar. Delete the copy you don't need (choose 'Don't send a response').");
                }
            }
        }


        /// <summary>
        /// Whether the UID can be the resource name: a plain path segment that fits sabre/dav's varchar(200) column
        /// with ".ics". A UID can be anything, e.g. a URL (Adobe Connect), which would replace the server's URL.
        /// </summary>
        private static bool IsUsableAsResourceName(string uid)
        {
            const int maximumNameLength = 196;
            return !string.IsNullOrEmpty(uid)
                   && uid.Length <= maximumNameLength
                   && uid.All(c => (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == '@')
                   && uid.Trim('.').Length > 0;
        }

        private string SerializeCalendar(IICalendar calendar)
        {
            return _calendarSerializer.SerializeToString(calendar);
        }

        private static bool TryDeserializeCalendar(
            string iCalData,
            out IICalendar calendar,
            WebResourceName uriOfCalendarForLogging,
            IStringSerializer calendarSerializer,
            ILoadEntityLogger logger)
        {
            calendar = null;
            try
            {
                calendar = DeserializeCalendar(iCalData, calendarSerializer);
                return true;
            }
            catch (Exception x)
            {
                s_logger.Error(string.Format("Could not deserialize ICalData of '{0}'.", uriOfCalendarForLogging.OriginalAbsolutePath));
                s_logger.Debug(string.Format("ICalData:\r\n{0}", iCalData), x);
                logger.LogSkipLoadBecauseOfError(uriOfCalendarForLogging, x);
                return false;
            }
        }

        private static IICalendar DeserializeCalendar(string iCalData, IStringSerializer calendarSerializer)
        {
            using (var reader = new StringReader(iCalData))
            {
                var calendarCollection = (iCalendarCollection) calendarSerializer.Deserialize(reader);
                return calendarCollection[0];
            }
        }

        public async Task<(IEntityStateCollection<WebResourceName, string> States, string NewToken)> GetFullRepositoryState(IEnumerable<WebResourceName> idsOfknownEntities, string stateToken, TContext context, IGetVersionsLogger logger)
        {
            var collectionSyncResult = await _calDavDataAccess.CollectionSync(stateToken, logger);
            return (new WebDavCollectionSyncEntityStates(collectionSyncResult.ChangedOrAddedItems, collectionSyncResult.DeletedItems, _versionComparer), collectionSyncResult.SyncToken);
        }
    }
}