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
using CalDavSynchronizer.DataAccess;
using Newtonsoft.Json.Linq;

namespace CalDavSynchronizer.AutomaticUpdates
{
    internal class AvailableVersionService : IAvailableVersionService
    {
        private Uri _downloadLink;
        private string _notes;

        /// <summary>
        /// Reads WebResourceUrls.SiteContainingNewestVersion, a JSON file on our release server:
        /// { "version": "5.0.1", "url": "https://.../setup.exe", "notes": "optional text" }
        /// </summary>
        public Version GetVersionOfDefaultDownload()
        {
            string site;

            using (var client = HttpUtility.CreateWebClient())
            {
                site = client.DownloadString(WebResourceUrls.SiteContainingNewestVersion);
            }

            var latest = JObject.Parse(site);

            _downloadLink = latest["url"] is JValue urlValue && Uri.TryCreate(urlValue.Value<string>(), UriKind.Absolute, out var url)
                ? url
                : null;
            _notes = (latest["notes"] as JValue)?.Value<string>();

            return latest["version"] is JValue versionValue && Version.TryParse(versionValue.Value<string>(), out var version)
                ? version
                : null;
        }

        public string GetWhatsNewNoThrow(Version oldVersion, Version newVersion) => _notes ?? string.Empty;

        public Uri DownloadLink => _downloadLink;
    }
}