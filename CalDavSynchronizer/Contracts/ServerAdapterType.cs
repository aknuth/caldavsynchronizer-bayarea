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
using System.Xml.Serialization;

namespace CalDavSynchronizer.Contracts
{
    public enum ServerAdapterType
    {
        [XmlEnum("Default")] WebDavHttpClientBased,

        // The Google adapters are removed. The values stay so that options files containing them can still be
        // deserialized; such profiles fail with an error when synchronizing instead of breaking the add-in's startup.

        [XmlEnum("GoogleOAuth")] WebDavHttpClientBasedWithGoogleOAuth,

        GoogleTaskApi,

        GoogleContactApi
    }
}