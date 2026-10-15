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

namespace GenSync.EntityRepositories
{
    /// <summary>
    /// Thrown by a repository's Create when the entity is deliberately not created. The entity is reported with a
    /// warning instead of an error, and the create is tried again on the next synchronization.
    /// </summary>
    public class EntityCreateSkippedException : Exception
    {
        public EntityCreateSkippedException(string message)
            : base(message)
        {
        }
    }
}
