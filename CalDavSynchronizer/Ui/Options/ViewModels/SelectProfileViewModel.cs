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
using System.Windows.Input;
using CalDavSynchronizer.Globalization;
using CalDavSynchronizer.ProfileTypes;

namespace CalDavSynchronizer.Ui.Options.ViewModels
{
    class SelectProfileViewModel
    {
        public event EventHandler<CloseEventArgs> CloseRequested;
        private readonly IUiService _uiService;

        public SelectProfileViewModel(IReadOnlyCollection<IProfileType> profileTypes, IUiService uiService)
        {
            if (profileTypes == null) throw new ArgumentNullException(nameof(profileTypes));
            if (uiService == null) throw new ArgumentNullException(nameof(uiService));

            _uiService = uiService;
            ProfileTypes = profileTypes.Select(p => new ProfileViewModel(p)).ToList();
            ProfileTypes.First().IsSelected = true;

            OkCommand = new DelegateCommand(_ => Close(true));
            CancelCommand = new DelegateCommand(_ => Close(false));
        }

        public ICommand CancelCommand { get; }
        public ICommand OkCommand { get; }
        public IReadOnlyList<ProfileViewModel> ProfileTypes { get; }
        public IProfileType SelectedProfile { get; private set; }

        void Close(bool okPressed)
        {
            if (okPressed)
            {
                if (ProfileTypes.Count(p => p.IsSelected) != 1)
                {
                    _uiService.ShowErrorDialog(Strings.Get($"Please select exactly one option."), ComponentContainer.MessageBoxTitle);
                    return;
                }

                SelectedProfile = ProfileTypes.Single(p => p.IsSelected).ProfileType;
            }

            CloseRequested?.Invoke(this, new CloseEventArgs(okPressed));
        }

        public static SelectProfileViewModel DesignInstance => new SelectProfileViewModel(
            new[]
                {
                    "Generic CalDAV_CardDAV"
                }
                .Select(u => new DesignProfileType(Path.GetFileName(u), u))
                .ToArray(),
            NullUiService.Instance);
    }
}