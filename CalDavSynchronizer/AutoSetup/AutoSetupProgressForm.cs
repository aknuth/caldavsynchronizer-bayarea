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

using System.Drawing;
using System.Windows.Forms;
using CalDavSynchronizer.Globalization;

namespace CalDavSynchronizer.AutoSetup
{
    /// <summary>
    /// Shown while AccountAutoSetup queries the server, so the user knows to wait.
    /// </summary>
    public class AutoSetupProgressForm : Form
    {
        public AutoSetupProgressForm()
        {
            Text = Strings.Get($"Set up calendars");
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 90);
            UseWaitCursor = true;

            Controls.Add(new Label
            {
                Text = Strings.Get($"Searching for calendars and address books on the server..."),
                Location = new Point(16, 16),
                Size = new Size(328, 20)
            });
            Controls.Add(new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Location = new Point(16, 46),
                Size = new Size(328, 20)
            });
        }
    }
}
