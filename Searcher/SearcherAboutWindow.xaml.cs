// <copyright file="SearcherAboutWindow.xaml.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
{
    /*
     * Searcher - Utility to search file content
     * Copyright (C) 2018  Dennis Joseph
     * 
     * This file is part of Searcher.

     * Searcher is free software: you can redistribute it and/or modify
     * it under the terms of the GNU General Public License as published by
     * the Free Software Foundation, either version 3 of the License, or
     * (at your option) any later version.
     * 
     * Searcher is distributed in the hope that it will be useful,
     * but WITHOUT ANY WARRANTY; without even the implied warranty of
     * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     * GNU General Public License for more details.
     * 
     * You should have received a copy of the GNU General Public License
     * along with Searcher.  If not, see <https://www.gnu.org/licenses/>.
     */

    using System;
    using System.Diagnostics;
    using System.Net;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Documents;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for SearcherAbout window.
    /// </summary>
    public partial class SearcherAboutWindow : Window
    {
        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="SearcherAboutWindow"/> class.
        /// </summary>
        public SearcherAboutWindow()
        {
            this.InitializeComponent();
            this.InitialiseControls();
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether a preference file exists.
        /// </summary>
        public bool AppHasPreferencesFile
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether checking for updates can be done.
        /// </summary>
        public bool CanCheckForUpdates
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the window is closing for application update.
        /// </summary>
        public bool ClosingForUpdate
        {
            get;
            set;
        }

        #endregion Public Properties

        #region Private Methods

        /// <summary>
        /// Save any config changes.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void AboutSearcher_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.CanCheckForUpdates = this.ChkUpdatesCheck.IsChecked.GetValueOrDefault() == true;
        }

        /// <summary>
        /// Determine window display.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void AboutSearcher_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.AppHasPreferencesFile)
            {
                this.ChkUpdatesCheck.IsChecked = this.CanCheckForUpdates;
            }
            else
            {
                this.SetupAppWithoutPreferences();
            }
        }

        /// <summary>
        /// Close window on pressing the escape key.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void AboutSearcher_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        /// <summary>
        /// Update the app.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void BtnUpdateSearcher_Click(object sender, RoutedEventArgs e)
        {
            if (Common.ApplicationUpdateExists)
            {
                // TODO: Use System.AppContext.BaseDirectory instead of AppDomain.CurrentDomain.BaseDirectory?
                string newProgPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg");
                string pathToOldProg = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Common.ApplicationExecutableName);
                string pathToNewProg = System.IO.Path.Combine(newProgPath, Common.ApplicationExecutableName);

                System.IO.File.Move(pathToOldProg, pathToOldProg + ".old");

                // Move content from update directory to current directory.
                System.IO.File.Move(pathToNewProg, pathToOldProg);
                System.IO.File.Copy(System.IO.Path.Combine(newProgPath, "COPYING.txt"), System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "COPYING.txt"), true);
                System.IO.File.Copy(System.IO.Path.Combine(newProgPath, "README.txt"), System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "README.txt"), true);

                // Delete update directory and close application for restart.
                System.IO.Directory.Delete(newProgPath, true);
                this.ClosingForUpdate = true;
                this.Close();
            }
        }

        /// <summary>
        /// Navigate to the URL specified.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">Used to get the URL to navigate to.</param>
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        /// <summary>
        /// Initialise controls to display on the form.
        /// </summary>
        private void InitialiseControls()
        {
            this.SetContentBasedOnLanguage();

            this.TblkVersionNumber.Text = Common.VersionNumber;
            string yearInfo = DateTime.Today.Year > 2018
                ? "2018 - " + DateTime.Today.Year
                : "2018";

            string copyRightText = string.Format("{0} {1}  {2}{3}", Application.Current.Resources["Copyright"].ToString(), yearInfo, "Dennis Joseph", Environment.NewLine);
            string warrantyText = Application.Current.Resources["WarrantyText"].ToString() + Environment.NewLine;

            Run copyRightAndWarranty = new (string.Join(Environment.NewLine, new string[] { copyRightText, warrantyText }));
            Run run2 = new (Application.Current.Resources["DistributeSoftware"].ToString() + " ");
            Run run4 = new (" " + Application.Current.Resources["ForDetails"].ToString());

            Hyperlink hyperlink = new(new Run(Application.Current.Resources["GNULicenceLink"].ToString()))
            {
                NavigateUri = new Uri("https://www.gnu.org/licenses/gpl-3.0.en.html")
            };

            hyperlink.RequestNavigate += this.Hyperlink_RequestNavigate;
            this.TblkShortLicenceNotice.Inlines.Clear();
            this.TblkShortLicenceNotice.Inlines.Add(copyRightAndWarranty);
            this.TblkShortLicenceNotice.Inlines.Add(run2);
            this.TblkShortLicenceNotice.Inlines.Add(hyperlink);
            this.TblkShortLicenceNotice.Inlines.Add(run4);
            this.BtnUpdateSearcher.Visibility = Common.ApplicationUpdateExists ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>
        /// Setup the about window without the preferences file.
        /// </summary>
        private void SetupAppWithoutPreferences()
        {
            this.ChkUpdatesCheck.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Set readable content based on selected language.
        /// </summary>
        private void SetContentBasedOnLanguage()
        {
            this.Title = Application.Current.Resources["AboutSearcher"].ToString();
            this.ChkUpdatesCheck.Content = Application.Current.Resources["MonthlyUpdateCheck"].ToString();
            this.BtnUpdateSearcher.Content = Application.Current.Resources["Update"].ToString();
        }

        #endregion Private Methods        
    }
}
