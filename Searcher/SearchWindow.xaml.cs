// <copyright file="SearchWindow.xaml.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shell;
using System.Windows.Threading;
using System.Xml.Linq;
using SearcherLibrary;
using Application = System.Windows.Application;
using Clipboard = System.Windows.Clipboard;
using ComboBox = System.Windows.Controls.ComboBox;
using ContextMenu = System.Windows.Controls.ContextMenu;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using MessageBoxOptions = System.Windows.MessageBoxOptions;
using TextBox = System.Windows.Controls.TextBox;
using Timer = System.Timers.Timer;

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
    /// <summary>
    /// Interaction logic for MainWindow.
    /// </summary>
    public partial class SearchWindow : Window, INotifyPropertyChanged
    {
        #region Private Fields

        /// <summary>
        /// The maximum scale value allowed.
        /// </summary>
        private const double MaxScaleValue = 3.0;

        /// <summary>
        /// Private store for the maximum length in the text for the search directories.
        /// </summary>
        private const int MaxSearchDirectoryTextLength = 1000;

        /// <summary>
        /// Private store for the maximum length in the filters used for searching.
        /// </summary>
        private const int MaxSearchFilterTextLength = 50;

        /// <summary>
        /// Private store for the maximum length in the search text.
        /// </summary>
        private const int MaxSearchTextLength = 200;

        /// <summary>
        /// Private store for limiting display of long strings.
        /// </summary>
        private const int MaxStringLengthCheck = 2000;

        /// <summary>
        /// Private store for setting the end index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const int MaxStringLengthDisplayIndexEnd = 200;

        /// <summary>
        /// Private store for setting the start index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const int MaxStringLengthDisplayIndexStart = 100;

        /// <summary>
        /// The minimum scale value allowed.
        /// </summary>
        private const double MinScaleValue = 0.5;

        /// <summary>
        /// Lock object to determine matches from which file will update the UI.
        /// </summary>
        private static object syncRoot = new object();

        /// <summary>
        /// Private store for the background colour of the application
        /// </summary>
        private SolidColorBrush applicationBackColour = new SolidColorBrush();

        /// <summary>
        /// Cancellation token source object to cancel file search.
        /// </summary>
        private CancellationTokenSource cancellationTokenSource;

        /// <summary>
        /// List of child result pop-out windows.
        /// </summary>
        private List<ResultsPopout> childWindows = new List<ResultsPopout>();

        /// <summary>
        /// Private store for the object that pops up the content.
        /// </summary>
        private ContentPopup contentPopup;

        /// <summary>
        /// Language used on the form.
        /// </summary>
        private string culture = "en-US";

        /// <summary>
        /// Private store for the list of directories that will not be searched (temporarily or always).
        /// </summary>
        private List<string> directoriesToExclude = new List<string>();

        /// <summary>
        /// Boolean to store whether the list of files being searched have been finalised.
        /// </summary>
        private bool distinctFilesFound = false;

        /// <summary>
        /// Variable to store the path for a custom editor.
        /// </summary>
        private string editorNamePath;

        /// <summary>
        /// Variable to store stopwatch that calculates execution time.
        /// </summary>
        private Stopwatch executionTime = new Stopwatch();

        /// <summary>
        /// Private store for the file name that will be used to update the UI when lots of matches exist.
        /// </summary>
        private string fileBeingDisplayed = string.Empty;

        /// <summary>
        /// Private store for the list of files that have already been searched.
        /// </summary>
        private List<string> filesSearched;

        /// <summary>
        /// Track count of number of files already searched.
        /// </summary>
        private int filesSearchedCounter = 0;

        /// <summary>
        /// Private store for the determining the progress of the files currently being searched.
        /// </summary>
        private ConcurrentDictionary<string, FileSearchResult> filesSearchProgress = new ConcurrentDictionary<string, FileSearchResult>();      // TODO: Use ConcurrentBag, but only .NET 5.0 has the Clear() method.

        /// <summary>
        /// Private store for the list of files that will not be searched (temporarily or always).
        /// </summary>
        private List<string> filesToExclude = new List<string>();

        /// <summary>
        /// Private store for the list of files that are being searched.
        /// </summary>
        private List<string> filesToSearch;

        /// <summary>
        /// Variable to store the number of files that have a search match.
        /// </summary>
        private int filesWithMatch = 0;

        /// <summary>
        /// Private store to determine whether the filters are to be excluded.
        /// </summary>
        private bool filterExclusionSet = false;

        /// <summary>
        /// The back colour value for matching result values.
        /// </summary>
        private SolidColorBrush highlightResultBackColour = Brushes.PeachPuff;

        /// <summary>
        /// Boolean indicating whether results are to be highlighted.
        /// </summary>
        private bool highlightResults = false;

        /// <summary>
        /// Boolean indicating whether the search is case sensitive.
        /// </summary>
        private bool matchCase = false;

        /// <summary>
        /// Private store for the matcher object used to search files.
        /// </summary>
        private Matcher matcherObj = new Matcher();

        /// <summary>
        /// Variable to store the number of search matches.
        /// </summary>
        private int matchesFound = 0;

        /// <summary>
        /// Boolean indicating whether the search is for the whole word.
        /// </summary>
        private bool matchWholeWord = false;

        /// <summary>
        /// Variable to store the maximum number of search history.
        /// </summary>
        private int maxDropDownItems = 10;

        /// <summary>
        /// The minimum search date to use for file creation.
        /// </summary>
        private DateTime maxSearchDate = DateTime.MaxValue;

        /// <summary>
        /// The minimum search date to use for file creation.
        /// </summary>
        private DateTime minSearchDate = DateTime.MinValue;

        /// <summary>
        /// Boolean indicating whether the search is across multiple lines using Regex.
        /// </summary>
        private bool multilineRegex = false;

        /// <summary>
        /// Private store for the default height of the content popup window.
        /// </summary>
        private int popupWindowHeight = 300;

        /// <summary>
        /// Private store for the timeout of the popup window.
        /// </summary>
        private int popupWindowTimeoutSeconds = 4;

        /// <summary>
        /// Private store for the default width of the content popup window.
        /// </summary>
        private int popupWindowWidth = 500;

        /// <summary>
        /// The regex options to use when searching.
        /// </summary>
        private RegexOptions regexOptions = RegexOptions.None;

        /// <summary>
        /// The paragraph object used as a search result.
        /// </summary>
        private Paragraph richTextboxParagraph;

        /// <summary>
        /// Private store for the scale for the search results.
        /// </summary>
        private double scale = 1.0;

        /// <summary>
        /// Boolean indicating whether the search mode is normal.
        /// </summary>
        private bool searchModeNormal = false;

        /// <summary>
        /// Boolean indicating whether the search mode uses regex.
        /// </summary>
        private bool searchModeRegex = false;

        /// <summary>
        /// Private store for the time the search was started.
        /// </summary>
        private DateTime searchStartTime = DateTime.Now;

        /// <summary>
        /// Boolean indicating whether the search should look in sub folders.
        /// </summary>
        private bool searchSubFolders = false;

        /// <summary>
        /// Private store for timer to determine time taken to run search.
        /// </summary>
        private Timer searchTimer = new Timer(1000);

        /// <summary>
        /// Boolean indicating whether the search must contain all of the search terms.
        /// </summary>
        private bool searchTypeAll = false;

        /// <summary>
        /// The separator character to identify individual terms;
        /// </summary>
        private string separatorCharacter = ";";

        /// <summary>
        /// Gets or sets a value indicating whether preferences should be saved.
        /// </summary>
        private bool shouldSavePreferences = true;

        /// <summary>
        /// Boolean indicating whether the search should show the execution time.
        /// </summary>
        private bool showExecutionTime = false;

        /// <summary>
        /// Boolean indicating whether match counts are displayed for each file.
        /// </summary>
        private bool showMatchCount = true;

        /// <summary>
        /// Private store for the window height.
        /// </summary>
        private double windowHeight = 0;

        /// <summary>
        /// Private store for the window width.
        /// </summary>
        private double windowWidth = 0;

        /// <summary>
        /// Private store to hide the zoom level text block.
        /// </summary>
        private DispatcherTimer zoomLabelTimer = new DispatcherTimer();

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchWindow"/> class.
        /// </summary>
        public SearchWindow()
        {
            this.InitializeComponent();
            this.InitialiseControls();
            this.DataContext = this;
        }

        #endregion Public Constructors

        #region Public Events

        /// <summary>
        /// Event handler for the Property changed event.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Public Events

        #region Public Properties

        /// <summary>
        /// Gets or sets the scale for the results window.
        /// </summary>
        public double Scale
        {
            get
            {
                return this.scale;
            }

            set
            {
                this.scale = value;
                this.NotifyPropertyChanged("Scale");
            }
        }

        #endregion Public Properties

        #region Protected Methods

        /// <summary>
        /// The Notify Property Changed event handler method.
        /// </summary>
        /// <param name="strPropertyName">The name of the property</param>
        protected void NotifyPropertyChanged(string strPropertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(strPropertyName));
            }
        }

        #endregion Protected Methods

        #region Private Methods

        /// <summary>
        /// Saves current search preferences for future use.
        /// </summary>
        /// <param name="comboBox">The combo box object.</param>
        /// <param name="preferenceElement">The search preference value.</param>
        private void AddItemsToPreferences(ComboBox comboBox, string preferenceElement)
        {
            IEnumerable<string> itemsToAdd = comboBox.Items.Cast<string>();
            PreferencesHandler.AddItemsToPreferences(itemsToAdd, preferenceElement);
        }

        /// <summary>
        /// Adds the historic search preferences if they have been saved.
        /// </summary>
        /// <param name="comboBox">The combo box object.</param>
        /// <param name="preferenceElement">The search preference value.</param>
        private void AddPreferencesToItems(ComboBox comboBox, string preferenceElement)
        {
            PreferencesHandler.GetPreference(preferenceElement).Descendants("Value").ToList().ForEach(p =>
            {
                comboBox.Items.Add(p.Value);
            });
        }

        /// <summary>
        /// Add the result to the display.
        /// </summary>
        /// <param name="matchedLines">The list of match objects to display matches</param>
        /// <returns>List of Inline to be added to the results.</returns>
        private async Task<List<Inline>> AddResult(List<MatchedLine> matchedLines)
        {
            List<Inline> retVal = new List<Inline>();
            string currentFileName = string.Empty;
            string matchCountStr = $" ({Application.Current.Resources["Matches"]}: {matchedLines.Count})";
            int matchCounter = 0;

            foreach (MatchedLine ml in matchedLines)
            {
                matchCounter++;

                if (!ml.DisplayProcessed && !this.cancellationTokenSource.Token.IsCancellationRequested)
                {
                    if (!string.IsNullOrEmpty(ml.FileName) && !this.filesSearched.Contains(ml.FileName))
                    {
                        this.filesWithMatch++;
                        this.filesSearched.Add(ml.FileName);
                        currentFileName = ml.FileName;

                        this.Dispatcher.Invoke(() =>
                        {
                            Hyperlink link = new Hyperlink(new Run(Environment.NewLine + ml.FileName));
                            link.NavigateUri = new Uri(ml.FileName);
                            link.RequestNavigate += Link_RequestNavigate;
                            link.PreviewMouseRightButtonDown += Link_PreviewMouseRightButtonDown;
                            retVal.Add(new Bold(link)
                            {
                                Foreground = Brushes.CornflowerBlue
                            });

                            if (this.showMatchCount)
                            {
                                retVal.Add(new Run(matchCountStr));
                            }

                            retVal.Add(new Run(Environment.NewLine));
                        });
                    }

                    this.matchesFound++;
                    this.Dispatcher.Invoke(() =>
                    {
                        retVal.Add(new Run(ml.Content.Substring(0, ml.StartIndex)));
                        retVal.Add(new Run(Regex.Unescape(Regex.Escape(ml.Content.Substring(ml.StartIndex, ml.Length))))
                        {
                            Background = this.highlightResults ? this.highlightResultBackColour : this.applicationBackColour
                        });
                        retVal.Add(new Run(ml.Content.Substring(ml.StartIndex + ml.Length, ml.Content.Length - (ml.StartIndex + ml.Length))));

                        retVal.Add(new Run(Environment.NewLine));
                    });

                    ml.DisplayProcessed = true;
                }

                if (matchCounter++ % 100 == 0)
                {
                    await Task.Delay(5);
                }
            }

            return retVal;
        }

        /// <summary>
        /// Adds the result matches to the textbox.
        /// </summary>
        /// <param name="resultsToAdd">The list of Inline to display.</param>
        /// <returns>Task to add contents to textbox.</returns>
        private async Task AddResultsToTextbox(List<Inline> resultsToAdd)
        {
            if (!this.cancellationTokenSource.IsCancellationRequested)
            {
                int rowsToTake = 100;

                if (resultsToAdd.Count < rowsToTake)
                {
                    this.richTextboxParagraph.Inlines.AddRange(resultsToAdd);
                }
                else
                {
                    int completed = 0;
                    while (completed < resultsToAdd.Count)
                    {
                        this.richTextboxParagraph.Inlines.AddRange(resultsToAdd.Skip(completed).Take(rowsToTake));
                        completed += rowsToTake;
                        await Task.Delay(5);       // If there are lots of rows, give the UI time to refresh itself.
                    }
                }
            }
        }

        /// <summary>
        /// Adds the search criteria to the combo box.
        /// </summary>
        /// <param name="comboBox">The combo box object.</param>
        /// <param name="criteria">The search criteria.</param>
        private void AddSearchCriteria(ComboBox comboBox, string criteria)
        {
            if (!comboBox.Items.Contains(criteria))
            {
                if (comboBox.Items.Count == this.maxDropDownItems)
                {
                    comboBox.Items.RemoveAt(this.maxDropDownItems - 1);
                }

                if (comboBox.Items.Count > 0)
                {
                    comboBox.Items.Add(string.Empty);
                    for (int counter = comboBox.Items.Count - 1; counter > 0; counter--)
                    {
                        comboBox.Items[counter] = comboBox.Items[counter - 1];
                    }

                    comboBox.Items[0] = criteria;
                    comboBox.SelectedIndex = 0;
                }
                else
                {
                    comboBox.Items.Add(criteria);
                }
            }
        }

        /// <summary>
        /// Add the file to the list of files that will not be searched.
        /// </summary>
        /// <param name="fullFileName">The full path and name of the file.</param>
        /// <returns>Boolean indicating whether the file was added.</returns>
        private bool AddToSearchFileExclusionList(string fullFileName)
        {
            bool retVal = false;

            if (string.IsNullOrEmpty(fullFileName))
            {
                this.ShowErrorPopup(Application.Current.Resources["ArchiveFileExclusionNotAllowed"].ToString());
            }
            else
            {
                if (!this.filesToExclude.Any(f => f.ToUpper() == fullFileName.ToUpper()) && File.Exists(fullFileName))
                {
                    this.filesToExclude.Add(fullFileName);
                    retVal = true;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Open the about window as a dialog.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void BtnAbout_Click(object sender, RoutedEventArgs e)
        {
            SearcherAboutWindow saw = new SearcherAboutWindow();

            if (PreferencesHandler.PreferencesFile != null)
            {
                saw.AppHasPreferencesFile = PreferencesHandler.PreferencesFile != null;
                saw.CanCheckForUpdates = PreferencesHandler.GetPreferenceValue("CheckForUpdates") == true.ToString();
            }

            saw.Owner = this;
            saw.ShowDialog();

            if (PreferencesHandler.PreferencesFile != null)
            {
                PreferencesHandler.SetPreferenceValue("CheckForUpdates", saw.CanCheckForUpdates.ToString());
            }

            if (saw.ClosingForUpdate)
            {
                Process.Start(Application.ResourceAssembly.Location);
                this.Close();
            }
        }

        /// <summary>
        /// Event handler for the cancel search button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (!this.cancellationTokenSource.IsCancellationRequested)
            {
                this.cancellationTokenSource.Cancel();
                this.searchTimer.Stop();
                this.EnableSearchControls(true);
                this.BtnSearch.IsEnabled = true;
                this.BtnCancel.IsEnabled = false;
                this.SetFileCounterProgressInformation(-1, Application.Current.Resources["SearchCancelled"].ToString());
            }
        }

        /// <summary>
        /// Event handler for the changing the editor.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnChangeEditor_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = Application.Current.Resources["ExecutableFiles"].ToString() + " (*.exe)|*.exe";
            ofd.Multiselect = false;
            ofd.InitialDirectory = @"C:\Windows\System32";
            ofd.RestoreDirectory = true;
            ofd.Title = Application.Current.Resources["SelectEditor"].ToString();

            if (PreferencesHandler.GetPreference("CustomEditor") != null && PreferencesHandler.GetPreference("CustomEditor").Count() == 1
                && PreferencesHandler.GetPreference("CustomEditor").FirstOrDefault() != null && !string.IsNullOrWhiteSpace(PreferencesHandler.GetPreferenceValue("CustomEditor")))
            {
                try
                {
                    DirectoryInfo parentDir = Directory.GetParent(PreferencesHandler.GetPreferenceValue("CustomEditor"));
                    if (parentDir.Exists)
                    {
                        ofd.InitialDirectory = parentDir.FullName;
                    }
                }
                catch (Exception)
                {
                    // Do not throw. Let the initial directory stay at the default.
                }
            }

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.editorNamePath = ofd.FileName;
                this.TxtEditor.Text = Path.GetFileNameWithoutExtension(ofd.FileName);
                PreferencesHandler.SetPreferenceValue("CustomEditor", this.editorNamePath);
            }
        }

        /// <summary>
        /// Event handler for the set directory button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnDirectory_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (!string.IsNullOrEmpty(this.CmbDirectory.Text))
            {
                fbd.SelectedPath = this.CmbDirectory.Text.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Last().Trim();
            }

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Break and recreate the paths cleanly, based on separating and joining using the ";" character.
                if (string.IsNullOrWhiteSpace(this.CmbDirectory.Text))
                {
                    this.CmbDirectory.Text = fbd.SelectedPath;
                }
                else
                {
                    List<string> directories = this.CmbDirectory.Text.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    directories.Add(fbd.SelectedPath);
                    this.CmbDirectory.Text = string.Join("; ", directories.Select(dir => dir.Trim()).Distinct());
                }

                this.CmbDirectory.Text = this.CmbDirectory.Text.Trim();
            }
        }

        /// <summary>
        /// Event handler for the search button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (this.CanSearch())
            {
                this.cancellationTokenSource = new CancellationTokenSource();
                this.SetSearchParameters();
                this.ResetSearch();
                this.EnableSearchControls(false);
                string searchText = this.CmbFindWhat.Text;
                string searchPath = this.CmbDirectory.Text;
                string filters = string.IsNullOrEmpty(this.CmbFilters.Text) ? "*.*" : this.CmbFilters.Text;
                this.TxtResults.IsDocumentEnabled = true;
                this.TaskBarItemInfoProgress.ProgressValue = 0;
                this.TaskBarItemInfoProgress.ProgressState = TaskbarItemProgressState.Normal;
                this.executionTime.Start();
                this.searchTimer.Start();

                try
                {
                    await this.PerformSearchAsync(searchText, searchPath, filters);
                }
                catch (Exception ex)
                {
                    if (ex is AggregateException)
                    {
                        foreach (Exception innerExc in ((AggregateException)ex).InnerExceptions)
                        {
                            if (innerExc != null && innerExc.InnerException != null)
                            {
                                this.SetSearchError(innerExc.InnerException.Message);
                            }
                        }
                    }
                    else
                    {
                        this.SetSearchError(ex.Message);
                    }
                }
                finally
                {
                    this.executionTime.Stop();
                    this.searchTimer.Stop();
                    this.TaskBarItemInfoProgress.ProgressState = TaskbarItemProgressState.None;
                    this.BtnCancel.IsEnabled = false;
                    this.BtnSearch.IsEnabled = true;
                    this.EnableSearchControls(true);
                    this.CmbFindWhat.Focus();
                    string elapsedTime = this.executionTime.Elapsed.ToString();
                    this.executionTime.Reset();

                    if (this.showExecutionTime)
                    {
                        MessageBox.Show(string.Format("{0}: {1}", Application.Current.Resources["TimeToSearch"].ToString(), elapsedTime), Application.Current.Resources["SearchComplete"].ToString(), MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
                    }
                }
            }
        }

        /// <summary>
        /// Determine whether the basic criteria is provided to perform a search.
        /// </summary>
        /// <returns>Boolean indicating whether the search can be performed.</returns>
        private bool CanSearch()
        {
            bool retVal = true;

            if (string.IsNullOrEmpty(this.CmbFilters.Text))
            {
                this.CmbFilters.Text = "*.*";
            }

            if (string.IsNullOrEmpty(this.CmbDirectory.Text) || string.IsNullOrEmpty(this.CmbFindWhat.Text))
            {
                MessageBox.Show(Application.Current.Resources["SpecifySearchDirectory"].ToString());
                retVal = false;
            }
            else if (this.DtpEndDate.SelectedDate != null && this.DtpEndDate.SelectedDate.HasValue && this.DtpStartDate.SelectedDate != null && this.DtpStartDate.SelectedDate.HasValue)
            {
                if (this.DtpStartDate.SelectedDate.Value > this.DtpEndDate.SelectedDate.Value)
                {
                    MessageBox.Show(Application.Current.Resources["StartDateLessThanEndDateError"].ToString());
                    retVal = false;
                }
            }
            else
            {
                List<string> errorPaths = new List<string>();
                this.CmbDirectory.Text.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList().ForEach(d =>
                {
                    if (!Directory.Exists(d))
                    {
                        errorPaths.Add(d);
                    }
                });

                if (errorPaths.Count > 0)
                {
                    // Display top 5 failed paths. If more, then first fix top 5 atleast.
                    string errorMessage = string.Format(
                        "{0}:{1}{2}{3}{4}",
                        Application.Current.Resources["DirectoryPathNotFound"].ToString(),
                        Environment.NewLine,
                        string.Join(Environment.NewLine, errorPaths.Take(5).ToArray()),
                        Environment.NewLine,
                        Application.Current.Resources["UseConfigDelimiterCharacter"].ToString());

                    MessageBox.Show(errorMessage);
                    retVal = false;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Set the language on the form.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void CmbLanguage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.culture = ((ComboBoxItem)CmbLanguage.SelectedItem).Tag.ToString();
            Uri imageFile = new Uri("/Images/" + this.culture + ".png", UriKind.Relative);
            this.ImgFlag.Source = new BitmapImage(imageFile);
            this.SetLanguage();
        }

        /// <summary>
        /// Handle the close event of the content popup to retain the window size.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void ContentPopup_Closed(object sender, EventArgs e)
        {
            this.popupWindowHeight = (int)this.contentPopup.Height;
            this.popupWindowWidth = (int)this.contentPopup.Width;
            this.contentPopup.Closed -= this.ContentPopup_Closed;
        }

        /// <summary>
        /// Context menu to copy the file name without extension to the clipboard.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void CopyFileNameNoExtToClipboard_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);
            Clipboard.SetText(Path.GetFileNameWithoutExtension(fullFilePath));
        }

        /// <summary>
        /// Context menu to copy the file name to the clipboard.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void CopyFileNameToClipboard_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);
            Clipboard.SetText(Path.GetFileName(fullFilePath));
        }

        /// <summary>
        /// Context menu to copy the full file path to the clipboard.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void CopyFullPathToClipboard_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);
            Clipboard.SetText(Path.GetFullPath(fullFilePath));
        }

        /// <summary>
        /// Determine whether updates can be checked for.
        /// </summary>
        private async void DownloadUpdates()
        {
            // This method is async void because if the search for update fails, we do not worry further as it does not impact the application usage.
            if (PreferencesHandler.PreferencesFile != null && PreferencesHandler.GetPreferenceValue("CheckForUpdates").ToUpper() == true.ToString().ToUpper())
            {
                UpdateApp updateApp = new UpdateApp(PreferencesHandler.PreferencesFile);

                if (updateApp.UpdatedAppExistsLocally())
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        this.Title = string.Format("{0} ({1})", this.Title, Application.Current.Resources["ClickAboutForUpdates"].ToString());
                    });
                }
                else
                {
                    string lastUpdateCheckDate = await updateApp.CheckAndDownloadUpdatedVersionAsync();
                    PreferencesHandler.SetPreferenceValue("LastUpdateCheckDate", lastUpdateCheckDate);
                }
            }
        }

        /// <summary>
        /// Perform search if enter key is pressed.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void DtpEndDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                this.BtnSearch_Click(sender, e);
            }
        }

        /// <summary>
        /// Set start/end date based on selected end date.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The SelectionChangedEventArgs object.</param>
        private void DtpEndDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DtpStartDate.SelectedDate.GetValueOrDefault() > DtpEndDate.SelectedDate.GetValueOrDefault())
            {
                DtpStartDate.SelectedDate = DtpEndDate.SelectedDate;
            }
        }

        /// <summary>
        /// Perform search if enter key is pressed.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void DtpStartDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                this.BtnSearch_Click(sender, e);
            }
        }

        /// <summary>
        /// Set start/end date based on selected start date.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The SelectionChangedEventArgs object.</param>
        private void DtpStartDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DtpEndDate.SelectedDate.GetValueOrDefault() < DtpStartDate.SelectedDate.GetValueOrDefault())
            {
                DtpEndDate.SelectedDate = DtpStartDate.SelectedDate;
            }
        }

        /// <summary>
        /// Enable/disable the search controls.
        /// </summary>
        /// <param name="enableControls">Boolean indicating the availability of search controls.</param>
        private void EnableSearchControls(bool enableControls)
        {
            this.CmbDirectory.IsEnabled = enableControls;
            this.BtnDirectory.IsEnabled = enableControls;
            this.CmbFindWhat.IsEnabled = enableControls;
            this.CmbFilters.IsEnabled = enableControls;
            this.CmbSearchType.IsEnabled = enableControls;

            this.GrpMatchOptions.IsEnabled = enableControls;
            this.GrpSearchMode.IsEnabled = enableControls;
            this.GrpMisc.IsEnabled = enableControls;
        }

        /// <summary>
        /// Sets up a directory to be excluded.
        /// </summary>
        /// <param name="sender">The sender object to get the file path.</param>
        /// <param name="e">The parameter is not used.</param>
        private void ExcludeDirectory_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);

            if (!string.IsNullOrEmpty(fullFilePath))
            {
                DirectoryExclude dirExcludeWindow = new DirectoryExclude(fullFilePath);
                dirExcludeWindow.PreferenceFileExists = PreferencesHandler.PreferencesFile != null;
                dirExcludeWindow.Owner = this;

                if (dirExcludeWindow.ShowDialog() == true)
                {
                    if (!this.directoriesToExclude.Contains(dirExcludeWindow.DirectoryToExclude) && Directory.Exists(dirExcludeWindow.DirectoryToExclude))
                    {
                        this.directoriesToExclude.Add(dirExcludeWindow.DirectoryToExclude);

                        if (dirExcludeWindow.IsExclusionPermanent)
                        {
                            PreferencesHandler.GetPreference("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", dirExcludeWindow.DirectoryToExclude));
                        }
                    }
                }
            }
            else
            {
                this.ShowErrorPopup(Application.Current.Resources["ExcludeArchiveAsFile"].ToString());
            }
        }

        /// <summary>
        /// Exclude file from being searched via config.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void ExcludeFromSearchAlways_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);
            if (this.AddToSearchFileExclusionList(fullFilePath))
            {
                PreferencesHandler.GetPreference("FilesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fullFilePath));
            }
        }

        /// <summary>
        /// Exclude file from being searched until the application is closed.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void ExcludeFromSearchTemporarily_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);
            this.AddToSearchFileExclusionList(fullFilePath);
        }

        /// <summary>
        /// Get the file path for which the mouse is double clicked.
        /// </summary>
        /// <returns>The file path for which the mouse is double clicked.</returns>
        private string GetFileForPopup()
        {
            string retVal = string.Empty;

            if (this.TxtResults.CaretPosition != null && this.TxtResults.CaretPosition.Parent != null && this.TxtResults.CaretPosition.Parent is Run)
            {
                Inline fileNameInline = ((Run)(Inline)this.TxtResults.CaretPosition.Parent).PreviousInline;

                while (!(fileNameInline is Bold) && fileNameInline != null)
                {
                    fileNameInline = (Inline)fileNameInline.PreviousInline;
                }

                if (fileNameInline != null)
                {
                    retVal = new TextRange(fileNameInline.ContentStart.GetLineStartPosition(1), fileNameInline.ContentStart.GetLineStartPosition(2)).Text;
                    retVal = retVal.Replace(Environment.NewLine, string.Empty);
                }

                int indexOfMatchCount = retVal.LastIndexOf($" ({Application.Current.Resources["Matches"]}: ");

                if (indexOfMatchCount > 0)
                {
                    retVal = retVal.Substring(0, indexOfMatchCount);
                }
            }

            return retVal;
        }

        /// <summary>
        /// Get list of files to search.
        /// </summary>
        /// <param name="path">The path to use for searching files.</param>
        /// <param name="filter">The filter to use when capturing files.</param>
        /// <returns>List of files to be searched.</returns>
        private List<string> GetFilesToSearch(string path, string filter)
        {
            List<string> filesToSearch = new List<string>();
            List<string> pathErrors = new List<string>();

            if (this.filterExclusionSet)
            {
                filter = "*.*";
            }

            if (!this.cancellationTokenSource.Token.IsCancellationRequested)
            {
                // Recurse directories to ensure permissions exist for accessing sub directories.
                if (this.searchSubFolders)
                {
                    try
                    {
                        foreach (string dirPath in Directory.EnumerateDirectories(path))
                        {
                            filesToSearch.AddRange(this.GetFilesToSearch(dirPath, filter));
                        }
                    }
                    catch (Exception ex)
                    {
                        pathErrors.Add(ex.Message);
                    }
                }

                try
                {
                    filesToSearch.AddRange(Directory.EnumerateFiles(path, filter));

                    if (this.minSearchDate > DateTime.MinValue)
                    {
                        filesToSearch.RemoveAll(f => File.GetCreationTime(f) < this.minSearchDate);
                    }

                    if (this.maxSearchDate < DateTime.MaxValue)
                    {
                        filesToSearch.RemoveAll(f => File.GetCreationTime(f) > this.maxSearchDate);
                    }
                }
                catch (Exception ex)
                {
                    // Add errors for path only if not set before.
                    if (!pathErrors.Contains(ex.Message))
                    {
                        pathErrors.Add(ex.Message);
                    }
                }

                if (pathErrors.Count > 0)
                {
                    // Show errors for failing paths.
                    pathErrors.ForEach(err =>
                    {
                        this.SetSearchError(err);
                    });
                }
            }

            return filesToSearch;
        }

        /// <summary>
        /// Get the line number for which the mouse is double clicked.
        /// </summary>
        /// <returns>The line number for which the mouse is double clicked.</returns>
        private int GetLineNumberForPopup()
        {
            int retVal = 0;

            if (this.TxtResults.CaretPosition.Parent != null && this.TxtResults.CaretPosition.Parent is Run && ((Run)this.TxtResults.CaretPosition.Parent).PreviousInline != null)
            {
                int startPosition = this.TxtResults.CaretPosition.IsAtLineStartPosition ? -1 : 0;
                int nextPosition = startPosition + 1;
                string text = string.Empty;
                bool lineFound = false;

                try
                {
                    while (!text.StartsWith("Line"))
                    {
                        text = new TextRange(this.TxtResults.CaretPosition.GetLineStartPosition(startPosition), this.TxtResults.CaretPosition.GetLineStartPosition(nextPosition)).Text + text;
                        startPosition--;
                        nextPosition--;
                        lineFound = text.StartsWith("Line");
                    }
                }
                catch (ArgumentNullException)
                {
                    lineFound = false;
                }

                if (lineFound)
                {
                    try
                    {
                        int.TryParse(text.Substring(5, text.IndexOf(":") - 5), out retVal);
                    }
                    catch (ArgumentOutOfRangeException aore)
                    {
                        if (this.multilineRegex)
                        {
                            // Unable to get caret position. Do not fail and close application.
                        }
                        else
                        {
                            this.ShowError(aore);
                        }
                    }
                }
            }

            return retVal;
        }

        /// <summary>
        /// Get the file that is taking the longest time to run.
        /// </summary>
        /// <param name="fileName">The default file name to be displayed.</param>
        /// <returns>String containing file that is taking the longest time to run.</returns>
        private string GetLongestRunningFile(string fileName)
        {
            string retVal = fileName;
            retVal = this.filesSearchProgress.Where(fsr => fsr.Value.SearchStartDateTimeTicks > 0).OrderBy(fsr => fsr.Value.SearchStartDateTimeTicks).FirstOrDefault().Value?.FileNamePath ?? fileName;
            return retVal;
        }

        /// <summary>
        /// Sets the title bar to display details of user running the application.
        /// </summary>
        /// <returns>String containing user name information.</returns>
        private string GetRunningUserInfo()
        {
            AssemblyTitleAttribute assemblyTitleAttribute = (AssemblyTitleAttribute)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyTitleAttribute), false);
            string programName = assemblyTitleAttribute != null ? assemblyTitleAttribute.Title : Application.Current.Resources["UnknownAssemblyName"].ToString();
            string userName = string.Format(@"{0}\{1}", Environment.UserDomainName, Environment.UserName);
            WindowsPrincipal principal = new WindowsPrincipal(WindowsIdentity.GetCurrent());

            if (principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                userName += " - " + Application.Current.Resources["Administrator"].ToString();
            }

            return string.Format("{0} ({1})", programName, userName);
        }

        /// <summary>
        /// Show the files used in the search
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseButtonEventArgs object.</param>
        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if ((this.distinctFilesFound && this.filesToSearch != null && this.filesToSearch.Count > 0) || (this.filesToExclude.Count > 0 || this.directoriesToExclude.Count > 0))
            {
                SearchedFileList searchedFileList;
                List<string> filesToAlwaysExclude = new List<string>();
                List<string> directoriesToAlwaysExclude = new List<string>();

                if (PreferencesHandler.PreferencesFile != null)
                {
                    filesToAlwaysExclude = PreferencesHandler.GetPreference("FilesToAlwaysExcludeFromSearch").Descendants("Value").Select(p => p.Value).ToList();
                    directoriesToAlwaysExclude = PreferencesHandler.GetPreference("DirectoriesToAlwaysExcludeFromSearch").Descendants("Value").Select(p => p.Value).ToList();

                    if (this.filesToExclude.Count > 0 || this.directoriesToExclude.Count > 0)
                    {
                        if (filesToAlwaysExclude.Count > 0 || directoriesToAlwaysExclude.Count > 0)
                        {
                            List<string> tempFilesToExclude = new List<string>();
                            tempFilesToExclude.AddRange(this.directoriesToExclude);
                            tempFilesToExclude.AddRange(this.filesToExclude);
                            tempFilesToExclude.RemoveAll(tfe => filesToAlwaysExclude.Any(fae => fae == tfe));
                            tempFilesToExclude.RemoveAll(tfe => directoriesToAlwaysExclude.Any(dae => dae == tfe));
                            searchedFileList = new SearchedFileList(this.filesToSearch, tempFilesToExclude, filesToAlwaysExclude.Concat(directoriesToAlwaysExclude));
                        }
                        else
                        {
                            searchedFileList = new SearchedFileList(this.filesToSearch, this.filesToExclude.Concat(this.directoriesToExclude));
                        }
                    }
                    else
                    {
                        searchedFileList = new SearchedFileList(this.filesToSearch);
                    }
                }
                else
                {
                    searchedFileList = new SearchedFileList(this.filesToSearch);
                }

                Point windowLocation = this.PgBarSearch.PointToScreen(new Point(0, 0));
                searchedFileList.Left = windowLocation.X + (this.PgBarSearch.ActualWidth / 4);
                searchedFileList.Top = windowLocation.Y;
                searchedFileList.Owner = this;
                searchedFileList.ShowDialog();

                if (PreferencesHandler.PreferencesFile != null)
                {
                    if (searchedFileList.FilesToInclude.Count > 0)
                    {
                        this.filesToExclude.RemoveAll(tfe => searchedFileList.FilesToInclude.Any(fi => File.Exists(fi) && fi == tfe));
                        this.directoriesToExclude.RemoveAll(tde => searchedFileList.FilesToInclude.Any(fi => Directory.Exists(fi) && fi == tde));
                        filesToAlwaysExclude.RemoveAll(fae => searchedFileList.FilesToInclude.Any(fi => File.Exists(fi) && fi == fae));
                        directoriesToAlwaysExclude.RemoveAll(fae => searchedFileList.FilesToInclude.Any(fi => Directory.Exists(fi) && fi == fae));
                        this.SavePathsToAlwaysExclude(filesToAlwaysExclude, directoriesToAlwaysExclude);
                    }
                }
            }
        }

        /// <summary>
        /// Hide the zoom level label.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The EventArgs object.</param>
        private void HideScaleTextBlock(object sender, EventArgs e)
        {
            this.TxtBlkScaleValue.Visibility = Visibility.Hidden;
            this.zoomLabelTimer.Stop();
        }

        /// <summary>
        /// Initialise controls on opening the search window for the first time.
        /// </summary>
        private void InitialiseControls()
        {
            try
            {
                this.searchTimer.Elapsed += this.SearchTimer_Elapsed;
                this.SetTitleInfo();

                if (!File.Exists(PreferencesHandler.PreferenceFilePath))
                {
                    string test = Application.Current.Resources["CreatePreferencesFileQuestion"].ToString();
                    if (MessageBox.Show(Application.Current.Resources["CreatePreferencesFileQuestion"].ToString(), Application.Current.Resources["NoSearchPreferences"].ToString(), MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.None) == MessageBoxResult.Yes)
                    {
                        PreferencesHandler.SetMainSearchWindow(this);
                        PreferencesHandler.SavePreferences();
                    }
                    else
                    {
                        shouldSavePreferences = false;
                        this.SetupAppWithoutPreferences();
                    }
                }
                else
                {
                    PreferencesHandler.SetMainSearchWindow(this);
                    PreferencesHandler.CheckPreferencesFile(PreferencesHandler.PreferenceFilePath);
                    this.SetInitialSearchOptions();

                    this.AddPreferencesToItems(this.CmbDirectory, "SearchDirectories");
                    this.AddPreferencesToItems(this.CmbFindWhat, "SearchContents");
                    this.AddPreferencesToItems(this.CmbFilters, "SearchFilters");
                    this.LoadPathsToAlwaysExlude();

                    this.CmbDirectory.Text = this.CmbDirectory.Items.Count > 0 ? this.CmbDirectory.Items[0].ToString() : string.Empty;
                    this.CmbFindWhat.Text = this.CmbFindWhat.Items.Count > 0 ? this.CmbFindWhat.Items[0].ToString() : string.Empty;
                    this.CmbFilters.Text = this.CmbFilters.Items.Count > 0 ? this.CmbFilters.Items[0].ToString() : string.Empty;
                }

                if (!string.IsNullOrEmpty(PreferencesHandler.PreferenceFilePath) && File.Exists(PreferencesHandler.PreferenceFilePath))
                {
                    this.SetDefaultCulture();
                    this.SetJumpList();
                    this.DownloadUpdates();
                }
            }
            catch (IOException ex)
            {
                this.ShowError(ex);
            }
        }

        /// <summary>
        /// Show right click menu options
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseButtonEventArgs object.</param>
        private void Link_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Hyperlink)
            {
                ContextMenu mnu = new ContextMenu();
                MenuItem copyFullPathToClipboard = new MenuItem();
                copyFullPathToClipboard.Header = Application.Current.Resources["CopyFullPath"].ToString();
                copyFullPathToClipboard.Tag = sender;
                copyFullPathToClipboard.Click += this.CopyFullPathToClipboard_Click;
                mnu.Items.Add(copyFullPathToClipboard);

                MenuItem copyFileNameToClipboard = new MenuItem();
                copyFileNameToClipboard.Header = Application.Current.Resources["CopyFileName"].ToString();
                copyFileNameToClipboard.Tag = sender;
                copyFileNameToClipboard.Click += this.CopyFileNameToClipboard_Click;
                mnu.Items.Add(copyFileNameToClipboard);

                MenuItem copyFileNameNoExtToClipboard = new MenuItem();
                copyFileNameNoExtToClipboard.Header = Application.Current.Resources["CopyFileNameNoExtension"].ToString();
                copyFileNameNoExtToClipboard.Tag = sender;
                copyFileNameNoExtToClipboard.Click += this.CopyFileNameNoExtToClipboard_Click;
                mnu.Items.Add(copyFileNameNoExtToClipboard);
                mnu.Items.Add(new Separator());

                MenuItem openContainingDirectory = new MenuItem();
                openContainingDirectory.Header = Application.Current.Resources["OpenDirectoryInExplorer"].ToString();
                openContainingDirectory.Tag = sender;
                openContainingDirectory.Click += this.OpenContainingDirectory_Click;
                mnu.Items.Add(openContainingDirectory);

                MenuItem saveAllResultsToFile = new MenuItem();
                saveAllResultsToFile.Header = Application.Current.Resources["SaveResultsToFile"].ToString();
                saveAllResultsToFile.Click += this.SaveAllResultsToFile_Click;
                mnu.Items.Add(saveAllResultsToFile);

                MenuItem popoutResults = new MenuItem();
                popoutResults.Header = Application.Current.Resources["PopOutResults"].ToString();
                popoutResults.Tag = sender;
                popoutResults.Click += this.PopoutResults_Click;
                mnu.Items.Add(popoutResults);

                mnu.Items.Add(new Separator());

                MenuItem excludeFileTemporarily = new MenuItem();
                excludeFileTemporarily.Header = Application.Current.Resources["ExcludeFileTemporarily"].ToString();
                excludeFileTemporarily.Tag = sender;
                excludeFileTemporarily.Click += this.ExcludeFromSearchTemporarily_Click;
                mnu.Items.Add(excludeFileTemporarily);

                if (PreferencesHandler.PreferencesFile != null)
                {
                    MenuItem excludeFileAlways = new MenuItem();
                    excludeFileAlways.Header = Application.Current.Resources["ExcludeFileAlways"].ToString();
                    excludeFileAlways.Tag = sender;
                    excludeFileAlways.Click += this.ExcludeFromSearchAlways_Click;
                    mnu.Items.Add(excludeFileAlways);
                }

                MenuItem excludeDirectory = new MenuItem();
                excludeDirectory.Header = Application.Current.Resources["ExcludeDirectory"].ToString();
                excludeDirectory.Tag = sender;
                excludeDirectory.Click += this.ExcludeDirectory_Click;
                mnu.Items.Add(excludeDirectory);

                ((Hyperlink)sender).ContextMenu = mnu;
            }
        }

        /// <summary>
        /// Event handler for the file name link.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RequestNavigateEventArgs object.</param>
        private void Link_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                if (Enum.GetNames(typeof(OtherExtensions)).Any(oe => Path.GetExtension(e.Uri.LocalPath).ToUpper().Contains(oe.ToUpper())))
                {
                    Process.Start(e.Uri.LocalPath);
                }
                else
                {
                    Process.Start(this.editorNamePath, "\"" + e.Uri.LocalPath + "\"");
                }
            }
            catch (Exception)
            {
                Process.Start("notepad", e.Uri.LocalPath);
                this.editorNamePath = "notepad";
            }
        }

        /// <summary>
        /// Set up the list of files that will be excluded always when searching.
        /// </summary>
        private void LoadPathsToAlwaysExlude()
        {
            PreferencesHandler.GetPreference("FilesToAlwaysExcludeFromSearch").Descendants("Value").ToList().ForEach(p =>
            {
                this.AddToSearchFileExclusionList(p.Value);
            });

            PreferencesHandler.GetPreference("DirectoriesToAlwaysExcludeFromSearch").Descendants("Value").ToList().ForEach(p =>
            {
                this.directoriesToExclude.Add(p.Value);
            });
        }

        /// <summary>
        /// Context menu to open windows explorer to the selected file.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void OpenContainingDirectory_Click(object sender, RoutedEventArgs e)
        {
            Common.OpenDirectoryForFile(sender);
        }

        /// <summary>
        /// Search for matches using parallel processing.
        /// </summary>
        /// <param name="searchText">The text containing terms to search.</param>
        /// <param name="searchPath">The text containing paths to search.</param>
        /// <param name="filters">The text containing filters to use for the search.</param>
        /// <returns>Task object that performs the search.</returns>
        private async Task PerformSearchAsync(string searchText, string searchPath, string filters)
        {
            List<string> filtersToUse = filters.Trim().Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList();
            List<string> searchPaths = searchPath.Trim().Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList();
            this.filesToSearch = new List<string>();
            this.filesSearched = new List<string>();
            this.distinctFilesFound = false;
            this.SetFileCounterProgressInformation(0, Application.Current.Resources["GettingFilesToSearch"].ToString());
            this.regexOptions = this.matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;

            if (this.searchModeRegex)
            {
                this.regexOptions = this.regexOptions | (this.multilineRegex ? RegexOptions.Multiline : RegexOptions.Singleline);
            }

            List<string> termsToSearch = this.SetTermsToSearchText(searchText);
            List<Task<List<string>>> pathsListTask = new List<Task<List<string>>>(searchPaths.Count());

            try
            {
                foreach (string path in searchPaths)
                {
                    foreach (string filter in filtersToUse)
                    {
                        pathsListTask.Add(Task.Run<List<string>>(() =>
                        {
                            return GetFilesToSearch(path, filter);
                        }));
                    }
                }

                await Task.WhenAll(pathsListTask);

                foreach (Task<List<string>> pathTask in pathsListTask)
                {
                    this.filesToSearch.AddRange(await pathTask);
                }

                if (this.filterExclusionSet)
                {
                    await RemoveFilesForFilterExclusion(filtersToUse);
                }

                await this.RemoveExclusionPaths();
                this.SetFileCounterProgressInformation(0, string.Format("{0}: {1}", Application.Current.Resources["FilesFound"].ToString(), this.filesToSearch.Count));
                this.filesToSearch = this.filesToSearch.Distinct().OrderBy(f => f).ToList();      // Remove duplicates that could be added via path filters that cover the same item mulitple times.
                this.distinctFilesFound = true;
                this.SetProgressMaxValue(this.filesToSearch.Count);
                await this.SearchParallelAsync(this.filesToSearch, termsToSearch);
            }
            catch (Exception ex)
            {
                if (!(ex is OperationCanceledException))
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Pop out the results of the current search to a new window.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void PopoutResults_Click(object sender, RoutedEventArgs e)
        {
            bool resultsCaptured = false;
            double horizontalOffset = this.TxtResults.HorizontalOffset;
            double verticalOffset = this.TxtResults.VerticalOffset;
            ResultsPopout newWnd = new ResultsPopout();
            newWnd.TxtResults.Background = this.applicationBackColour;
            newWnd.Width = this.Width;
            newWnd.Height = this.TxtResults.ActualHeight;
            newWnd.Title = string.Format("{0} - {1}", Application.Current.Resources["Results"].ToString(), this.CmbFindWhat.Text.Substring(0, Math.Min(this.CmbFindWhat.Text.Length, 50)));
            this.childWindows.Add(newWnd);
            newWnd.Closed += (windowSender, windowEventArgs) =>
            {
                this.childWindows.Remove(newWnd);
            };

            while (!resultsCaptured)
            {
                this.TxtResults.SelectAll();
                this.TxtResults.Copy();
                newWnd.TxtResults.Paste();
                this.TxtResults.ScrollToHorizontalOffset(horizontalOffset);
                this.TxtResults.ScrollToVerticalOffset(verticalOffset);
                this.TxtResults.CaretPosition = this.TxtResults.Document.ContentStart;  // Removes the "SelectAll" from the richtextbox
                this.CmbFindWhat.Focus();                                               // Set focus back on the search term.

                // If the copy did not succeed, only "\r\n" is set in the new window. So, retry the copy.
                TextRange textRange = new TextRange(newWnd.TxtResults.Document.ContentStart, newWnd.TxtResults.Document.ContentEnd);
                resultsCaptured = textRange.Text != "\r\n";
            }

            newWnd.Show();
        }

        /// <summary>
        /// If filters exclusion is set, remove the files that are to be exlcuded from search.
        /// </summary>
        /// <param name="filtersToUse">The list of filters to exclude.</param>
        private async Task RemoveFilesForFilterExclusion(List<string> filtersToUse)
        {
            await Task.Run(() =>
            {
                filesToSearch = filesToSearch.Distinct().ToList();

                foreach (string filter in filtersToUse)
                {
                    filesToSearch.RemoveAll(f => Regex.Matches(f, filter.Replace("*", ".*"), RegexOptions.IgnoreCase).Count > 0);
                }
            });
        }

        /// <summary>
        /// Event handler for the Regex radio button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void RbtnRegexSearch_Checked(object sender, RoutedEventArgs e)
        {
            if (this.RbtnRegexSearch.IsChecked.Value == true)
            {
                // For regex searches, unable to
                // 1. Match whole word
                this.ChkMatchWholeWord.IsEnabled = false;

                if (this.ChkMatchWholeWord.IsChecked.Value == true)
                {
                    this.ChkMatchWholeWord.IsChecked = false;
                }
            }
            else
            {
                this.ChkMatchWholeWord.IsEnabled = true;
                this.ChkHighlightResults.IsEnabled = true;
                this.ChkRegexMultiline.IsChecked = false;
            }

            this.ChkRegexMultiline.IsEnabled = this.RbtnRegexSearch.IsChecked.Value == true;
        }

        /// <summary>
        /// Remove files that will not be part of the current search.
        /// </summary>
        /// <returns>Task indicating the status of the operation.</returns>
        private async Task RemoveExclusionPaths()
        {
            this.filesToSearch.RemoveAll(f => this.filesToExclude.Any(excludeFile => excludeFile.ToUpper() == f.ToUpper()));

            await Task.Run(() =>
            {
                foreach (string dirToExclude in this.directoriesToExclude)
                {
                    string exclDir = dirToExclude.ToUpper();
                    this.filesToSearch.RemoveAll(f =>
                    {
                        try
                        {
                            return Path.GetDirectoryName(f).ToUpper().StartsWith(exclDir);
                        }
                        catch (PathTooLongException)
                        {
                            return false;
                        }
                    });
                }
            });
        }

        /// <summary>
        /// Reset the search progress on initiating a search.
        /// </summary>
        private void ResetSearch()
        {
            this.PgBarSearch.Value = 0;
            this.TblkProgress.Text = string.Empty;
            this.TxtResults.Document.Blocks.Clear();
            this.richTextboxParagraph = new Paragraph();
            this.searchStartTime = DateTime.Now;
            this.TblkProgressTime.Text = string.Format("{0:hh\\:mm\\:ss}", DateTime.Now.Subtract(this.searchStartTime));
            this.SetProgressInformation(string.Empty);
            this.SetFileCounterProgressInformation(0, Application.Current.Resources["GettingFilesToSearch"].ToString());
            this.TxtResults.Document = new FlowDocument(this.richTextboxParagraph);
            this.filesSearchProgress.Clear();
            this.BtnSearch.IsEnabled = false;
            this.BtnCancel.IsEnabled = true;
        }

        /// <summary>
        /// Saves all the search results to a file.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void SaveAllResultsToFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            if (!string.IsNullOrEmpty(this.CmbDirectory.Text))
            {
                sfd.InitialDirectory = this.CmbDirectory.Text.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                sfd.Filter = "txt files (*.txt)|*.txt";
            }

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                File.WriteAllText(sfd.FileName, new TextRange(this.TxtResults.Document.ContentStart, this.TxtResults.Document.ContentEnd).Text);
            }
        }

        /// <summary>
        /// Set the paths that will be always excluded.
        /// </summary>
        /// <param name="filesToAlwaysExclude">The list of file paths.</param>
        /// <param name="directoriesToAlwaysExclude">The list of directory paths.</param>
        private void SavePathsToAlwaysExclude(List<string> filesToAlwaysExclude, List<string> directoriesToAlwaysExclude)
        {
            PreferencesHandler.GetPreference("FilesToAlwaysExcludeFromSearch").FirstOrDefault().RemoveAll();
            PreferencesHandler.GetPreference("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().RemoveAll();

            filesToAlwaysExclude.ForEach(fae =>
            {
                if (File.Exists(fae))
                {
                    PreferencesHandler.GetPreference("FilesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fae));
                }
            });

            directoriesToAlwaysExclude.ForEach(fae =>
            {
                if (Directory.Exists(fae))
                {
                    PreferencesHandler.GetPreference("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fae));
                }
            });
        }

        /// <summary>
        /// Get the search result and display in text box.
        /// </summary>
        /// <param name="fileName">The file name to search.</param>
        /// <param name="termsToSearch">The terms to search.</param>
        /// <returns>Task object that searches and displays the result.</returns>
        private async Task<List<Inline>> SearchAndDisplayResult(string fileName, IEnumerable<string> termsToSearch)
        {
            List<Inline> retVal = new List<Inline>();
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                matchedLines = await Task.Run(() =>
                                              {
                                                  var longestRunningFile = GetLongestRunningFile(fileName);
                                                  SetProgressInformation(String.Format("{0}: {1}", Application.Current.Resources["Searching"], longestRunningFile));
                                                  return FileSearchHandlerFactory.Search(fileName, termsToSearch, matcherObj);
                                              },
                                              cancellationTokenSource.Token);
            }
            catch (Exception ex)
            {
                if (ex.Message != "A task was canceled.")
                {
                    SetSearchError(ex.Message);
                }
            }

            if (matchedLines != null && matchedLines.Count > 0)
            {
                // If a file returned multiple results (i.e. zip file with content), handle differently.
                if (matchedLines.Select(ml => ml.FileName).Distinct().Count() > 1)
                {
                    var groupedMatches = matchedLines.GroupBy(f => f.FileName, (key, g) => new { FileName = key, MatchLine = g.ToList() });

                    foreach (var groupedMatch in groupedMatches)
                    {
                        foreach (MatchedLine matchLine in groupedMatch.MatchLine)
                        {
                            matchLine.FileName = string.Empty;
                        }

                        groupedMatch.MatchLine[0].FileName = groupedMatch.FileName;
                        retVal.AddRange(await this.AddResult(groupedMatch.MatchLine));
                    }
                }
                else
                {
                    matchedLines = matchedLines.OrderBy(ml => ml.MatchId).ToList();
                    string resultFileName = matchedLines.Select(ml => ml.FileName).FirstOrDefault();
                    matchedLines.ForEach(ml => ml.FileName = string.Empty);
                    matchedLines[0].FileName = (resultFileName != null && resultFileName.Length > fileName.Length) ? resultFileName : fileName;
                    retVal = await this.AddResult(matchedLines);
                }
            }

            Interlocked.Increment(ref this.filesSearchedCounter);
            return retVal;
        }

        /// <summary>
        /// Search for matches using parallel processing.
        /// </summary>
        /// <param name="fileNamePaths">The filenames with path to search for.</param>
        /// <param name="termsToSearch">The terms to search for.</param>
        /// <returns>Task object to search files asynchronously</returns>
        private async Task SearchParallelAsync(IEnumerable<string> fileNamePaths, IEnumerable<string> termsToSearch)
        {
            this.filesSearchedCounter = 0;
            this.matcherObj.MatchWholeWord = this.matchWholeWord;
            this.matcherObj.IsRegexSearch = this.searchModeRegex;
            this.matcherObj.IsMultiLineRegex = this.multilineRegex;
            this.matcherObj.AllMatchesInFile = this.searchTypeAll;
            this.matcherObj.RegexOptions = this.regexOptions;
            this.matcherObj.CancellationTokenSource = this.cancellationTokenSource;
            List<Task> searchTasks = new List<Task>();

            try
            {
                foreach (string fileNamePath in fileNamePaths.AsParallel())
                {
                    if (!this.cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        FileSearchResult fileSearchResult = new FileSearchResult { FileNamePath = fileNamePath, SearchStartDateTimeTicks = DateTime.UtcNow.Ticks };
                        this.filesSearchProgress.TryAdd(fileSearchResult.FileNamePath, fileSearchResult);

                        Task searchTask = Task.Run(
                            async () =>
                            {
                                fileSearchResult.SearchMatches = await this.SearchAndDisplayResult(fileNamePath, termsToSearch);
                                await this.Dispatcher.Invoke(async () =>
                                {
                                    if (fileSearchResult.SearchMatches.Count > 0)
                                    {
                                        await this.UpdateResults(fileSearchResult);
                                    }

                                    this.TaskBarItemInfoProgress.ProgressValue = this.filesSearchedCounter * 1.0 / this.filesToSearch.Count();
                                    this.SetFileCounterProgressInformation(this.filesSearchedCounter, string.Format("{0} {1} {2} {3} ({4} %)", Application.Current.Resources["ProcessingFiles"].ToString(), this.filesSearchedCounter, Application.Current.Resources["Of"].ToString(), this.filesToSearch.Count(), (int)(this.filesSearchedCounter * 100) / this.filesToSearch.Count()));
                                });

                                this.filesSearchProgress.TryRemove(fileSearchResult.FileNamePath, out _);
                            },
                        this.cancellationTokenSource.Token);

                        searchTasks.Add(searchTask);
                    }
                }

                await Task.WhenAll(searchTasks);
            }
            finally
            {
                this.SetSearchCompletedDetails(this.filesSearchedCounter, fileNamePaths.Count(), this.matchesFound, this.filesWithMatch);
            }
        }

        /// <summary>
        /// Display time taken to search.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The ElapsedEventArgs object.</param>
        private void SearchTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        this.TblkProgressTime.Text = string.Format("{0:hh\\:mm\\:ss}", DateTime.Now.Subtract(this.searchStartTime));

                        string longestRunningFile = this.GetLongestRunningFile(string.Empty);
                        this.SetProgressInformation(string.Format("{0}: {1}", Application.Current.Resources["Searching"].ToString(), longestRunningFile));
                    });
                }
            }
            catch (TaskCanceledException)
            {
                this.searchTimer.Stop();

                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.cancellationTokenSource.Cancel();
                }
            }
        }

        /// <summary>
        /// Set the custom background colour for the application.
        /// </summary>
        private void SetApplicationCustomBackground()
        {
            string backGroundColourValue = PreferencesHandler.GetPreferenceValue("BackGroundColour");
            backGroundColourValue = backGroundColourValue.StartsWith("#") ? backGroundColourValue : "#" + backGroundColourValue;

            try
            {
                this.applicationBackColour = (SolidColorBrush)new BrushConverter().ConvertFromString(backGroundColourValue);
                this.Background = this.applicationBackColour;
                this.CmbDirectory.Background = this.applicationBackColour;
                this.CmbFindWhat.Background = this.applicationBackColour;
                this.CmbFilters.Background = this.applicationBackColour;
                this.TxtErrors.Background = this.applicationBackColour;
                this.TxtEditor.Background = this.applicationBackColour;
                this.DtpStartDate.Background = this.applicationBackColour;
                this.DtpEndDate.Background = this.applicationBackColour;
                this.TxtResults.Background = this.applicationBackColour;
            }
            catch (FormatException)
            {
                // Do nothing. Leave the background colour as default.
                PreferencesHandler.SetPreferenceValue("HighlightResultsColour", "#FFFFFF");
            }
        }

        /// <summary>
        /// Set readable content based on selected language.
        /// </summary>
        private void SetContentBasedOnLanguage()
        {
            this.LblDirectory.Content = Application.Current.Resources["Directory"].ToString();
            this.LblFindWhat.Content = Application.Current.Resources["FindKeywords"].ToString();
            this.LblFilters.Content = Application.Current.Resources["Filters"].ToString();
            this.ExpndrOptions.Header = Application.Current.Resources["Options"].ToString();
            this.GrpMatchOptions.Header = Application.Current.Resources["MatchOptions"].ToString();
            this.ChkMatchWholeWord.Content = Application.Current.Resources["WholeWord"].ToString();
            this.ChkMatchCase.Content = Application.Current.Resources["MatchCase"].ToString();
            this.GrpSearchMode.Header = Application.Current.Resources["SearchMode"].ToString();
            this.RbtnNormalSearch.Content = Application.Current.Resources["Normal"].ToString();
            this.RbtnRegexSearch.Content = Application.Current.Resources["Regex"].ToString();
            this.ChkRegexMultiline.Content = Application.Current.Resources["Multiline"].ToString();
            this.GrpMisc.Header = Application.Current.Resources["Misc"].ToString();
            this.ChkSearchSubfolders.Content = Application.Current.Resources["Subfolders"].ToString();
            this.ChkHighlightResults.Content = Application.Current.Resources["HighlightResults"].ToString();
            this.TblkEditor.Text = Application.Current.Resources["Editor"].ToString();
            this.TblkStartDate.Text = Application.Current.Resources["StartDate"].ToString();
            this.TblkEndDate.Text = Application.Current.Resources["EndDate"].ToString();
            this.TblkLanguage.Text = Application.Current.Resources["Language"].ToString();
            this.ChkShowMatchCount.Content = Application.Current.Resources["FileMatchCount"].ToString();
            this.BtnSearch.Content = Application.Current.Resources["Search"].ToString();
            this.BtnCancel.Content = Application.Current.Resources["Cancel"].ToString();
            this.CmbItemAny.Content = Application.Current.Resources["Any"].ToString();
            this.CmbItemAll.Content = Application.Current.Resources["All"].ToString();
            this.ChkExcludeFilters.Content = Application.Current.Resources["Exclude"].ToString();
            this.BtnAbout.Content = Application.Current.Resources["About"].ToString();

            string separatorTooltipMessage = string.Format("{0} '{1}'.", Application.Current.Resources["SeparateUsingCharacter"].ToString(), this.separatorCharacter);
            this.CmbDirectory.ToolTip = separatorTooltipMessage;
            this.CmbFindWhat.ToolTip = separatorTooltipMessage;
            this.CmbFilters.ToolTip = separatorTooltipMessage;
            this.ChkMatchWholeWord.ToolTip = Application.Current.Resources["TooltipMatchWholeWord"].ToString();
            this.ChkMatchCase.ToolTip = Application.Current.Resources["TooltipMatchCase"].ToString();
            this.RbtnNormalSearch.ToolTip = Application.Current.Resources["TooltipNormalSearch"].ToString();
            this.RbtnRegexSearch.ToolTip = Application.Current.Resources["TooltipSearchRegex"].ToString();
            this.ChkRegexMultiline.ToolTip = Application.Current.Resources["TooltipSearchRegexMultiLine"].ToString();
            this.ChkSearchSubfolders.ToolTip = Application.Current.Resources["TooltipSearchSubfolders"].ToString();
            this.ChkHighlightResults.ToolTip = Application.Current.Resources["TooltipHighlightMatches"].ToString();
            this.StkPnlEditor.ToolTip = Application.Current.Resources["TooltipEditorForFile"].ToString();
            this.StkChangeLanguage.ToolTip = Application.Current.Resources["TooltipChangeLanguage"].ToString();
            this.StkFileCreationEndDate.ToolTip = Application.Current.Resources["TooltipFileCreateEndDate"].ToString();
            this.StkFileCreationStartDate.ToolTip = Application.Current.Resources["TooltipFileCreateStartDate"].ToString();
            this.GrdProgressResult.ToolTip = Application.Current.Resources["TooltipOpenSearchedFileList"].ToString();
        }

        /// <summary>
        /// Set the default culture of the application if none found.
        /// </summary>
        private void SetDefaultCulture()
        {
            bool defaultCultureSet = false;
            this.culture = "en-US";     // Default culture if none found.

            foreach (ComboBoxItem item in this.CmbLanguage.Items)
            {
                if (item.Tag.ToString() == CultureInfo.CurrentUICulture.Name)
                {
                    this.culture = item.Tag.ToString();
                    this.CmbLanguage.SelectedItem = item;
                    defaultCultureSet = true;
                    break;
                }
            }

            if (!defaultCultureSet)
            {
                foreach (ComboBoxItem item in this.CmbLanguage.Items)
                {
                    if (item.Tag.ToString() == this.culture)
                    {
                        this.culture = item.Tag.ToString();
                        this.CmbLanguage.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Show the number of files that have been progressed.
        /// </summary>
        /// <param name="value">The current value of the progress.</param>
        /// <param name="text">The file name being searched for.</param>
        private void SetFileCounterProgressInformation(int value, string text)
        {
            try
            {
                if (value >= 0)
                {
                    this.PgBarSearch.Value = value;
                }

                this.TblkProgress.Text = text;
            }
            catch (TaskCanceledException)
            {
                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.cancellationTokenSource.Cancel();
                }
            }
        }

        /// <summary>
        /// Sets up the search options on initial form load.
        /// </summary>
        private void SetInitialSearchOptions()
        {
            PreferencesHandler.LoadPreferences();
            this.ChkMatchWholeWord.IsChecked = PreferencesHandler.GetPreferenceValue("MatchWholeWord").ToUpper() == true.ToString().ToUpper();
            this.ChkMatchCase.IsChecked = PreferencesHandler.GetPreferenceValue("MatchCase").ToUpper() == true.ToString().ToUpper();
            this.searchModeNormal = this.RbtnNormalSearch.IsChecked.Value == true;
            this.searchModeRegex = this.RbtnRegexSearch.IsChecked.Value == true;
            this.ChkSearchSubfolders.IsChecked = PreferencesHandler.GetPreferenceValue("SearchSubfolders").ToUpper() == true.ToString().ToUpper();
            this.ChkHighlightResults.IsChecked = PreferencesHandler.GetPreferenceValue("HighlightResults").ToUpper() == true.ToString().ToUpper();
            this.searchTypeAll = PreferencesHandler.GetPreferenceValue("SearchContentMode").ToUpper() == "All".ToUpper();
            this.showExecutionTime = PreferencesHandler.GetPreferenceValue("ShowExecutionTime").ToUpper() == true.ToString().ToUpper();
            this.culture = PreferencesHandler.GetPreferenceValue("Culture") ?? "en-US";
            this.ChkShowMatchCount.IsChecked = PreferencesHandler.GetPreferenceValue("ShowFileMatchCount").ToUpper() == true.ToString().ToUpper();
            this.CmbSearchType.SelectedItem = this.searchTypeAll ?
                this.CmbSearchType.Items.Cast<ComboBoxItem>().Where(i => i.Content.ToString().ToUpper() == Application.Current.Resources["All"].ToString().ToUpper()).First() :
                this.CmbSearchType.Items.Cast<ComboBoxItem>().First();

            this.SetSearchDate(this.DtpStartDate, "MinFileCreateSearchDate");
            this.SetSearchDate(this.DtpEndDate, "MaxFileCreateSearchDate");
            this.SetWindowDimensions();
            this.SetApplicationCustomBackground();
            this.SetResultHighlightColour();
            this.SetSearchItemsSeparator();
            this.SetLanguage();

            if (!int.TryParse(PreferencesHandler.GetPreferenceValue("MaxDropDownItems"), out this.maxDropDownItems))
            {
                this.maxDropDownItems = 10;
            }

            this.editorNamePath = PreferencesHandler.GetPreferenceValue("CustomEditor") ?? "notepad";
            this.TxtEditor.Text = Path.GetFileNameWithoutExtension(this.editorNamePath);
        }

        /// <summary>
        /// Set a jump list for the window.
        /// </summary>
        private void SetJumpList()
        {
            JumpTask jmpPreferences = new JumpTask();
            jmpPreferences.ApplicationPath = PreferencesHandler.PreferenceFilePath;
            jmpPreferences.IconResourcePath = @"C:\Windows\notepad.exe";
            jmpPreferences.Title = "Preferences";
            jmpPreferences.Description = Application.Current.Resources["OpenPreferencesFile"].ToString();
            jmpPreferences.CustomCategory = "Settings";

            JumpList jmpList = JumpList.GetJumpList(Application.Current);
            jmpList.JumpItems.Add(jmpPreferences);
            JumpList.AddToRecentCategory(jmpPreferences);
            jmpList.Apply();
        }

        /// <summary>
        /// Set language based on culture.
        /// </summary>
        private void SetLanguage()
        {
            this.SetResourceDictionary();

            foreach (ComboBoxItem item in this.CmbLanguage.Items)
            {
                if (item.Tag.ToString() == this.culture)
                {
                    this.CmbLanguage.SelectedItem = item;
                    break;
                }
            }

            CultureInfo cultureInfo = new CultureInfo(this.culture);
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
            this.matcherObj.CultureInfo = cultureInfo;

            this.SetContentBasedOnLanguage();
        }

        /// <summary>
        /// Set progress information showing the name of the file being progressed with.
        /// </summary>
        /// <param name="text">The progress text.</param>
        private void SetProgressInformation(string text)
        {
            try
            {
                this.Dispatcher.Invoke(() =>
                {
                    this.TblkProgressFile.Text = text;
                });
            }
            catch (TaskCanceledException)
            {
                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.cancellationTokenSource.Cancel();
                }
            }
        }

        /// <summary>
        /// Set the maximum value of the progress bar control.
        /// </summary>
        /// <param name="value">The maximum value of the progress bar.</param>
        private void SetProgressMaxValue(int value)
        {
            this.Dispatcher.Invoke(() =>
            {
                this.PgBarSearch.Maximum = value;
            });
        }

        /// <summary>
        /// Sets up the resource dictionary for language.
        /// </summary>
        private void SetResourceDictionary()
        {
            List<ResourceDictionary> resourceDictionaries = Application.Current.Resources.MergedDictionaries.ToList();
            string requestedCultureDictionary = string.Format("Resources/StringResources.{0}.xaml", this.culture);
            ResourceDictionary resourceDictionary = resourceDictionaries.FirstOrDefault(d => d.Source.OriginalString == requestedCultureDictionary);

            if (resourceDictionary == null)
            {
                requestedCultureDictionary = "StringResources.xaml";
                resourceDictionary = resourceDictionaries.FirstOrDefault(d => d.Source.OriginalString == requestedCultureDictionary);
            }

            if (resourceDictionary != null)
            {
                Application.Current.Resources.MergedDictionaries.Remove(resourceDictionary);
                Application.Current.Resources.MergedDictionaries.Add(resourceDictionary);
            }
        }

        /// <summary>
        /// Set the highlight colour to be used for matching results.
        /// </summary>
        private void SetResultHighlightColour()
        {
            string highlightResultColourValue = PreferencesHandler.GetPreferenceValue("HighlightResultsColour");
            highlightResultColourValue = highlightResultColourValue.StartsWith("#") ? highlightResultColourValue : "#" + highlightResultColourValue;        // Prepend with "#" if required.

            try
            {
                SolidColorBrush newColour = (SolidColorBrush)new BrushConverter().ConvertFromString(highlightResultColourValue);
                this.highlightResultBackColour = newColour;
            }
            catch (FormatException)
            {
                this.highlightResultBackColour = Brushes.PeachPuff;
            }

            PreferencesHandler.SetPreferenceValue("HighlightResultsColour", this.highlightResultBackColour.ToString());
        }

        /// <summary>
        /// Set details after the search has been completed
        /// </summary>
        /// <param name="filesProcessedCount">The number files that were processed for the search.</param>
        /// <param name="filesProcessedCountOriginal">The number of files that were to be processed for the search.</param>
        /// <param name="matchesFound">The total number of matches.</param>
        /// <param name="filesWithMatches">The total number of files with matches.</param>
        private void SetSearchCompletedDetails(int filesProcessedCount, int filesProcessedCountOriginal, int matchesFound, int filesWithMatches)
        {
            if (filesProcessedCountOriginal > 0)
            {
                this.SetFileCounterProgressInformation(filesProcessedCount, string.Format("{0} {1} {2} {3} ({4} %).", Application.Current.Resources["ProcessingFiles"].ToString(), filesProcessedCount, Application.Current.Resources["Of"].ToString(), filesProcessedCountOriginal, (int)(filesProcessedCount * 100) / filesProcessedCountOriginal));
            }

            string progressMessage = string.Format("{0} {1} {2} {3} {4} {5}.", Application.Current.Resources["Found"].ToString(), matchesFound, Application.Current.Resources["Matches"].ToString(), Application.Current.Resources["In"].ToString(), filesWithMatches, Application.Current.Resources["Files"].ToString());

            if (this.cancellationTokenSource.Token.IsCancellationRequested)
            {
                progressMessage = Application.Current.Resources["SearchCancelled"].ToString() + ". " + progressMessage;
            }

            this.SetProgressInformation(progressMessage);
        }

        /// <summary>
        /// Set the dates used by the application.
        /// </summary>
        /// <param name="datePicker">The date picker control.</param>
        /// <param name="preferenceElement">The preference element to use.</param>
        private void SetSearchDate(DatePicker datePicker, string preferenceElement)
        {
            if (PreferencesHandler.PreferencesFile != null && PreferencesHandler.GetPreference(preferenceElement) != null)
            {
                string strDateToSet = PreferencesHandler.GetPreferenceValue(preferenceElement);
                DateTime dateToSet;

                if (DateTime.TryParse(strDateToSet, out dateToSet))
                {
                    datePicker.SelectedDate = dateToSet;
                }
            }
        }

        /// <summary>
        /// Show any error on attempting to find matches.
        /// </summary>
        /// <param name="text">The error incurred on searching.</param>
        private void SetSearchError(string text)
        {
            try
            {
                this.Dispatcher.Invoke(() =>
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        this.TxtErrors.Text = string.Empty;
                    }
                    else
                    {
                        this.TxtErrors.Text += text + Environment.NewLine;
                    }

                    this.GrdRowErrors.Height = (GridLength)new GridLengthConverter().ConvertFromString(string.IsNullOrEmpty(this.TxtErrors.Text) ? "auto" : "2*");
                    this.BrdrErrors.Visibility = string.IsNullOrEmpty(this.TxtErrors.Text) ? Visibility.Collapsed : Visibility.Visible;
                });
            }
            catch (TaskCanceledException)
            {
                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.cancellationTokenSource.Cancel();
                }
            }
        }

        /// <summary>
        /// Sets the maximum lengths for the search inputs.
        /// </summary>
        /// <param name="comboBox">The combo box object whose input is to be restricted.</param>
        /// <param name="maxAllowedLength">The maximum input length allowed for the combo box object.</param>
        private void SetSearchInputMaxLengths(ComboBox comboBox, int maxAllowedLength)
        {
            TextBox editableTextBox = (TextBox)comboBox.Template.FindName("PART_EditableTextBox", comboBox);
            if (editableTextBox != null)
            {
                editableTextBox.MaxLength = maxAllowedLength;
            }
        }

        /// <summary>
        /// Set up the separator character to be used when searching.
        /// </summary>
        private void SetSearchItemsSeparator()
        {
            if (PreferencesHandler.GetPreference("SeparatorCharacter") != null && PreferencesHandler.GetPreference("SeparatorCharacter").FirstOrDefault() != null && !string.IsNullOrWhiteSpace(PreferencesHandler.GetPreferenceValue("SeparatorCharacter")))
            {
                this.separatorCharacter = PreferencesHandler.GetPreferenceValue("SeparatorCharacter").Substring(0, 1);

                string separatorTooltipMessage = string.Format("{0} '{1}'.", Application.Current.Resources["SeparateUsingCharacter"].ToString(), this.separatorCharacter);
                this.CmbDirectory.ToolTip = separatorTooltipMessage;
                this.CmbFindWhat.ToolTip = separatorTooltipMessage;
                this.CmbFilters.ToolTip = separatorTooltipMessage;
            }
        }

        /// <summary>
        /// Set up the search parameters to initiate a search.
        /// </summary>
        private void SetSearchParameters()
        {
            this.TxtErrors.Text = string.Empty;
            this.matchWholeWord = this.ChkMatchWholeWord.IsChecked.Value == true;
            this.matchCase = this.ChkMatchCase.IsChecked.Value == true;
            this.searchModeNormal = this.RbtnNormalSearch.IsChecked.Value == true;
            this.searchModeRegex = this.RbtnRegexSearch.IsChecked.Value == true;
            this.searchSubFolders = this.ChkSearchSubfolders.IsChecked.Value == true;
            this.highlightResults = this.ChkHighlightResults.IsChecked.Value == true;
            this.searchTypeAll = this.CmbFindWhat.Text.Contains(this.separatorCharacter) && ((ComboBoxItem)this.CmbSearchType.SelectedValue).Content.ToString().ToUpper() == Application.Current.Resources["All"].ToString().ToUpper();
            this.multilineRegex = this.ChkRegexMultiline.IsChecked.Value == true;
            this.minSearchDate = this.DtpStartDate.SelectedDate.HasValue ? this.DtpStartDate.SelectedDate.Value : DateTime.MinValue;
            this.maxSearchDate = this.DtpEndDate.SelectedDate.HasValue ? this.DtpEndDate.SelectedDate.Value.Add(new TimeSpan(23, 59, 59)) : DateTime.MaxValue;   // Add time else uses 00:00:00
            this.showMatchCount = this.ChkShowMatchCount.IsChecked.Value == true;
            this.filterExclusionSet = this.ChkExcludeFilters.IsChecked.Value == true;
            this.SetSearchError(string.Empty);
            this.matchesFound = 0;
            this.filesWithMatch = 0;

            this.AddSearchCriteria(this.CmbDirectory, this.CmbDirectory.Text);
            this.AddSearchCriteria(this.CmbFindWhat, this.CmbFindWhat.Text);
            this.AddSearchCriteria(this.CmbFilters, this.CmbFilters.Text);

            this.UpdateDropdownListOrder(this.CmbDirectory);
            this.UpdateDropdownListOrder(this.CmbFindWhat);
            this.UpdateDropdownListOrder(this.CmbFilters);
        }

        /// <summary>
        /// Set up the terms to search based on the search criteria.
        /// </summary>
        /// <param name="searchText">The terms to be searched (This will be split based on ";" characters in the string).</param>
        /// <returns>List of terms to search based on the search criteria.</returns>
        private List<string> SetTermsToSearchText(string searchText)
        {
            List<string> termsToSearch = searchText.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).Distinct().ToList();
            termsToSearch.RemoveAll(str => string.IsNullOrEmpty(str));
            bool isRegexEscaped = false;

            if (!this.searchModeRegex)
            {
                termsToSearch = termsToSearch.Select(st => st = Regex.Escape(st)).ToList();     // If not regex, escape the characters to perform exact search.
                isRegexEscaped = true;
            }

            if (this.matchWholeWord)
            {
                termsToSearch = termsToSearch.Select(st => st = "\\b" + (isRegexEscaped ? st : Regex.Escape(st)) + "\\b").ToList();
            }

            return termsToSearch;
        }

        /// <summary>
        /// Sets the information to be displayed on the title.
        /// </summary>
        private void SetTitleInfo()
        {
            this.Title = this.GetRunningUserInfo();
        }

        /// <summary>
        /// Setup the application without the preferences file.
        /// </summary>
        private void SetupAppWithoutPreferences()
        {
            this.BtnChangeEditor.Visibility = Visibility.Collapsed;
            this.SetDefaultCulture();
        }

        /// <summary>
        /// Set up the height and width of the search window.
        /// </summary>
        private void SetWindowDimensions()
        {
            string strWindowHeight = PreferencesHandler.GetPreferenceValue("WindowHeight");
            string strWindowWidth = PreferencesHandler.GetPreferenceValue("WindowWidth");
            string strPopupWindowHeight = PreferencesHandler.GetPreferenceValue("PopupWindowHeight");
            string strPopupWindowWidth = PreferencesHandler.GetPreferenceValue("PopupWindowWidth");
            string strPopupWindowTimeoutSeconds = PreferencesHandler.GetPreferenceValue("PopupWindowTimeoutSeconds");

            if (double.TryParse(strWindowHeight, out this.windowHeight))
            {
                this.Height = this.windowHeight;
            }

            if (double.TryParse(strWindowWidth, out this.windowWidth))
            {
                this.Width = this.windowWidth;
            }

            int.TryParse(strPopupWindowWidth, out this.popupWindowWidth);
            int.TryParse(strPopupWindowHeight, out this.popupWindowHeight);
            int.TryParse(strPopupWindowTimeoutSeconds, out this.popupWindowTimeoutSeconds);
        }

        /// <summary>
        /// Show any error on attempting to read file or find matches.
        /// </summary>
        /// <param name="ex">The incurred exception object.</param>
        private void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, Application.Current.Resources["ErrorOccurred"].ToString());
        }

        /// <summary>
        /// Show an error as a quick popup, and hide it afterwards.
        /// </summary>
        /// <param name="errorMessage">The error message to display.</param>
        private void ShowErrorPopup(string errorMessage)
        {
            this.SetSearchError(errorMessage);

            Task.Delay(5000).ContinueWith((t) =>
            {
                this.SetSearchError(string.Empty);
            });
        }

        /// <summary>
        /// Handler to show file content as popup.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void TxtResults_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.SetSearchError(string.Empty);
            int lineNo = this.GetLineNumberForPopup();

            if (lineNo > 0)
            {
                string file = this.GetFileForPopup();

                if (this.contentPopup != null)
                {
                    this.contentPopup.Close();
                }

                if (!string.IsNullOrWhiteSpace(file) && Common.IsAsciiSearch(file))
                {
                    try
                    {
                        this.contentPopup = new ContentPopup(file, lineNo);
                        this.contentPopup.Width = this.popupWindowWidth;
                        this.contentPopup.Height = this.popupWindowHeight;
                        this.contentPopup.WindowCloseTimeoutSeconds = (this.popupWindowTimeoutSeconds < 2 || this.popupWindowTimeoutSeconds > 20)
                            ? 4
                            : this.popupWindowTimeoutSeconds;
                        this.contentPopup.Show();
                        this.contentPopup.Closed += this.ContentPopup_Closed;
                    }
                    catch (FileNotFoundException fnfe)
                    {
                        this.SetSearchError(fnfe.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Zoom the content of the results window.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseWheelEventArgs object.</param>
        private void TxtResults_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                bool zoomIn = e.Delta > 0;
                bool scaleChanged = false;

                if (zoomIn && this.Scale < MaxScaleValue)
                {
                    scaleChanged = true;
                    this.Scale += 0.1;
                }
                else if (!zoomIn && this.Scale > MinScaleValue)
                {
                    scaleChanged = true;
                    this.Scale -= 0.1;
                }

                if (scaleChanged)
                {
                    this.TxtBlkScaleValue.Visibility = Visibility.Visible;

                    this.zoomLabelTimer.Interval = new TimeSpan(0, 0, 2);
                    this.zoomLabelTimer.Tick += this.HideScaleTextBlock;
                    this.zoomLabelTimer.Start();
                }
            }
        }

        /// <summary>
        /// Update the content of the drop down combo box.
        /// </summary>
        /// <param name="comboBox">The combo box control object.</param>
        private void UpdateDropdownListOrder(ComboBox comboBox)
        {
            if (comboBox.SelectedItem != null)
            {
                List<string> directories = new List<string>();
                string selectedItem = comboBox.Text;
                directories.Add(selectedItem);

                foreach (var item in comboBox.Items)
                {
                    if (item.ToString() != selectedItem)
                    {
                        directories.Add(item.ToString());
                    }
                }

                comboBox.Items.Clear();
                directories.ForEach(s => comboBox.Items.Add(s));
                comboBox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Update the UI with results from search.
        /// </summary>
        /// <param name="fileSearchResult">The FileSearchResult object that contains one or more matches.</param>
        /// <returns>Task to update the results.</returns>
        private async Task UpdateResults(FileSearchResult fileSearchResult)
        {
            await this.AddResultsToTextbox(fileSearchResult.SearchMatches);
        }

        /// <summary>
        /// Handle window closing event to save preferences.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The CancelEventArgs object.</param>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (shouldSavePreferences && PreferencesHandler.PreferencesFile != null)
            {
                PreferencesHandler.SetPreferenceValue("MatchWholeWord",  this.ChkMatchWholeWord.IsChecked.Value.ToString());
                PreferencesHandler.SetPreferenceValue("MatchCase", this.ChkMatchCase.IsChecked.Value.ToString());
                PreferencesHandler.SetPreferenceValue("SearchSubfolders", this.ChkSearchSubfolders.IsChecked.Value.ToString());
                PreferencesHandler.SetPreferenceValue("HighlightResults", this.ChkHighlightResults.IsChecked.Value.ToString());
                PreferencesHandler.SetPreferenceValue("MinFileCreateSearchDate", (this.DtpStartDate.SelectedDate != null && this.DtpStartDate.SelectedDate.HasValue) ? this.DtpStartDate.SelectedDate.Value.ToString() : string.Empty);
                PreferencesHandler.SetPreferenceValue("MaxFileCreateSearchDate", (this.DtpEndDate.SelectedDate != null && this.DtpEndDate.SelectedDate.HasValue) ? this.DtpEndDate.SelectedDate.Value.ToString() : string.Empty);
                PreferencesHandler.SetPreferenceValue("WindowHeight", this.Height.ToString());
                PreferencesHandler.SetPreferenceValue("WindowWidth", this.Width.ToString());
                PreferencesHandler.SetPreferenceValue("PopupWindowHeight", this.popupWindowHeight.ToString());
                PreferencesHandler.SetPreferenceValue("PopupWindowWidth", this.popupWindowWidth.ToString());
                PreferencesHandler.SetPreferenceValue("WindowWidth", this.Width.ToString());
                PreferencesHandler.SetPreferenceValue("Culture", this.culture.ToString());
                PreferencesHandler.SetPreferenceValue("SearchContentMode", this.searchTypeAll ? "All" : "Any");
                PreferencesHandler.SetPreferenceValue("ShowFileMatchCount", this.ChkShowMatchCount.IsChecked.Value.ToString());
                this.AddItemsToPreferences(this.CmbDirectory, "SearchDirectories");
                this.AddItemsToPreferences(this.CmbFindWhat, "SearchContents");
                this.AddItemsToPreferences(this.CmbFilters, "SearchFilters");

                try
                {
                    if (this.cancellationTokenSource != null && !this.cancellationTokenSource.IsCancellationRequested)
                    {
                        this.cancellationTokenSource.Cancel();
                    }

                    PreferencesHandler.SavePreferences();
                }
                catch (Exception ex)
                {
                    this.ShowError(ex);
                }
            }

            int totalChildWindows = this.childWindows.Count;

            for (int counter = 0; counter < totalChildWindows; counter++)
            {
                this.childWindows[0].Close();       // Using index 0 as the list is reduced each time a child window is closed.
            }
        }

        /// <summary>
        /// Handler for window key down events
        /// </summary>
        /// <param name="sender">The window object.</param>
        /// <param name="e">The key event args object.</param>
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Set focus on search combo box if Ctrl+F is clicked.
            if ((Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && e.Key == Key.F)
            {
                this.CmbFindWhat.Focus();
            }
        }

        /// <summary>
        /// Sets up the controls and validations after load.
        /// </summary>
        /// <param name="sender">The window object.</param>
        /// <param name="e">The routed event args object.</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.SetSearchInputMaxLengths(this.CmbDirectory, MaxSearchDirectoryTextLength);
            this.SetSearchInputMaxLengths(this.CmbFindWhat, MaxSearchTextLength);
            this.SetSearchInputMaxLengths(this.CmbFilters, MaxSearchFilterTextLength);
            this.CmbFindWhat.Focus();
        }

        #endregion Private Methods
    }
}
