// <copyright file="SearchWindow.xaml.cs" company="dennjose">
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
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Navigation;
    using System.Windows.Shell;
    using System.Xml.Linq;
    using SearcherLibrary;

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
        private const int MaxSearchTextLength = 100;

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
        /// The full path of the preferences file.
        /// </summary>
        private string filePreferencesName = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SearcherPreferences.xml");

        /// <summary>
        /// Private store for the list of files that have already been searched.
        /// </summary>
        private List<string> filesSearched;

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
        /// Private field to store class that allows searching NON-ASCII files.
        /// </summary>
        private SearchOtherExtensions otherExtensions = new SearchOtherExtensions();

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
        /// The preferences file in XML format.
        /// </summary>
        private XDocument preferenceFile = null;

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
        /// Boolean indicating whether the search should use parallel processing.
        /// </summary>
        private bool searchModeParallel = true;

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
        private System.Timers.Timer searchTimer = new System.Timers.Timer(1000);

        /// <summary>
        /// Boolean indicating whether the search must contain all of the search terms.
        /// </summary>
        private bool searchTypeAll = false;

        /// <summary>
        /// The separator character to identify individual terms;
        /// </summary>
        private string separatorCharacter = ";";

        /// <summary>
        /// Boolean indicating whether the search should show the execution time.
        /// </summary>
        private bool showExecutionTime = false;

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
        private System.Windows.Threading.DispatcherTimer zoomLabelTimer = new System.Windows.Threading.DispatcherTimer();

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchWindow"/> class.
        /// </summary>
        public SearchWindow()
        {
            this.InitializeComponent();
            this.InitialiseControls();
            bool allowDummySearch = false;
            this.DataContext = this;

            if (allowDummySearch)
            {
                this.SetDummySearch();
            }
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
            this.preferenceFile.Descendants(preferenceElement).Descendants().Remove();

            for (int counter = 0; counter < comboBox.Items.Count; counter++)
            {
                this.preferenceFile.Descendants(preferenceElement).FirstOrDefault().Add(new XElement("Value", comboBox.Items[counter].ToString()));
            }
        }

        /// <summary>
        /// Adds the historic search preferences if they have been saved.
        /// </summary>
        /// <param name="comboBox">The combo box object.</param>
        /// <param name="preferenceElement">The search preference value.</param>
        private void AddPreferencesToItems(ComboBox comboBox, string preferenceElement)
        {
            this.preferenceFile.Descendants(preferenceElement).Descendants("Value").ToList().ForEach(p =>
            {
                comboBox.Items.Add(p.Value);
            });
        }

        /// <summary>
        /// Add the result to the display.
        /// </summary>
        /// <param name="matchedLines">The list of match objects to display matches</param>
        private void AddResult(List<MatchedLine> matchedLines)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    InlineCollection inlines = new Paragraph().Inlines;
                    string overlapFileName = string.Empty;

                    matchedLines.ForEach(ml =>
                    {
                        if (!ml.DisplayProcessed)
                        {
                            if (!string.IsNullOrEmpty(ml.FileName) && !this.filesSearched.Contains(ml.FileName))
                            {
                                filesWithMatch++;
                                filesSearched.Add(ml.FileName);
                                overlapFileName = ml.FileName;
                                Hyperlink link = new Hyperlink(new Run(Environment.NewLine + ml.FileName + Environment.NewLine));
                                link.NavigateUri = new Uri(ml.FileName);
                                link.RequestNavigate += Link_RequestNavigate;
                                link.PreviewMouseRightButtonDown += Link_PreviewMouseRightButtonDown;
                                inlines.Add(new Bold(link)
                                {
                                    Foreground = Brushes.CornflowerBlue
                                });
                            }

                            List<string> contentArray = new List<string>();

                            List<MatchedLine> otherMatches = matchedLines.Where(oml => oml.LineNumber == ml.LineNumber && oml.Content == ml.Content && oml.StartIndex != ml.StartIndex && oml.DisplayProcessed == false).OrderBy(oml => oml.StartIndex).ToList();

                            contentArray.Add(ml.Content.Substring(0, ml.StartIndex));
                            contentArray.Add(ml.Content.Substring(ml.StartIndex, ml.Length));

                            if (otherMatches.Count > 0)
                            {
                                matchesFound += otherMatches.Count;

                                while (otherMatches.Count > 0)
                                {
                                    try
                                    {
                                        contentArray.Add(otherMatches[0].Content.Substring(string.Join(string.Empty, contentArray).Length, otherMatches[0].StartIndex - string.Join(string.Empty, contentArray).Length));
                                        contentArray.Add(otherMatches[0].Content.Substring(otherMatches[0].StartIndex, otherMatches[0].Length));

                                        if (otherMatches.Count == 1)
                                        {
                                            contentArray.Add(otherMatches[0].Content.Substring(string.Join(string.Empty, contentArray).Length));
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException)
                                    {
                                        this.SetSearchError(string.Format("Search term overlap on line {0} in file {1}.", ml.LineNumber, overlapFileName));
                                        contentArray.Add(otherMatches[0].Content.Substring(string.Join(string.Empty, contentArray).Length));
                                    }

                                    otherMatches[0].DisplayProcessed = true;
                                    otherMatches.Remove(otherMatches[0]);
                                }
                            }
                            else
                            {
                                contentArray.Add(ml.Content.Substring(ml.StartIndex + ml.Length, ml.Content.Length - (ml.StartIndex + ml.Length)));
                            }

                            matchesFound++;

                            for (int counter = 0; counter < contentArray.Count; counter++)
                            {
                                if (counter % 2 == 1)
                                {
                                    inlines.Add(new Run(Regex.Unescape(Regex.Escape(contentArray[counter])))
                                    {
                                        Background = this.highlightResults ? this.highlightResultBackColour : this.applicationBackColour
                                    });
                                }
                                else
                                {
                                    inlines.Add(contentArray[counter]);
                                }
                            }

                            ml.DisplayProcessed = true;
                            inlines.Add(Environment.NewLine);
                        }
                    });

                    this.richTextboxParagraph.Inlines.AddRange(inlines.ToList());
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
                this.ShowErrorPopup("Unable to exclude file. Archived or zipped files cannot be excluded. To be handled in a future version.");
            }
            else
            {
                if (!this.filesToExclude.Any(f => f.ToUpper() == fullFileName.ToUpper()) && System.IO.File.Exists(fullFileName))
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

            if (this.preferenceFile != null)
            {
                saw.AppHasPreferencesFile = this.preferenceFile != null;
                saw.CanCheckForUpdates = this.preferenceFile.Descendants("CheckForUpdates").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();
            }

            saw.Owner = this;
            saw.ShowDialog();

            if (this.preferenceFile != null)
            {
                this.preferenceFile.Descendants("CheckForUpdates").FirstOrDefault().Value = saw.CanCheckForUpdates.ToString();
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
                this.SetFileCounterProgressInformation(-1, string.Format("Search cancelled."));
            }
        }

        /// <summary>
        /// Event handler for the changing the editor.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnChangeEditor_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "Executable Files (*.exe)|*.exe";
            ofd.Multiselect = false;
            ofd.InitialDirectory = @"C:\Windows\System32";
            ofd.RestoreDirectory = true;
            ofd.Title = "Select editor";

            if (this.preferenceFile.Descendants("CustomEditor") != null && this.preferenceFile.Descendants("CustomEditor").Count() == 1
                && this.preferenceFile.Descendants("CustomEditor").FirstOrDefault() != null && !string.IsNullOrWhiteSpace(this.preferenceFile.Descendants("CustomEditor").FirstOrDefault().Value))
            {
                try
                {
                    DirectoryInfo parentDir = System.IO.Directory.GetParent(this.preferenceFile.Descendants("CustomEditor").FirstOrDefault().Value);
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
                this.preferenceFile.Descendants("CustomEditor").FirstOrDefault().Value = this.editorNamePath;
            }
        }

        /// <summary>
        /// Event handler for the set directory button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnDirectory_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
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
        private void BtnSearch_Click(object sender, RoutedEventArgs e)
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
                this.executionTime.Start();
                this.searchTimer.Start();
                Task.Factory.StartNew(
                () =>
                {
                    try
                    {
                        this.PerformSearch(searchText, searchPath, filters);
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
                },
                this.cancellationTokenSource.Token)
                .ContinueWith((t) =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        this.executionTime.Stop();
                        this.searchTimer.Stop();
                        this.BtnCancel.IsEnabled = false;
                        this.BtnSearch.IsEnabled = true;
                        this.EnableSearchControls(true);
                        this.CmbFindWhat.Focus();
                        string elapsedTime = this.executionTime.Elapsed.ToString();
                        this.executionTime.Reset();

                        if (this.showExecutionTime)
                        {
                            MessageBox.Show(string.Format("Time taken for search: {0}", elapsedTime), "Search complete", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
                        }
                    });
                });
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
                MessageBox.Show("Please specify directory and contents to search.");
                retVal = false;
            }
            else if (this.DtpEndDate.SelectedDate != null && this.DtpEndDate.SelectedDate.HasValue && this.DtpStartDate.SelectedDate != null && this.DtpStartDate.SelectedDate.HasValue)
            {
                if (this.DtpStartDate.SelectedDate.Value > this.DtpEndDate.SelectedDate.Value)
                {
                    MessageBox.Show("Start date cannot be later than end date.");
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
                        "Unable to access directory path(s) or path(s) cannot be found:{0}{1}{2}{3}",
                        Environment.NewLine,
                        string.Join(Environment.NewLine, errorPaths.Take(5).ToArray()),
                        Environment.NewLine,
                        "If using a specific character separator via configuration, please ensure it is used when searching.");

                    MessageBox.Show(errorMessage);
                    retVal = false;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Check if any missing xml elements are there from the default settings. If yes, create the missing elements. Generally done for upgrade. Not ideal, but avoids XSD.
        /// </summary>
        private void CheckPreferencesFile()
        {
            XDocument prefFile = null;

            try
            {
                prefFile = XDocument.Load(this.filePreferencesName);
            }
            catch (IOException ioe)
            {
                throw new IOException(string.Format("Unable to load preferences file {0}. {1}", this.filePreferencesName, ioe.Message));
            }

            if (prefFile != null)
            {
                IOrderedEnumerable<string> defaultNodes = this.CreatePreferencesFile().Descendants().Select(t => t.Name.LocalName).OrderBy(t => t);
                IOrderedEnumerable<string> prefFileNodes = prefFile.Descendants().Select(t => t.Name.LocalName).OrderBy(t => t);
                List<string> missingNodes = defaultNodes.Except(prefFileNodes).ToList();

                if (missingNodes.Count > 0)
                {
                    foreach (string node in missingNodes)
                    {
                        if (node == "MatchWholeWord")
                        {
                            prefFile.Root.Add(new XElement("MatchWholeWord", false));
                        }
                        else if (node == "MatchCase")
                        {
                            prefFile.Root.Add(new XElement("MatchCase", false));
                        }
                        else if (node == "MaxDropDownItems")
                        {
                            prefFile.Root.Add(new XElement("MaxDropDownItems", 10));
                        }
                        else if (node == "SearchSubfolders")
                        {
                            prefFile.Root.Add(new XElement("SearchSubfolders", true));
                        }
                        else if (node == "HighlightResults")
                        {
                            prefFile.Root.Add(new XElement("HighlightResults", true));
                        }
                        else if (node == "MinFileCreateSearchDate")
                        {
                            prefFile.Root.Add(new XElement("MinFileCreateSearchDate", DateTime.MinValue));
                        }
                        else if (node == "MaxFileCreateSearchDate")
                        {
                            prefFile.Root.Add(new XElement("MaxFileCreateSearchDate", DateTime.MaxValue));
                        }
                        else if (node == "ShowExecutionTime")
                        {
                            prefFile.Root.Add(new XElement("ShowExecutionTime", false));
                        }
                        else if (node == "SeparatorCharacter")
                        {
                            prefFile.Root.Add(new XElement("SeparatorCharacter", ";"));
                        }
                        else if (node == "SearchContentMode")
                        {
                            prefFile.Root.Add(new XElement("SearchContentMode", "Any"));
                        }
                        else if (node == "BackGroundColour")
                        {
                            prefFile.Root.Add(new XElement("BackGroundColour", "#FFFFFF"));
                        }
                        else if (node == "HighlightResultsColour")
                        {
                            prefFile.Root.Add(new XElement("HighlightResultsColour", "#FFDAB9"));
                        }
                        else if (node == "CustomEditor")
                        {
                            prefFile.Root.Add(new XElement("CustomEditor", string.Empty));
                        }
                        else if (node == "CheckForUpdates")
                        {
                            prefFile.Root.Add(new XElement("CheckForUpdates", true));
                        }
                        else if (node == "LastUpdateCheckDate")
                        {
                            prefFile.Root.Add(new XElement("LastUpdateCheckDate", DateTime.Today.AddMonths(-1).ToShortDateString()));
                        }
                        else if (node == "WindowHeight")
                        {
                            prefFile.Root.Add(new XElement("WindowHeight", this.MinHeight));
                        }
                        else if (node == "WindowWidth")
                        {
                            prefFile.Root.Add(new XElement("WindowWidth", this.MinWidth));
                        }
                        else if (node == "PopupWindowHeight")
                        {
                            prefFile.Root.Add(new XElement("PopupWindowHeight", 300));
                        }
                        else if (node == "PopupWindowWidth")
                        {
                            prefFile.Root.Add(new XElement("PopupWindowWidth", 500));
                        }
                        else if (node == "PopupWindowTimeoutSeconds")
                        {
                            prefFile.Root.Add(new XElement("PopupWindowTimeoutSeconds", 4));
                        }
                        else if (node == "SearchDirectories")
                        {
                            prefFile.Root.Add(new XElement("SearchDirectories", new XElement[] { null }));
                        }
                        else if (node == "SearchContents")
                        {
                            prefFile.Root.Add(new XElement("SearchContents", new XElement[] { null }));
                        }
                        else if (node == "SearchFilters")
                        {
                            prefFile.Root.Add(new XElement("SearchFilters", new XElement[] { null }));
                        }
                        else if (node == "FilesToAlwaysExcludeFromSearch")
                        {
                            prefFile.Root.Add(new XElement("FilesToAlwaysExcludeFromSearch", new XElement[] { null }));
                        }
                        else if (node == "DirectoriesToAlwaysExcludeFromSearch")
                        {
                            prefFile.Root.Add(new XElement("DirectoriesToAlwaysExcludeFromSearch", new XElement[] { null }));
                        }
                    }

                    prefFile.Save(this.filePreferencesName);
                }
            }
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
        /// Create a preference file.
        /// </summary>
        /// <returns>The search preference file.</returns>
        private XDocument CreatePreferencesFile()
        {
            // Not bothering with XSD, as this is a one-off config operation, and not used for data exchange with other systems.
            XElement[] initialPreferences = new XElement[]
            {
                new XElement("MatchWholeWord", false),
                new XElement("MatchCase", false),
                new XElement("MaxDropDownItems", 10),
                new XElement("SearchSubfolders", true),
                new XElement("HighlightResults", true),
                new XElement("MinFileCreateSearchDate", DateTime.MinValue),
                new XElement("MaxFileCreateSearchDate", DateTime.MaxValue),
                new XElement("SearchContentMode", "Any"),
                new XElement("ShowExecutionTime", false),
                new XElement("SeparatorCharacter", ";"),
                new XElement("BackGroundColour", "#FFFFFF"),
                new XElement("HighlightResultsColour", "#FFDAB9"),        // Brushes.PeachPuff
                new XElement("CustomEditor", string.Empty),
                new XElement("CheckForUpdates", true),
                new XElement("LastUpdateCheckDate", DateTime.Today.AddMonths(-1).ToShortDateString()),
                new XElement("WindowHeight", this.MinHeight),
                new XElement("WindowWidth", this.MinWidth),
                new XElement("PopupWindowHeight", 300),
                new XElement("PopupWindowWidth", 500),
                new XElement("PopupWindowTimeoutSeconds", 4),
                new XElement("SearchDirectories", new XElement[] { null }),
                new XElement("SearchContents", new XElement[] { null }),
                new XElement("SearchFilters", new XElement[] { null }),
                new XElement("FilesToAlwaysExcludeFromSearch", new XElement[] { null }),
                new XElement("DirectoriesToAlwaysExcludeFromSearch", new XElement[] { null })
            };

            XDocument retVal = XDocument.Parse(new XElement("SearcherPreferences", initialPreferences).ToString(), LoadOptions.None);
            return retVal;
        }

        /// <summary>
        /// Event handler for the context menu to search button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void CtxSearch_Click(object sender, RoutedEventArgs e)
        {
            this.searchModeParallel = false;
            this.BtnSearch_Click(sender, e);
        }

        /// <summary>
        /// Event handler for the context menu to parallel search button.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void CtxSearchParallel_Click(object sender, RoutedEventArgs e)
        {
            this.BtnSearch_Click(sender, e);
        }

        /// <summary>
        /// Determine whether updates can be checked for.
        /// </summary>
        private void DownloadUpdates()
        {
            if (this.preferenceFile != null && this.preferenceFile.Descendants("CheckForUpdates").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper())
            {
                DateTime lastUpdateCheckDate;

                if (this.preferenceFile.Descendants("LastUpdateCheckDate") != null && this.preferenceFile.Descendants("LastUpdateCheckDate").FirstOrDefault() != null
                    && DateTime.TryParse(this.preferenceFile.Descendants("LastUpdateCheckDate").FirstOrDefault().Value, out lastUpdateCheckDate))
                {
                    // Delete any previous version.
                    System.IO.File.Delete(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Searcher.exe.old"));

                    if (Common.ApplicationUpdateExists)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            this.Title = string.Format("{0} ({1})", this.Title, "Click 'About' for updates");
                        });
                    }

                    // Check for updates monthly. Why bother the user more frequently. Can look to make this configurable in the future.
                    if (lastUpdateCheckDate.AddMonths(1) < DateTime.Today)
                    {
                        Task.Run(async () =>
                        {
                            {
                                if (await this.NewReleaseExistsAsync())
                                {
                                    string newProgPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg");
                                    if (!System.IO.Directory.Exists(newProgPath))
                                    {
                                        string downloadedFile = await this.GetDownloadedUpdateFilenameAsync();
                                        System.IO.Compression.ZipFile.ExtractToDirectory(downloadedFile, newProgPath);
                                        System.IO.File.Delete(downloadedFile);
                                    }
                                }
                            }
                        });

                        // Save the current date as when the last check for updates was performed. Next check must be after 1 month atleast.
                        this.preferenceFile.Descendants("LastUpdateCheckDate").FirstOrDefault().Value = DateTime.Today.ToShortDateString();
                    }
                }
            }
        }

        /// <summary>
        /// Perform search if enter key is pressed.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void DtpEndDate_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Return)
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
        private void DtpStartDate_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Return)
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
        /// Sets up a directory to be exlcuded.
        /// </summary>
        /// <param name="sender">The sender object to get the file path.</param>
        /// <param name="e">The parameter is not used.</param>
        private void ExcludeDirectory_Click(object sender, RoutedEventArgs e)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);

            if (!string.IsNullOrEmpty(fullFilePath))
            {
                DirectoryExclude dirExcludeWindow = new DirectoryExclude(fullFilePath);
                dirExcludeWindow.PreferenceFileExists = this.preferenceFile != null;

                if (dirExcludeWindow.ShowDialog() == true)
                {
                    if (!this.directoriesToExclude.Contains(dirExcludeWindow.DirectoryToExclude) && Directory.Exists(dirExcludeWindow.DirectoryToExclude))
                    {
                        this.directoriesToExclude.Add(dirExcludeWindow.DirectoryToExclude);

                        if (dirExcludeWindow.IsExclusionPermanent)
                        {
                            this.preferenceFile.Descendants("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", dirExcludeWindow.DirectoryToExclude));
                        }
                    }
                }
            }
            else
            {
                this.ShowErrorPopup("Cannot exclude. If this is an archive, please exclude it as a file.");
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
                this.preferenceFile.Descendants("FilesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fullFilePath));
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
        /// Get the name of the updated application download file.
        /// </summary>
        /// <returns>The name of the updated application download file.</returns>
        private async Task<string> GetDownloadedUpdateFilenameAsync()
        {
            string retVal = string.Empty;

            retVal = await Task.Run<string>(() =>
            {
                string downloadUrl = "https://sourceforge.net/projects/searcher/files/latest/download";

                try
                {
                    string newFileName = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Searcher_new.zip");

                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(new Uri(downloadUrl), newFileName);
                    }

                    return newFileName;
                }
                catch
                {
                    return string.Empty;
                }
            });

            return retVal;
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
        /// Sets the title bar to display details of user running the application.
        /// </summary>
        /// <returns>String containing user name information.</returns>
        private string GetRunningUserInfo()
        {
            System.Reflection.AssemblyTitleAttribute assemblyTitleAttribute = (System.Reflection.AssemblyTitleAttribute)Attribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly(), typeof(System.Reflection.AssemblyTitleAttribute), false);
            string programName = assemblyTitleAttribute != null ? assemblyTitleAttribute.Title : "Unknown Assembly Name";
            string userName = string.Format(@"{0}\{1}", Environment.UserDomainName, Environment.UserName);
            System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent());

            if (principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator))
            {
                userName += " - Administrator";
            }

            return string.Format("{0} ({1})", programName, userName);
        }

        /// <summary>
        /// Show the files used in the search
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseButtonEventArgs object.</param>
        private void Grid_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (this.distinctFilesFound && this.filesToSearch != null && this.filesToSearch.Count > 0 || (this.filesToExclude.Count > 0 || this.directoriesToExclude.Count > 0))
            {
                SearchedFileList searchedFileList;
                List<string> filesToAlwaysExclude = new List<string>();
                List<string> directoriesToAlwaysExclude = new List<string>();

                if (this.preferenceFile != null)
                {
                    filesToAlwaysExclude = this.preferenceFile.Descendants("FilesToAlwaysExcludeFromSearch").Descendants("Value").Select(p => p.Value).ToList();
                    directoriesToAlwaysExclude = this.preferenceFile.Descendants("DirectoriesToAlwaysExcludeFromSearch").Descendants("Value").Select(p => p.Value).ToList();

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
                searchedFileList.ShowDialog();

                if (this.preferenceFile != null)
                {
                    if (searchedFileList.FilesToInclude.Count > 0)
                    {
                        this.filesToExclude.RemoveAll(tfe => searchedFileList.FilesToInclude.Any(fi => File.Exists(fi) && fi == tfe));
                        this.directoriesToExclude.RemoveAll(tde => searchedFileList.FilesToInclude.Any(fi => Directory.Exists(fi) && fi == tde));
                        filesToAlwaysExclude.RemoveAll(fae => searchedFileList.FilesToInclude.Any(fi => File.Exists(fi) && fi == fae));
                        directoriesToAlwaysExclude.RemoveAll(fae => searchedFileList.FilesToInclude.Any(fi => Directory.Exists(fi) && fi == fae));
                        SavePathsToAlwaysExclude(filesToAlwaysExclude, directoriesToAlwaysExclude);
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
            this.TxtBlkScaleValue.Visibility = System.Windows.Visibility.Hidden;
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

                if (!File.Exists(this.filePreferencesName))
                {
                    if (MessageBox.Show("No search preferences found. Do you want to save your search preferences?", "No search preferences", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.None) == MessageBoxResult.Yes)
                    {
                        this.preferenceFile = this.CreatePreferencesFile();
                    }
                    else
                    {
                        this.SetupAppWithoutPreferences();
                    }
                }
                else
                {
                    this.CheckPreferencesFile();
                    this.preferenceFile = XDocument.Load(this.filePreferencesName);
                    this.ChkMatchWholeWord.IsChecked = this.preferenceFile.Descendants("MatchWholeWord").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();
                    this.ChkMatchCase.IsChecked = this.preferenceFile.Descendants("MatchCase").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();
                    this.searchModeNormal = this.RbtnNormalSearch.IsChecked.Value == true;
                    this.searchModeRegex = this.RbtnRegexSearch.IsChecked.Value == true;
                    this.ChkSearchSubfolders.IsChecked = this.preferenceFile.Descendants("SearchSubfolders").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();
                    this.ChkHighlightResults.IsChecked = this.preferenceFile.Descendants("HighlightResults").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();
                    this.searchTypeAll = this.preferenceFile.Descendants("SearchContentMode").FirstOrDefault().Value.ToUpper() == "All".ToUpper();
                    this.showExecutionTime = this.preferenceFile.Descendants("ShowExecutionTime").FirstOrDefault().Value.ToUpper() == true.ToString().ToUpper();

                    this.SetSearchDate(this.DtpStartDate, "MinFileCreateSearchDate");
                    this.SetSearchDate(this.DtpEndDate, "MaxFileCreateSearchDate");
                    this.SetWindowDimensions();
                    this.SetApplicationCustomBackground();
                    this.SetResultHighlightColour();
                    this.SetSearchItemsSeparator();

                    if (!int.TryParse(this.preferenceFile.Descendants("MaxDropDownItems").FirstOrDefault().Value, out this.maxDropDownItems))
                    {
                        this.maxDropDownItems = 10;
                    }

                    XElement custEditorNamePath = this.preferenceFile.Descendants("CustomEditor").FirstOrDefault();
                    this.editorNamePath = (custEditorNamePath == null || string.IsNullOrWhiteSpace(custEditorNamePath.Value)) ? "notepad" : custEditorNamePath.Value;
                    this.TxtEditor.Text = Path.GetFileNameWithoutExtension(this.editorNamePath);

                    this.AddPreferencesToItems(this.CmbDirectory, "SearchDirectories");
                    this.AddPreferencesToItems(this.CmbFindWhat, "SearchContents");
                    this.AddPreferencesToItems(this.CmbFilters, "SearchFilters");
                    this.LoadPathsToAlwaysExlude();

                    this.CmbDirectory.Text = this.CmbDirectory.Items.Count > 0 ? this.CmbDirectory.Items[0].ToString() : string.Empty;
                    this.CmbFindWhat.Text = this.CmbFindWhat.Items.Count > 0 ? this.CmbFindWhat.Items[0].ToString() : string.Empty;
                    this.CmbFilters.Text = this.CmbFilters.Items.Count > 0 ? this.CmbFilters.Items[0].ToString() : string.Empty;
                }

                if (!string.IsNullOrEmpty(this.filePreferencesName) && File.Exists(this.filePreferencesName))
                {
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
        /// Set up the list of files that will be excluded always when searching.
        /// </summary>
        private void LoadPathsToAlwaysExlude()
        {
            this.preferenceFile.Descendants("FilesToAlwaysExcludeFromSearch").Descendants("Value").ToList().ForEach(p =>
            {
                this.AddToSearchFileExclusionList(p.Value);
            });

            this.preferenceFile.Descendants("DirectoriesToAlwaysExcludeFromSearch").Descendants("Value").ToList().ForEach(p =>
            {
                this.directoriesToExclude.Add(p.Value);
            });
        }

        /// <summary>
        /// Show right click menu options
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseButtonEventArgs object.</param>
        private void Link_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is Hyperlink)
            {
                ContextMenu mnu = new ContextMenu();
                MenuItem copyFullPathToClipboard = new MenuItem();
                copyFullPathToClipboard.Header = "Copy full path to clipboard";
                copyFullPathToClipboard.Tag = sender;
                copyFullPathToClipboard.Click += this.CopyFullPathToClipboard_Click;
                mnu.Items.Add(copyFullPathToClipboard);

                MenuItem copyFileNameToClipboard = new MenuItem();
                copyFileNameToClipboard.Header = "Copy file name to clipboard";
                copyFileNameToClipboard.Tag = sender;
                copyFileNameToClipboard.Click += this.CopyFileNameToClipboard_Click;
                mnu.Items.Add(copyFileNameToClipboard);

                MenuItem copyFileNameNoExtToClipboard = new MenuItem();
                copyFileNameNoExtToClipboard.Header = "Copy file name (without extension) to clipboard";
                copyFileNameNoExtToClipboard.Tag = sender;
                copyFileNameNoExtToClipboard.Click += this.CopyFileNameNoExtToClipboard_Click;
                mnu.Items.Add(copyFileNameNoExtToClipboard);
                mnu.Items.Add(new Separator());

                MenuItem openContainingDirectory = new MenuItem();
                openContainingDirectory.Header = "Open containing directory in explorer";
                openContainingDirectory.Tag = sender;
                openContainingDirectory.Click += this.OpenContainingDirectory_Click;
                mnu.Items.Add(openContainingDirectory);

                MenuItem saveAllResultsToFile = new MenuItem();
                saveAllResultsToFile.Header = "Save all results to file";
                saveAllResultsToFile.Click += this.SaveAllResultsToFile_Click;
                mnu.Items.Add(saveAllResultsToFile);

                MenuItem popoutResults = new MenuItem();
                popoutResults.Header = "Pop out results";
                popoutResults.Tag = sender;
                popoutResults.Click += this.PopoutResults_Click;
                mnu.Items.Add(popoutResults);

                mnu.Items.Add(new Separator());

                MenuItem excludeFileTemporarily = new MenuItem();
                excludeFileTemporarily.Header = "Exclude file from search temporarily";
                excludeFileTemporarily.Tag = sender;
                excludeFileTemporarily.Click += this.ExcludeFromSearchTemporarily_Click;
                mnu.Items.Add(excludeFileTemporarily);

                if (this.preferenceFile != null)
                {
                    MenuItem excludeFileAlways = new MenuItem();
                    excludeFileAlways.Header = "Always exclude file from search";
                    excludeFileAlways.Tag = sender;
                    excludeFileAlways.Click += this.ExcludeFromSearchAlways_Click;
                    mnu.Items.Add(excludeFileAlways);
                }

                MenuItem excludeDirectory = new MenuItem();
                excludeDirectory.Header = "Exclude directory from search";
                excludeDirectory.Tag = sender;
                excludeDirectory.Click += ExcludeDirectory_Click;
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
                Process.Start(this.editorNamePath, "\"" + e.Uri.LocalPath + "\"");
            }
            catch (Exception)
            {
                Process.Start("notepad", e.Uri.LocalPath);
                this.editorNamePath = "notepad";
            }
        }

        /// <summary>
        /// Determine if a new application version exists.
        /// </summary>
        /// <returns>Boolean indicating whether a new application version exists.</returns>
        private async Task<bool> NewReleaseExistsAsync()
        {
            bool retVal = false;

            if (System.IO.Directory.Exists(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg")))
            {
                retVal = true;
            }
            else
            {
                retVal = await Task.Run<string>(async () =>
                {
                    string path = @"https://sourceforge.net/projects/searcher/best_release.json";
                    using (HttpClient client = new HttpClient())
                    {
                        client.Timeout = new TimeSpan(0, 0, 0, 30);
                        return await client.GetStringAsync(new Uri(path));
                    }
                })
                .ContinueWith((s) =>
                {
                    if (s.Status == TaskStatus.RanToCompletion)
                    {
                        Newtonsoft.Json.Linq.JObject jsonSiteVersion = Newtonsoft.Json.Linq.JObject.Parse(s.Result);
                        string fileName = jsonSiteVersion["release"]["filename"].ToString();
                        string strSiteVersion = fileName.Replace("/Searcher_v", string.Empty).Replace(".zip", string.Empty);
                        Version appVersion = new Version(Common.VersionNumber);
                        Version siteVersion;

                        if (Version.TryParse(strSiteVersion, out siteVersion))
                        {
                            retVal = siteVersion > appVersion;
                        }
                    }

                    return retVal;
                });
            }

            return retVal;
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
        private void PerformSearch(string searchText, string searchPath, string filters)
        {
            List<string> filtersToUse = filters.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList();
            List<string> searchPaths = searchPath.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList();
            this.filesToSearch = new List<string>();
            this.filesSearched = new List<string>();
            this.distinctFilesFound = false;
            this.SetFileCounterProgressInformation(0, string.Format("Getting files to search."));
            this.regexOptions = this.matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;

            if (this.searchModeRegex)
            {
                this.regexOptions = this.regexOptions | (this.multilineRegex ? RegexOptions.Multiline : RegexOptions.Singleline);
            }

            List<string> termsToSearch = this.SetTermsToSearchText(searchText);
            this.otherExtensions = new SearchOtherExtensions { RegexOptions = this.regexOptions, IsRegexSearch = this.searchModeRegex };
            List<Task<List<string>>> pathsListTask = new List<Task<List<string>>>(searchPaths.Count());

            searchPaths.ForEach(path =>
            {
                filtersToUse.ForEach(filter =>
                {
                    if (!this.cancellationTokenSource.IsCancellationRequested)
                    {
                        pathsListTask.Add(Task.Run<List<string>>(() => { return GetFilesToSearch(path, filter); }));
                    }
                });
            });

            try
            {
                Task<List<string>[]> result = Task.WhenAll(pathsListTask);
                result.Wait();
                pathsListTask.ForEach(task => { this.filesToSearch.AddRange(task.Result); });
                RemoveExclusionPaths();
                this.SetFileCounterProgressInformation(0, string.Format("Files found: {0}", this.filesToSearch.Count));
                this.filesToSearch = this.filesToSearch.Distinct().OrderBy(f => f).ToList();      // Remove duplicates that could be added via path filters that cover the same item mulitple times.
                this.distinctFilesFound = true;
                this.SetProgressMaxValue(this.filesToSearch.Count);

                if (this.searchModeParallel)
                {
                    this.SearchParallel(this.filesToSearch, termsToSearch);
                }
                else
                {
                    this.searchModeParallel = true;
                    this.Search(this.filesToSearch, termsToSearch);
                }
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
            newWnd.Title = string.Format("{0} - {1}", "Results", this.CmbFindWhat.Text.Substring(0, Math.Min(this.CmbFindWhat.Text.Length, 50)));
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
        private void RemoveExclusionPaths()
        {
            this.filesToSearch.RemoveAll(f => this.filesToExclude.Any(excludeFile => excludeFile.ToUpper() == f.ToUpper()));

            foreach (string dirToExclude in this.directoriesToExclude)
            {
                string exclDir = dirToExclude.ToUpper();
                filesToSearch.RemoveAll(f => 
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
            this.SetFileCounterProgressInformation(0, string.Format("Getting files to search."));
            this.TxtResults.Document = new FlowDocument(this.richTextboxParagraph);
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
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            if (!string.IsNullOrEmpty(this.CmbDirectory.Text))
            {
                sfd.InitialDirectory = this.CmbDirectory.Text.Split(new string[] { this.separatorCharacter }, StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                sfd.Filter = "txt files (*.txt)|*.txt";
            }

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                System.IO.File.WriteAllText(sfd.FileName, new TextRange(this.TxtResults.Document.ContentStart, this.TxtResults.Document.ContentEnd).Text);
            }
        }

        /// <summary>
        /// Search for matches.
        /// </summary>
        /// <param name="fileNamePaths">The filenames with path to search for.</param>
        /// <param name="termsToSearch">The terms to search for.</param>
        private void Search(IEnumerable<string> fileNamePaths, IEnumerable<string> termsToSearch)
        {
            int fileCounter = 0;
            Matcher matcher = new Matcher(this.matchWholeWord, this.searchModeRegex, this.multilineRegex, this.searchTypeAll, this.cancellationTokenSource, this.regexOptions);

            foreach (string fileName in fileNamePaths)
            {
                fileCounter++;
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    try
                    {
                        this.SearchAndDisplayResult(fileName, matcher, termsToSearch);
                    }
                    catch (Exception ex)
                    {
                        this.SetSearchError(string.Format("Error reading file: {0}. {1}", fileName, ex.Message));
                    }

                    this.SetFileCounterProgressInformation(fileCounter, string.Format("Processed Files {0} of {1} ({2} %)", fileCounter + 1, fileNamePaths.Count(), ((fileCounter + 1) * 100) / fileNamePaths.Count()));

                    if (this.cancellationTokenSource.IsCancellationRequested)
                    {
                        break;
                    }
                }
            }

            this.SetSearchCompletedDetails(fileCounter, fileNamePaths.Count(), this.matchesFound, this.filesWithMatch);
        }

        /// <summary>
        /// Get the search result and display in text box.
        /// </summary>
        /// <param name="fileName">The file name to search.</param>
        /// <param name="matcher">The matcher object with configuration to determine how to search.</param>
        /// <param name="termsToSearch">The terms to search.</param>
        private void SearchAndDisplayResult(string fileName, Matcher matcher, IEnumerable<string> termsToSearch)
        {
            this.SetProgressInformation(string.Format("Searching: {0}", fileName));
            List<MatchedLine> matchedLines = matcher.GetMatch(fileName, termsToSearch);

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
                        this.AddResult(groupedMatch.MatchLine);
                    }
                }
                else
                {
                    matchedLines = matchedLines.OrderBy(ml => ml.LineNumber).ThenBy(ml => ml.StartIndex).ToList();
                    matchedLines.ForEach(ml => ml.FileName = string.Empty);
                    matchedLines[0].FileName = fileName;
                    this.AddResult(matchedLines);
                }
            }
        }

        /// <summary>
        /// Search for matches using parallel processing.
        /// </summary>
        /// <param name="fileNamePaths">The filenames with path to search for.</param>
        /// <param name="termsToSearch">The terms to search for.</param>
        private void SearchParallel(IEnumerable<string> fileNamePaths, IEnumerable<string> termsToSearch)
        {
            int counter = 0;
            ParallelOptions parallelOptions = new ParallelOptions();
            parallelOptions.CancellationToken = this.cancellationTokenSource.Token;
            Matcher matcher = new Matcher(this.matchWholeWord, this.searchModeRegex, this.multilineRegex, this.searchTypeAll, this.cancellationTokenSource, this.regexOptions);
            parallelOptions.MaxDegreeOfParallelism = 4;     // Not dependant on CPU cores. The value 4 seems good enough for the algorithm.
            Parallel.ForEach(
                fileNamePaths,
                parallelOptions,
                (string fileName, ParallelLoopState state) =>
                {
                    if (!string.IsNullOrWhiteSpace(fileName))
                    {
                        try
                        {
                            this.SearchAndDisplayResult(fileName, matcher, termsToSearch);
                        }
                        catch (Exception ex)
                        {
                            this.SetSearchError(ex.Message);
                        }

                        counter += 1;
                        this.SetFileCounterProgressInformation(counter, string.Format("Processed Files {0} of {1} ({2} %)", counter, fileNamePaths.Count(), (int)(counter * 100) / fileNamePaths.Count()));

                        if (this.cancellationTokenSource.IsCancellationRequested)
                        {
                            state.Stop();
                        }
                    }
                });

            counter = this.cancellationTokenSource.IsCancellationRequested ? counter : fileNamePaths.Count();     // If task cancelled, show the actual file processed count, else all files are processed.
            this.SetSearchCompletedDetails(counter, fileNamePaths.Count(), this.matchesFound, this.filesWithMatch);
        }

        /// <summary>
        /// Display time taken to search.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The ElapsedEventArgs object.</param>
        private void SearchTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (!this.cancellationTokenSource.IsCancellationRequested)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        this.TblkProgressTime.Text = string.Format("{0:hh\\:mm\\:ss}", DateTime.Now.Subtract(this.searchStartTime));
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
            string backGroundColourValue = this.preferenceFile.Descendants("BackGroundColour").FirstOrDefault().Value;
            backGroundColourValue = backGroundColourValue.StartsWith("#") ? backGroundColourValue : "#" + backGroundColourValue;

            try
            {
                this.applicationBackColour = (SolidColorBrush)new System.Windows.Media.BrushConverter().ConvertFromString(backGroundColourValue);
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
                this.preferenceFile.Descendants("HighlightResultsColour").FirstOrDefault().Value = "#FFFFFF";
            }
        }

        /// <summary>
        /// Dummy search for testing.
        /// </summary>
        private void SetDummySearch()
        {
            this.CmbDirectory.Text = @"D:\Code\SvnCode\SQL Server\Stored Procedures";
            this.CmbFindWhat.Text = @"t_new_material; insert";
            this.CmbFilters.Text = @"*.sql";
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
                this.Dispatcher.Invoke(() =>
                {
                    if (value >= 0)
                    {
                        this.PgBarSearch.Value = value;
                    }

                    this.TblkProgress.Text = text;
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
        /// Set a jump list for the window.
        /// </summary>
        private void SetJumpList()
        {
            JumpTask jmpPreferences = new JumpTask();
            jmpPreferences.ApplicationPath = this.filePreferencesName;
            jmpPreferences.IconResourcePath = @"C:\Windows\notepad.exe";
            jmpPreferences.Title = "Preferences";
            jmpPreferences.Description = "Open preferences file";
            jmpPreferences.CustomCategory = "Settings";

            JumpList jmpList = JumpList.GetJumpList(Application.Current);
            jmpList.JumpItems.Add(jmpPreferences);
            JumpList.AddToRecentCategory(jmpPreferences);
            jmpList.Apply();
        }

        /// <summary>
        /// Set the paths that will be always excluded.
        /// </summary>
        /// <param name="filesToAlwaysExclude">The list of file paths.</param>
        /// <param name="directoriesToAlwaysExclude">The list of directory paths.</param>
        private void SavePathsToAlwaysExclude(List<string> filesToAlwaysExclude, List<string> directoriesToAlwaysExclude)
        {
            this.preferenceFile.Descendants("FilesToAlwaysExcludeFromSearch").FirstOrDefault().RemoveAll();
            this.preferenceFile.Descendants("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().RemoveAll();

            filesToAlwaysExclude.ForEach(fae =>
            {
                if (File.Exists(fae))
                {
                    this.preferenceFile.Descendants("FilesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fae));
                }
            });

            directoriesToAlwaysExclude.ForEach(fae =>
            {
                if (Directory.Exists(fae))
                {
                    this.preferenceFile.Descendants("DirectoriesToAlwaysExcludeFromSearch").FirstOrDefault().Add(new XElement("Value", fae));
                }
            });
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
        /// Set the highlight colour to be used for matching results.
        /// </summary>
        private void SetResultHighlightColour()
        {
            string highlightResultColourValue = this.preferenceFile.Descendants("HighlightResultsColour").FirstOrDefault().Value;
            highlightResultColourValue = highlightResultColourValue.StartsWith("#") ? highlightResultColourValue : "#" + highlightResultColourValue;        // Prepend with "#" if required.

            try
            {
                SolidColorBrush newColour = (SolidColorBrush)new System.Windows.Media.BrushConverter().ConvertFromString(highlightResultColourValue);
                this.highlightResultBackColour = newColour;
            }
            catch (FormatException)
            {
                this.highlightResultBackColour = Brushes.PeachPuff;
            }

            this.preferenceFile.Descendants("HighlightResultsColour").FirstOrDefault().Value = this.highlightResultBackColour.ToString();
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
                this.SetFileCounterProgressInformation(filesProcessedCount, string.Format("Processed Files {0} of {1} ({2} %)", filesProcessedCount, filesProcessedCountOriginal, ((int)filesProcessedCount / filesProcessedCountOriginal) * 100));
            }

            this.SetProgressInformation(string.Format("Found {0} match(es) in {1} file(s)", matchesFound, filesWithMatches));
        }

        /// <summary>
        /// Set the dates used by the application.
        /// </summary>
        /// <param name="datePicker">The date picker control.</param>
        /// <param name="preferenceElement">The preference element to use.</param>
        private void SetSearchDate(DatePicker datePicker, string preferenceElement)
        {
            if (this.preferenceFile != null && this.preferenceFile.Descendants(preferenceElement) != null)
            {
                string strDateToSet = this.preferenceFile.Descendants(preferenceElement).FirstOrDefault().Value;
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
                    this.BrdrErrors.Visibility = string.IsNullOrEmpty(this.TxtErrors.Text) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
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
            if (this.preferenceFile.Descendants("SeparatorCharacter") != null && this.preferenceFile.Descendants("SeparatorCharacter").FirstOrDefault() != null && !string.IsNullOrWhiteSpace(this.preferenceFile.Descendants("SeparatorCharacter").FirstOrDefault().Value))
            {
                this.separatorCharacter = this.preferenceFile.Descendants("SeparatorCharacter").FirstOrDefault().Value.Substring(0, 1);

                if (this.separatorCharacter != ";")
                {
                    string separatorTooltipMessage = string.Format("Separate additional entries by using the '{0}' character.", this.separatorCharacter);
                    this.CmbDirectory.ToolTip = separatorTooltipMessage;
                    this.CmbFindWhat.ToolTip = separatorTooltipMessage;
                    this.CmbFilters.ToolTip = separatorTooltipMessage;
                }
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
            this.searchTypeAll = this.CmbFindWhat.Text.Contains(this.separatorCharacter) && ((ComboBoxItem)this.CmbSearchType.SelectedValue).Content.ToString().ToUpper() == "All".ToUpper();
            this.multilineRegex = this.ChkRegexMultiline.IsChecked.Value == true;
            this.minSearchDate = this.DtpStartDate.SelectedDate.HasValue ? this.DtpStartDate.SelectedDate.Value : DateTime.MinValue;
            this.maxSearchDate = this.DtpEndDate.SelectedDate.HasValue ? this.DtpEndDate.SelectedDate.Value.Add(new TimeSpan(23, 59, 59)) : DateTime.MaxValue;   // Add time else uses 00:00:00
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
        }

        /// <summary>
        /// Set up the height and width of the search window.
        /// </summary>
        private void SetWindowDimensions()
        {
            string strWindowHeight = this.preferenceFile.Descendants("WindowHeight").FirstOrDefault().Value;
            string strWindowWidth = this.preferenceFile.Descendants("WindowWidth").FirstOrDefault().Value;
            string strPopupWindowHeight = this.preferenceFile.Descendants("PopupWindowHeight").FirstOrDefault().Value;
            string strPopupWindowWidth = this.preferenceFile.Descendants("PopupWindowWidth").FirstOrDefault().Value;
            string strPopupWindowTimeoutSeconds = this.preferenceFile.Descendants("PopupWindowTimeoutSeconds").FirstOrDefault().Value;

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
            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Exception occurred");
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
                    this.contentPopup = new ContentPopup(file, lineNo);
                    this.contentPopup.Width = this.popupWindowWidth;
                    this.contentPopup.Height = this.popupWindowHeight;
                    this.contentPopup.WindowCloseTimeoutSeconds = (this.popupWindowTimeoutSeconds < 2 || this.popupWindowTimeoutSeconds > 20)
                        ? 4
                        : this.popupWindowTimeoutSeconds;
                    this.contentPopup.Show();
                    this.contentPopup.Closed += this.ContentPopup_Closed;
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
            if (System.Windows.Input.Keyboard.IsKeyDown(Key.LeftCtrl) || System.Windows.Input.Keyboard.IsKeyDown(Key.RightCtrl))
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
                    this.TxtBlkScaleValue.Visibility = System.Windows.Visibility.Visible;

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
        /// Handle window closing event to save preferences.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The CancelEventArgs object.</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.preferenceFile != null)
            {
                this.preferenceFile.Descendants("MatchWholeWord").FirstOrDefault().Value = this.ChkMatchWholeWord.IsChecked.Value.ToString();
                this.preferenceFile.Descendants("MatchCase").FirstOrDefault().Value = this.ChkMatchCase.IsChecked.Value.ToString();
                this.preferenceFile.Descendants("SearchSubfolders").FirstOrDefault().Value = this.ChkSearchSubfolders.IsChecked.Value.ToString();
                this.preferenceFile.Descendants("HighlightResults").FirstOrDefault().Value = this.ChkHighlightResults.IsChecked.Value.ToString();
                this.preferenceFile.Descendants("MinFileCreateSearchDate").FirstOrDefault().Value = (this.DtpStartDate.SelectedDate != null && this.DtpStartDate.SelectedDate.HasValue) ? this.DtpStartDate.SelectedDate.Value.ToString() : string.Empty;
                this.preferenceFile.Descendants("MaxFileCreateSearchDate").FirstOrDefault().Value = (this.DtpEndDate.SelectedDate != null && this.DtpEndDate.SelectedDate.HasValue) ? this.DtpEndDate.SelectedDate.Value.ToString() : string.Empty;
                this.preferenceFile.Descendants("WindowHeight").FirstOrDefault().Value = this.Height.ToString();
                this.preferenceFile.Descendants("WindowWidth").FirstOrDefault().Value = this.Width.ToString();
                this.preferenceFile.Descendants("WindowWidth").FirstOrDefault().Value = this.Width.ToString();
                this.preferenceFile.Descendants("PopupWindowHeight").FirstOrDefault().Value = this.popupWindowHeight.ToString();
                this.preferenceFile.Descendants("PopupWindowWidth").FirstOrDefault().Value = this.popupWindowWidth.ToString();
                this.preferenceFile.Descendants("WindowWidth").FirstOrDefault().Value = this.Width.ToString();
                
                this.preferenceFile.Descendants("SearchContentMode").FirstOrDefault().Value = this.searchTypeAll ? "All" : "Any";
                this.AddItemsToPreferences(this.CmbDirectory, "SearchDirectories");
                this.AddItemsToPreferences(this.CmbFindWhat, "SearchContents");
                this.AddItemsToPreferences(this.CmbFilters, "SearchFilters");

                try
                {
                    if (this.cancellationTokenSource != null && !this.cancellationTokenSource.IsCancellationRequested)
                    {
                        this.cancellationTokenSource.Cancel();
                    }

                    this.preferenceFile.Save(this.filePreferencesName);
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
            if ((System.Windows.Input.Keyboard.IsKeyDown(Key.LeftCtrl) || System.Windows.Input.Keyboard.IsKeyDown(Key.RightCtrl)) && e.Key == Key.F)
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
        }

        #endregion Private Methods
    }
}
