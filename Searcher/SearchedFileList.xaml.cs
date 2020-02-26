// <copyright file="SearchedFileList.xaml.cs" company="dennjose">
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
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for SearchedFileList class.
    /// </summary>
    public partial class SearchedFileList : Window
    {
        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        public SearchedFileList()
        {
            this.InitializeComponent();
            this.FilesToInclude = new List<string>();
            this.SetContentBasedOnLanguage();
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        /// <param name="fileNamePaths">The list of file name paths.</param>
        public SearchedFileList(IEnumerable<string> fileNamePaths)
            : this()
        {
            if (fileNamePaths != null && fileNamePaths.Count() > 0)
            {
                fileNamePaths.OrderBy(f => f).ToList().ForEach(f =>
                {
                    this.LstFileList.Items.Add(f);
                });
            }
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        /// <param name="fileNamePaths">The list of file name paths.</param>
        /// <param name="excludedFileNames">The list of file names excluded from search.</param>
        public SearchedFileList(IEnumerable<string> fileNamePaths, IEnumerable<string> excludedFileNames)
            : this(fileNamePaths)
        {
            if (excludedFileNames.Count() > 0)
            {
                this.GrdFilesExcluded.Visibility = System.Windows.Visibility.Visible;
                excludedFileNames.OrderBy(f => f).ToList().ForEach(f =>
                {
                    this.LstFilesExcluded.Items.Add(f);
                });
            }
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        /// <param name="fileNamePaths">The list of file name paths.</param>
        /// <param name="excludedFileNames">The list of file names excluded from search.</param>
        /// <param name="alwaysExcludedFileNames">The list of file names excluded from search via configuration.</param>
        public SearchedFileList(IEnumerable<string> fileNamePaths, IEnumerable<string> excludedFileNames, IEnumerable<string> alwaysExcludedFileNames)
            : this(fileNamePaths, excludedFileNames)
        {
            if (alwaysExcludedFileNames.Count() > 0)
            {
                this.GrdFilesExcludedAlways.Visibility = System.Windows.Visibility.Visible;

                alwaysExcludedFileNames.OrderBy(f => f).ToList().ForEach(f =>
                {
                    this.LstFilesAlwaysExcluded.Items.Add(f);
                });
            }
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets the list of file to be removed from the exclusion list.
        /// </summary>
        public List<string> FilesToInclude
        {
            get;
            set;
        }

        #endregion Public Properties

        #region Private Methods

        /// <summary>
        /// Copy the list of file names searched to clipboard.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void BtnCopyList_Click(object sender, RoutedEventArgs e)
        {
            if (this.LstFileList.Items != null)
            {
                List<string> filesToCopy = new List<string>();

                if (this.LstFileList.SelectedItems != null && this.LstFileList.SelectedItems.Count > 0)
                {
                    foreach (string item in this.LstFileList.SelectedItems)
                    {
                        filesToCopy.Add(item);
                    }
                }
                else
                {
                    foreach (string item in this.LstFileList.Items)
                    {
                        filesToCopy.Add(item);
                    }
                }

                Clipboard.SetData(System.Windows.DataFormats.Text, string.Join(Environment.NewLine, filesToCopy));
                this.Close();
            }
        }

        /// <summary>
        /// Context menu for item.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The MouseButtonEventArgs object.</param>
        private void LstFileList_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (LstFileList.Items != null && LstFileList.Items.Count > 0 && LstFileList.SelectedItem != null && !string.IsNullOrWhiteSpace(LstFileList.SelectedItem.ToString()) && LstFileList.SelectedItems.Count == 1)
            {
                System.Windows.Controls.ContextMenu mnu = new System.Windows.Controls.ContextMenu();
                MenuItem openContainingDirectoryInExplorer = new MenuItem();
                openContainingDirectoryInExplorer.Header = "Open containing directory in explorer";
                openContainingDirectoryInExplorer.Tag = sender;
                openContainingDirectoryInExplorer.Click += this.OpenContainingDirectoryInExplorer_Click;

                mnu.Items.Add(openContainingDirectoryInExplorer);
                LstFileList.ContextMenu = mnu;
            }
        }

        /// <summary>
        /// Get items to bring back into search list from the list of always excluded files.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void LstFilesAlwaysExcluded_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && sender is System.Windows.Controls.ListBox)
            {
                this.ReincludeItemsFromList((System.Windows.Controls.ListBox)sender);
            }
        }

        /// <summary>
        /// Get items to bring back into search list from the list of temporarily excluded files.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void LstFilesExcluded_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && sender is System.Windows.Controls.ListBox)
            {
                this.ReincludeItemsFromList((System.Windows.Controls.ListBox)sender);
            }
        }

        /// <summary>
        /// Open the file location in windows explorer.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The RoutedEventArgs object.</param>
        private void OpenContainingDirectoryInExplorer_Click(object sender, RoutedEventArgs e)
        {
            Common.OpenDirectoryForFile(sender);
        }

        /// <summary>
        /// Get items to remove from the files excluded from search.
        /// </summary>
        /// <param name="listbox">The list box control.</param>
        private void ReincludeItemsFromList(System.Windows.Controls.ListBox listbox)
        {
            List<string> itemsToInclude = listbox.SelectedItems.Cast<string>().ToList();
            this.FilesToInclude.AddRange(itemsToInclude.Where(i => (File.Exists(i) || Directory.Exists(i)) && !this.FilesToInclude.Any(fi => fi == i)));

            itemsToInclude.ForEach(item =>
            {
                listbox.Items.Remove(item);
            });
        }

        /// <summary>
        /// Close the window.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The KeyEventArgs object.</param>
        private void SearchedFileList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        /// <summary>
        /// Set readable content based on selected language.
        /// </summary>
        private void SetContentBasedOnLanguage()
        {
            this.Title = Application.Current.Resources["SearchedFileList"].ToString();
            this.BtnCopyList.Content = Application.Current.Resources["Directory"].ToString();
            this.TblkAlwaysExcluded.Text = Application.Current.Resources["TemporarilyExcluded"].ToString();
            this.TblkTemporarilyExcluded.Text = Application.Current.Resources["AlwaysExcluded"].ToString();
        }

        #endregion Private Methods
    }
}
