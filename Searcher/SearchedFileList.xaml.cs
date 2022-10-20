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
        #region Private Fields

        /// <summary>
        /// Private filed to retain window height between instances.
        /// </summary>
        private static double windowHeight;

        /// <summary>
        /// Private filed to retain window left position between instances.
        /// </summary>
        private static double windowLeft;

        /// <summary>
        /// Private filed to retain window top position between instances.
        /// </summary>
        private static double windowTop;

        /// <summary>
        /// Private filed to retain window width between instances.
        /// </summary>
        private static double windowWidth;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        public SearchedFileList()
        {
            this.InitializeComponent();
            this.FilesToInclude = new List<string>();
            this.SetContentBasedOnLanguage();

            if (windowWidth > 0 && windowHeight > 0 && windowLeft > 0 && windowTop > 0)
            {
                this.Height = windowHeight;
                this.Width = windowWidth;
                this.Left = windowLeft;
                this.Top = windowTop;
            }
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
        /// <param name="excludedContent">The list of file names excluded from search.</param>
        public SearchedFileList(IEnumerable<string> fileNamePaths, IEnumerable<string> excludedContent)
            : this(fileNamePaths)
        {
            if (excludedContent.Count() > 0)
            {
                this.ExpandTemporarilyExcluded.IsExpanded = true;

                excludedContent.OrderBy(f => f).ToList().ForEach(f =>
                {
                    this.LstTemporarilyExcluded.Items.Add(f);
                });
            }
            else
            {
                this.ExpandTemporarilyExcluded.IsExpanded = false;
            }
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="SearchedFileList"/> class.
        /// </summary>
        /// <param name="fileNamePaths">The list of file name paths.</param>
        /// <param name="temporarilyExcludedContent">The list of file names excluded from search.</param>
        /// <param name="alwaysExcludedContent">The list of file names excluded from search via configuration.</param>
        public SearchedFileList(IEnumerable<string> fileNamePaths, IEnumerable<string> temporarilyExcludedContent, IEnumerable<string> alwaysExcludedContent)
            : this(fileNamePaths, temporarilyExcludedContent)
        {
            if (alwaysExcludedContent.Count() > 0)
            {
                this.ExpandAlwaysExcluded.IsExpanded = true;

                alwaysExcludedContent.OrderBy(f => f).ToList().ForEach(f =>
                {
                    this.LstAlwaysExcluded.Items.Add(f);
                });
            }
            else
            {
                this.ExpandAlwaysExcluded.IsExpanded = false;
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
            this.BtnCopyList.Content = Application.Current.Resources["CopyAll"].ToString();
            this.ExpandAlwaysExcluded.Header = Application.Current.Resources["AlwaysExcluded"].ToString();
            this.ExpandTemporarilyExcluded.Header = Application.Current.Resources["TemporarilyExcluded"].ToString();
            this.LstTemporarilyExcluded.ToolTip = Application.Current.Resources["DeleteToRemoveItem"].ToString();
            this.LstAlwaysExcluded.ToolTip = Application.Current.Resources["DeleteToRemoveItem"].ToString();
        }

        /// <summary>
        /// Remember window dimensions for reopening next time.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void Window_Closed(object sender, EventArgs e)
        {
            windowHeight = this.Height;
            windowLeft = this.Left;
            windowTop = this.Top;
            windowWidth = this.Width;
        }

        #endregion Private Methods  
    }
}
