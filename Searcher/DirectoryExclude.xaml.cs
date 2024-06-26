﻿// <copyright file="DirectoryExclude.xaml.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for DirectoryExclude window.
    /// </summary>
    public partial class DirectoryExclude : Window
    {
        /// <summary>
        /// Initialises a new instance of the <see cref="DirectoryExclude"/> class.
        /// </summary>
        public DirectoryExclude()
        {
            this.InitializeComponent();
            this.SetContentBasedOnLanguage();
            this.DirectoryToExclude = string.Empty;
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="DirectoryExclude"/> class.
        /// </summary>
        /// <param name="childPath">The child path whose hierarchy is to be displayed.</param>
        public DirectoryExclude(string childPath)
            : this()
        {
            List<string> dirPaths = [];
            string tempPath = childPath;

            while (Directory.GetParent(tempPath)?.FullName != Directory.GetDirectoryRoot(childPath))
            {
                tempPath = Directory.GetParent(tempPath)?.FullName ?? string.Empty;
                dirPaths.Add(tempPath);
            }

            string nodeToFind = string.Empty;

            foreach (string dirPath in dirPaths.OrderBy(p => p))
            {
                TreeViewItem item = new()
                {
                    Header = dirPath,
                    IsExpanded = true,
                    IsSelected = true
                };

                if (this.TvDirectoryStructure.Items == null || this.TvDirectoryStructure.Items.Count == 0)
                {
                    this.TvDirectoryStructure.Items?.Add(item);
                }
                else
                {
                    TreeViewItem? nodeToUse = GetItemWithText((TreeViewItem)this.TvDirectoryStructure.Items[0], nodeToFind);
                    nodeToUse?.Items.Add(item);
                }

                nodeToFind = dirPath;
            }
        }

        /// <summary>
        /// Gets the directory that is to be excluded.
        /// </summary>
        public string DirectoryToExclude
        {
            get;

            private set;
        }

        /// <summary>
        /// Gets a value indicating whether the directory is permanently excluded.
        /// </summary>
        public bool IsExclusionPermanent
        {
            get;

            private set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether a preference file exists (to determine if the directory can be excluded permanently).
        /// </summary>
        public bool PreferenceFileExists
        {
            get;
            set;
        }

        /// <summary>
        /// Sets the directory to be excluded.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            string? selectedDirectory = ((TreeViewItem)this.TvDirectoryStructure.SelectedItem).Header.ToString();

            if (!string.IsNullOrWhiteSpace(selectedDirectory))
            {
                this.DirectoryToExclude = selectedDirectory;
                this.IsExclusionPermanent = this.RbtnPermanent.IsChecked.GetValueOrDefault();
            }

            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Get the TreeViewItem object containing the header text specified.
        /// </summary>
        /// <param name="treeViewItem">The TreeViewItem parent object.</param>
        /// <param name="textToFind">The text to find in the TreeViewItem object or its child nodes.</param>
        /// <returns>The TreeViewItem object containing the header text specified.</returns>
        private static TreeViewItem? GetItemWithText(TreeViewItem treeViewItem, string textToFind)
        {
            TreeViewItem? retVal = null;

            if (treeViewItem.Header.ToString() == textToFind)
            {
                retVal = treeViewItem;
            }
            else
            {
                foreach (TreeViewItem item in treeViewItem.Items)
                {
                    retVal = GetItemWithText(item, textToFind);

                    if (retVal != null)
                    {
                        break;
                    }
                }
            }

            return retVal;
        }

        /// <summary>
        /// Set readable content based on selected language.
        /// </summary>
        private void SetContentBasedOnLanguage()
        {
            this.Title = Application.Current.Resources["SelectDirToExclude"].ToString();
            this.RbtnTemporary.Content = Application.Current.Resources["Temporary"].ToString();
            this.RbtnPermanent.Content = Application.Current.Resources["Permanent"].ToString();
            this.TblkExclusionType.Text = Application.Current.Resources["ExclusionType"].ToString();
            this.BtnCancel.Content = Application.Current.Resources["Cancel"].ToString();
            this.BtnOk.Content = Application.Current.Resources["OK"].ToString();
        }

        /// <summary>
        /// Set up window for display.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.GrdExclusionType.Visibility = this.PreferenceFileExists ? Visibility.Visible : Visibility.Collapsed;
            this.TvDirectoryStructure.Focus();
        }
    }
}
