// <copyright file="Common.cs" company="dennjose">
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
    using System.IO;
    using System.Linq;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using SearcherLibrary;

    /// <summary>
    /// Class to perform common tasks.
    /// </summary>
    public class Common
    {
        /// <summary>
        /// Gets a value indicating whether an application update has been downloaded and is pending to be applied.
        /// </summary>
        public static bool ApplicationUpdateExists
        {
            get
            {
                return System.IO.Directory.Exists(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg"));
            }
        }

        /// <summary>
        /// Gets the version number for the application.
        /// </summary>
        public static string VersionNumber
        {
            get
            {
                System.Version programVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string version = string.Format("{0}.{1}.{2}", programVersion.Major, programVersion.Minor, programVersion.Build);
                return version;
            }
        }

        /// <summary>
        /// Get the file path from the hyperlink object.
        /// </summary>
        /// <param name="sender">The Hyperlink object.</param>
        /// <returns>Full file path contained in the Hyperlink's NavigateUri property.</returns>
        public static string GetLinkUriDetails(object sender)
        {
            string retVal = string.Empty;

            if (sender is MenuItem)
            {
                string fileNamePath = string.Empty;

                if (((MenuItem)sender).Tag is Hyperlink)
                {
                    Hyperlink selectedPathLink = (Hyperlink)((MenuItem)sender).Tag;

                    if (selectedPathLink.NavigateUri != null && selectedPathLink.NavigateUri.ToString() != string.Empty && !string.IsNullOrWhiteSpace(selectedPathLink.NavigateUri.LocalPath))
                    {
                        fileNamePath = selectedPathLink.NavigateUri.LocalPath;
                    }
                }
                else if (((MenuItem)sender).Tag is ListBox)
                {
                    if (((ListBox)((MenuItem)sender).Tag).SelectedItems != null && ((ListBox)((MenuItem)sender).Tag).SelectedItems.Count == 1)
                    {
                        fileNamePath = ((ListBox)((MenuItem)sender).Tag).SelectedItems[0].ToString();
                    }
                }

                if (!string.IsNullOrWhiteSpace(fileNamePath) && File.Exists(fileNamePath))
                {
                    retVal = Path.GetFullPath(fileNamePath);
                }
            }

            return retVal;
        }

        /// <summary>
        /// Determine if the file is a file that can be read (i.e. not .PDF, .DOCX etc.)
        /// </summary>
        /// <param name="fileName">The full name of the file.</param>
        /// <returns>Boolean indicating whether the file can be read in the default editor.</returns>
        public static bool IsAsciiSearch(string fileName)
        {
            bool retVal = false;
            string fileExtension = Path.GetExtension(fileName).ToUpper();
            retVal = !Enum.GetNames(typeof(OtherExtensions)).Any(s => fileExtension.Contains(s.ToUpper()));
            return retVal;
        }
        
        /// <summary>
        /// Context menu to open windows explorer to the selected file.
        /// </summary>
        /// <param name="sender">The sender object that contains the file name.</param>
        public static void OpenDirectoryForFile(object sender)
        {
            string fullFilePath = Common.GetLinkUriDetails(sender);

            if (!string.IsNullOrWhiteSpace(fullFilePath) && File.Exists(fullFilePath))
            {
                Process explorerWindowProcess = new Process();
                explorerWindowProcess.StartInfo.FileName = "explorer.exe";
                explorerWindowProcess.StartInfo.Arguments = "/select,\"" + @fullFilePath + "\"";
                explorerWindowProcess.Start();
            }
        }
    }
}
