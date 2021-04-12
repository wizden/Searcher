// <copyright file="SearchOtherExtensions.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary
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
    using System.IO.Compression;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Wordprocessing;
    using FileExtensions;
    using SharpCompress.Common;
    using SharpCompress.Readers;

    /// <summary>
    /// Class to search files with NON-ASCII extensions.
    /// </summary>
    public class SearchOtherExtensions
    {
        #region Private Fields

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        internal const int IndexBoundary = 50;

        /// <summary>
        /// Private variable to hold the name of the temporary extraction directory.
        /// </summary>
        internal const string TempExtractDirectoryName = "→_extract_←";

        /// <summary>
        /// List of characters not allowed for the file system.
        /// </summary>
        private static List<char> disallowedCharactersByOperatingSystem;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether the search mode uses regex.
        /// </summary>
        public bool IsRegexSearch { get; set; }

        /// <summary>
        /// Gets or sets the regex options to use when searching.
        /// </summary>
        public RegexOptions RegexOptions { get; set; }

        #endregion Public Properties

        #region Private Properties

        /// <summary>
        /// Gets the list of characters not allowed by the operating system.
        /// </summary>
        internal static List<char> DisallowedCharactersByOperatingSystem
        {
            get
            {
                if (disallowedCharactersByOperatingSystem == null)
                {
                    disallowedCharactersByOperatingSystem = new List<char>();

                    if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        // Based on: https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file. Excluding "\" and "/" as these get used for IEntry paths.
                        disallowedCharactersByOperatingSystem = new List<char>() { '<', '>', ':', '|', '?', '*' };
                    }
                }

                return disallowedCharactersByOperatingSystem;
            }
        }

        #endregion Private Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in NON-ASCII files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">Optional matcher object to search for zipped files that may contain ASCII files.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public List<MatchedLine> SearchFileForMatch(string fileName, IEnumerable<string> searchTerms, Matcher matcher = null)
        {
            string fileExtension = Path.GetExtension(fileName).ToUpper();

            if (matcher != null)
            {
                this.RegexOptions = matcher.RegexOptions;
                this.IsRegexSearch = matcher.IsRegexSearch;
            }

            List<MatchedLine> matchedLines = new List<MatchedLine>();

            switch (fileExtension)
            {
                case ".7Z":
                case ".GZ":
                case ".RAR":
                case ".TAR":
                case ".ZIP":
                    matchedLines = new SearchArchive().GetMatchesInZip(fileName, searchTerms, matcher);
                    break;
                case ".DOCX":
                case ".DOCM":
                    matchedLines = new SearchWinword().GetMatchesInDocx(fileName, searchTerms, matcher);
                    break;
                case ".ODP":
                    matchedLines = new SearchOdp().GetMatchesInOdp(fileName, searchTerms, matcher);
                    break;
                case ".ODS":
                    matchedLines = new SearchOds().GetMatchesInOds(fileName, searchTerms, matcher);
                    break;
                case ".ODT":
                    matchedLines = new SearchOdt().GetMatchesInOdt(fileName, searchTerms, matcher);
                    break;
                case ".EML":
                case ".MSG":
                case ".OFT":
                    matchedLines = new SearchOutlook().GetMatchesInOutlook(fileName, searchTerms, matcher);
                    break;
                case ".PDF":
                    matchedLines = new SearchPdf().GetMatchesInPdf(fileName, searchTerms, matcher);
                    break;
                case ".PPTM":
                case ".PPTX":
                    matchedLines = new SearchPowerpoint().GetMatchesInPptx(fileName, searchTerms, matcher);
                    break;
                case ".XLSM":
                case ".XLSX":
                    matchedLines = new SearchExcel().GetMatchesInExcel(fileName, searchTerms, matcher);
                    break;
                default:
                    break;
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Remove the temporarily created directory.
        /// </summary>
        /// <param name="tempDirPath">The temporarily created directory containing archive contents.</param>
        internal void RemoveTempDirectory(string tempDirPath)
        {
            IOException fileAccessException;
            int counter = 0;    // If unable to delete after 10 attempts, get out instead of staying stuck.

            do
            {
                fileAccessException = null;

                try
                {
                    // Clean up to delete the temporarily created directory. If files are still in use, an exception will be thrown.
                    Directory.Delete(tempDirPath, true);
                }
                catch (IOException ioe)
                {
                    fileAccessException = ioe;
                    System.Threading.Thread.Sleep(counter);     // Sleep for tiny increments of time, while waiting for file to be released (max 45 milliseconds).
                    counter++;
                }
            } while (fileAccessException != null && fileAccessException.Message.ToUpper().Contains("The process cannot access the file".ToUpper()) && counter < 10);
        }

        #endregion Private Methods
    }
}
