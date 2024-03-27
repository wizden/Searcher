// <copyright file="OdsSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
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
    using System.Text.RegularExpressions;
    using System.Xml.Linq;
    using SharpCompress.Common;
    using SharpCompress.Readers;

    /// <summary>
    /// Class to search ODS files.
    /// </summary>
    public class OdsSearchHandler : FileSearchHandler
    {
        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => new() { ".ODS" };

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .PDF files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new();
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                SharpCompress.Archives.IArchive? archive = null;

                if (fileName.ToUpper().EndsWith(".ODS") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
                {
                    archive = SharpCompress.Archives.Zip.ZipArchive.Open(fileName);
                }

                if (archive != null)
                {
                    IReader reader = archive.ExtractAllEntries();
                    while (reader.MoveToNextEntry())
                    {
                        if (!reader.Entry.IsDirectory && reader.Entry.Key == "content.xml")
                        {
                            // Ignore symbolic links as these are captured by the original target.
                            if (string.IsNullOrWhiteSpace(reader.Entry.LinkTarget) && !reader.Entry.Key.Any(c => DisallowedCharactersByOperatingSystem.Any(dc => dc == c)))
                            {
                                try
                                {
                                    reader.WriteEntryToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                                    string fullFilePath = System.IO.Path.Combine(tempDirPath, reader.Entry.Key.Replace(@"/", @"\"));
                                    matchedLines = GetMatchesFromOdsContentXml(fileName, fullFilePath, searchTerms, matcher);
                                }
                                catch (PathTooLongException ptlex)
                                {
                                    throw new PathTooLongException(string.Format("{0} {1} {2} {3} - {4}", Resources.Strings.ErrorAccessingEntry, reader.Entry.Key, Resources.Strings.InArchive, fileName, ptlex.Message));
                                }
                            }
                        }
                    }

                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        matchedLines.Clear();
                    }

                    archive.Dispose();
                }
            }
            finally
            {
                RemoveTempDirectory(tempDirPath);
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Get the matched items from the content.xml file for the document.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="extractedContentFile">The name of the extracted content xml file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private static List<MatchedLine> GetMatchesFromOdsContentXml(string fileName, string extractedContentFile, IEnumerable<string> searchTerms, Matcher matcher)
        {
            int matchCounter = 0;
            List<MatchedLine> matchedLines = new();

            // Loop through each sheet.
            foreach (XElement element in XDocument.Load(extractedContentFile, LoadOptions.None).Descendants().Where(d => d.Name.LocalName == "table"))
            {
                string sheetName = element.Attributes().Where(a => a.Name.LocalName == "name").Select(a => a.Value).FirstOrDefault() ?? string.Empty;
                int rowCount = 0;

                // Loop through each table-row.
                foreach (XElement tableContent in element.Elements().Where(e => e.Name.LocalName == "table-row"))
                {
                    int colCount = 1;

                    string emptyRowCount = tableContent.Attributes().Where(c => c.Name.LocalName == "number-rows-repeated").Select(c => c.Value).FirstOrDefault() ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(emptyRowCount))
                    {
                        rowCount += int.Parse(emptyRowCount);
                    }
                    else
                    {
                        rowCount++;
                    }

                    // Loop through each table-cell.
                    foreach (XElement rowContent in tableContent.Elements().Where(e => e.Name.LocalName == "table-cell"))
                    {
                        string emptyColCount = rowContent.Attributes().Where(c => c.Name.LocalName == "number-columns-repeated").Select(c => c.Value).FirstOrDefault() ?? string.Empty;

                        if (!string.IsNullOrWhiteSpace(emptyColCount))
                        {
                            colCount += int.Parse(emptyColCount);
                        }

                        // Loop through each table-cell content.
                        string rowColumnLocation = $"{GetSpreadsheetColumnNameFromIndex(colCount)}{rowCount}";

                        foreach (string searchTerm in searchTerms)
                        {
                            string cellContent = string.Join(Environment.NewLine, rowContent.Elements().Where(e => e.Name.LocalName == "p").Select(e => e.Value));
                            MatchCollection matches = Regex.Matches(cellContent, searchTerm, matcher.RegularExpressionOptions);

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches.Cast<Match>())
                                {
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
                                        Content = string.Format("{0}\\{1}\t\t{2}", sheetName, rowColumnLocation, cellContent),
                                        SearchTerm = searchTerm,
                                        FileName = fileName,
                                        LineNumber = 1,
                                        StartIndex = match.Index + (sheetName.Length + rowColumnLocation.Length + 3),
                                        Length = match.Length
                                    });
                                }
                            }
                        }
                    }
                }

                if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    break;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Get the column header based on index.
        /// </summary>
        /// <param name="index">The index value.</param>
        /// <returns>The column header value.</returns>
        private static string GetSpreadsheetColumnNameFromIndex(int index)
        {
            // Based on: https://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
            index -= 1;     // Adjust so it matches 0-indexed array rather than 1-indexed column

            int quotient = index / 26;
            return (quotient > 0)
                ? GetSpreadsheetColumnNameFromIndex(quotient) + ((char)((index % 26) + 'A')).ToString()
                : ((char)((index % 26) + 'A')).ToString();
        }

        #endregion Private Methods
    }
}