// <copyright file="OdsSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
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
        public static new List<string> Extensions => new List<string> { ".ODS" };

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
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                SharpCompress.Archives.IArchive archive = null;

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
                                    matchedLines = this.GetMatchesFromOdsContentXml(fileName, fullFilePath, searchTerms, matcher);
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                this.RemoveTempDirectory(tempDirPath);
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
        private List<MatchedLine> GetMatchesFromOdsContentXml(string fileName, string extractedContentFile, IEnumerable<string> searchTerms, Matcher matcher)
        {
            int matchCounter = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            // Loop through each sheet.
            foreach (XElement element in XDocument.Load(extractedContentFile, LoadOptions.None).Descendants().Where(d => d.Name.LocalName == "table"))
            {
                string sheetName = element.Attributes().Where(a => a.Name.LocalName == "name").Select(a => a.Value).FirstOrDefault();
                int rowCount = 0;

                // Loop through each table-row.
                foreach (XElement tableContent in element.Elements().Where(e => e.Name.LocalName == "table-row"))
                {
                    int colCount = 1;

                    string emptyRowCount = tableContent.Attributes().Where(c => c.Name.LocalName == "number-rows-repeated").Select(c => c.Value).FirstOrDefault();

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
                        string emptyColCount = rowContent.Attributes().Where(c => c.Name.LocalName == "number-columns-repeated").Select(c => c.Value).FirstOrDefault();

                        if (!string.IsNullOrWhiteSpace(emptyColCount))
                        {
                            colCount += int.Parse(emptyColCount);
                        }

                        // Loop through each table-cell content.
                        foreach (XElement cellContent in rowContent.Elements().Where(e => e.Name.LocalName == "p"))
                        {
                            string rowColumnLocation = $"{GetSpreadsheetColumnNameFromIndex(colCount)}{rowCount}";

                            foreach (string searchTerm in searchTerms)
                            {
                                MatchCollection matches = Regex.Matches(cellContent.Value, searchTerm, this.RegexOptions);

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches)
                                    {
                                        matchedLines.Add(new MatchedLine
                                        {
                                            MatchId = matchCounter++,
                                            Content = string.Format("{0}\\{1}\t\t{2}", sheetName, rowColumnLocation, cellContent.Value),
                                            SearchTerm = searchTerm,
                                            FileName = fileName,
                                            LineNumber = 1,
                                            StartIndex = match.Index + (sheetName.Length + rowColumnLocation.Length + 3),
                                            Length = match.Length
                                        });
                                    }
                                }
                            }

                            colCount++;
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
        private string GetSpreadsheetColumnNameFromIndex(int index)
        {
            // Based on: https://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
            index -= 1;     // Adjust so it matches 0-indexed array rather than 1-indexed column

            int quotient = index / 26;
            return (quotient > 0)
                ? this.GetSpreadsheetColumnNameFromIndex(quotient) + ((char)((index % 26) + 'A')).ToString()
                : ((char)((index % 26) + 'A')).ToString();
        }

        #endregion Private Methods
    }
}