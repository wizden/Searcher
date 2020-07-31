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
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Wordprocessing;
    using iTextSharp.text.pdf;
    using iTextSharp.text.pdf.parser;
    using SharpCompress.Common;
    using SharpCompress.Readers;

    /// <summary>
    /// Class to search files with NON-ASCII extensions.
    /// </summary>
    public class SearchOtherExtensions
    {
        #region Private Fields

        /// <summary>
        /// Private variable to hold the name of the 7-zip executable.
        /// </summary>
        private const string AppName7Zip = "7-zip";

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        private const int IndexBoundary = 50;

        /// <summary>
        /// The maximum length of a content to check. If longer than this, split the search output result to minimise the length of results content displayed.
        /// </summary>
        private const int MaxContentLengthCheck = 200;

        /// <summary>
        /// The maximum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MaxExcelDate = 2958465.99999;

        /// <summary>
        /// The minimum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MinExcelDate = -657434;

        /// <summary>
        /// Private variable to hold the name of the temporary extraction directory.
        /// </summary>
        private const string TempExtractDirectoryName = "→_extract_←";

        /// <summary>
        /// List of characters not allowed for the file system.
        /// </summary>
        private static List<char> disallowedCharactersByOperatingSystem;

        /// <summary>
        /// Private variable to hold the install location of the 7-zip executable.
        /// </summary>
        private static string fileLocation7Zip = string.Empty;

        /// <summary>
        /// Private store for the Matcher object to prevent re-creation for each call.
        /// </summary>
        private Matcher localMatcherObj = null;

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
        private static List<char> DisallowedCharactersByOperatingSystem
        {
            get
            {
                if (disallowedCharactersByOperatingSystem == null)
                {
                    disallowedCharactersByOperatingSystem = new List<char>();

                    if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        // Based on: https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file. Excluding "\" and "/" as these get used for IEntry paths.
                        disallowedCharactersByOperatingSystem.AddRange(new char[] { '<', '>', ':', '|', '?', '*' });
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
            string fileExtension = System.IO.Path.GetExtension(fileName).ToUpper();

            if (matcher != null)
            {
                this.RegexOptions = matcher.RegexOptions;
                this.IsRegexSearch = matcher.IsRegexSearch;
                this.localMatcherObj = matcher;
            }

            List<MatchedLine> matchedLines = new List<MatchedLine>();

            switch (fileExtension)
            {
                case ".7Z":
                case ".GZ":
                case ".JAR":
                case ".RAR":
                case ".TAR":
                case ".ZIP":
                    matchedLines = this.GetMatchesInZip(fileName, searchTerms);
                    break;
                case ".DOCX":
                case ".DOCM":
                    matchedLines = this.GetMatchesInDocx(fileName, searchTerms);
                    break;
                case ".PDF":
                    matchedLines = this.GetMatchesInPdf(fileName, searchTerms);
                    break;
                case ".PPTM":
                case ".PPTX":
                    matchedLines = this.GetMatchesInPptx(fileName, searchTerms);
                    break;
                case ".XLSM":
                case ".XLSX":
                    matchedLines = this.GetMatchesInXlsx(fileName, searchTerms);
                    break;
                default:
                    break;
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Decompress a GZIP file stream.
        /// </summary>
        /// <param name="fileName">The GZIP file name.</param>
        /// <param name="searchTerms">The terms to search for.</param>
        /// <returns>List of matched lines for GZIP file contents.</returns>
        private List<MatchedLine> DecompressGZipStream(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string newFileName = string.Empty;

            FileInfo fileToDecompress = new FileInfo(fileName);
            using (FileStream originalFileStream = fileToDecompress.OpenRead())
            {
                string currentFileName = fileToDecompress.FullName;
                newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);

                using (FileStream decompressedFileStream = File.Create(newFileName))
                {
                    using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(newFileName))
            {
                matchedLines.AddRange(this.localMatcherObj.GetMatch(newFileName, searchTerms));
            }

            return matchedLines;
        }

        /// <summary>
        /// Get the position of the first word.
        /// </summary>
        /// <param name="content">The match content.</param>
        /// <param name="startIndex">The index boundary.</param>
        /// <returns>Position of the first word.</returns>
        private int GetLocationOfFirstWord(string content, int startIndex)
        {
            int retVal = startIndex;

            if (retVal > 0)
            {
                while (content[retVal] != ' ' && retVal > 0)
                {
                    retVal--;
                }

                retVal++;
            }

            return retVal;
        }

        /// <summary>
        /// Get the position of the last word.
        /// </summary>
        /// <param name="content">The match content.</param>
        /// <param name="endIndex">The index boundary.</param>
        /// <returns>Position of the last word.</returns>
        private int GetLocationOfLastWord(string content, int endIndex)
        {
            int retVal = endIndex;

            if (retVal > 0)
            {
                while (retVal < (content.Length - 1) && content[retVal] != ' ')
                {
                    retVal++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Search for matches in zipped archive files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="tempDirPath">The temporary extract directory.</param>
        /// <param name="archive">The archive to be searched.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchedLinesInZipArchive(string fileName, IEnumerable<string> searchTerms, string tempDirPath, SharpCompress.Archives.IArchive archive)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                IReader reader = archive.ExtractAllEntries();
                while (reader.MoveToNextEntry())
                {
                    if (!reader.Entry.IsDirectory)
                    {
                        // Ignore symbolic links as these are captured by the original target.
                        if (string.IsNullOrWhiteSpace(reader.Entry.LinkTarget) && !reader.Entry.Key.Any(c => DisallowedCharactersByOperatingSystem.Any(dc => dc == c)))
                        {
                            try
                            {
                                reader.WriteEntryToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                                string fullFilePath = System.IO.Path.Combine(tempDirPath, reader.Entry.Key.Replace(@"/", @"\"));
                                matchedLines.AddRange(this.localMatcherObj.GetMatch(fullFilePath, searchTerms));

                                if (matchedLines != null && matchedLines.Count > 0)
                                {
                                    // Want the exact path of the file - without the .extract part.
                                    string dirNameToDisplay = fullFilePath.Replace(TempExtractDirectoryName, string.Empty);
                                    matchedLines.Where(ml => string.IsNullOrEmpty(ml.FileName) || ml.FileName.Contains(TempExtractDirectoryName)).ToList()
                                        .ForEach(ml => ml.FileName = dirNameToDisplay);
                                }
                            }
                            catch (PathTooLongException ptlex)
                            {
                                throw new PathTooLongException(string.Format("{0} {1} {2} {3} - {4}", Resources.Strings.ErrorAccessingEntry, reader.Entry.Key, Resources.Strings.InArchive, fileName, ptlex.Message));
                            }
                        }
                    }
                }

                if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    matchedLines.Clear();
                }
            }
            catch (ArgumentNullException ane)
            {
                if (ane.Message.Contains("Value cannot be null") && fileName.EndsWith(".gz", StringComparison.OrdinalIgnoreCase))
                {
                    matchedLines = this.DecompressGZipStream(fileName, searchTerms);
                }
                else if (ane.Message.Contains("String reference not set to an instance of a String."))
                {
                    throw new NotSupportedException(string.Format("{0} {1}. {2}", Resources.Strings.ErrorAccessingFile, fileName, Resources.Strings.FileEncrypted));
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in DOCX files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInDocx(string fileName, IEnumerable<string> searchTerms)
        {
            int pageNumber = 1;
            int startIndex = 0;
            int endIndex = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
                {
                    string allContent = string.Empty;
                    Body body = document.MainDocumentPart.Document.Body;
                    body.Descendants().Where(bce => bce is Paragraph && bce.HasChildren).ToList().ForEach(bce =>
                    {
                        StringBuilder contentText = new StringBuilder();                                                                                 // Set content for each paragraph and detect page breaks (dependant on word processing application).

                        foreach (OpenXmlElement r in bce.Descendants().Where(r => r is DocumentFormat.OpenXml.Wordprocessing.Run && r.HasChildren))
                        {
                            if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            r.Descendants().ToList().ForEach(rce =>
                            {
                                if (rce is DocumentFormat.OpenXml.Wordprocessing.Text)
                                {
                                    contentText.Append(rce.InnerText);
                                }

                                if (rce is LastRenderedPageBreak)
                                {
                                    pageNumber++;
                                }
                            });
                        }

                        string content = contentText.ToString();
                        allContent = content;

                        if (!string.IsNullOrEmpty(content))
                        {
                            foreach (string searchTerm in searchTerms)
                            {
                                if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    break;
                                }

                                try
                                {
                                    content = content.Trim();
                                    MatchCollection matches = Regex.Matches(content, searchTerm, this.RegexOptions);           // Use this match for getting the locations of the match.

                                    if (matches.Count > 0)
                                    {
                                        foreach (Match match in matches)
                                        {
                                            if (content.Length > MaxContentLengthCheck)
                                            {
                                                startIndex = match.Index >= IndexBoundary ? this.GetLocationOfFirstWord(content, match.Index - IndexBoundary) : 0;
                                                endIndex = (content.Length >= match.Index + match.Length + IndexBoundary) ? this.GetLocationOfLastWord(content, match.Index + match.Length + IndexBoundary) : content.Length;
                                            }
                                            else
                                            {
                                                startIndex = 0;
                                                endIndex = content.Length;
                                            }

                                            string matchLine = content.Substring(startIndex, endIndex - startIndex);
                                            Match searchMatch = Regex.Match(matchLine, searchTerm, this.RegexOptions);         // Use this match for the result highlight, based on additional characters being selected before and after the match.

                                            // Only add, if it does not already exist (Not sure how the body elements manage to bring back the same content again.
                                            if (!matchedLines.Any(ml => ml.SearchTerm == searchTerm && ml.Content == string.Format("{0} {1}:\t{2}", Resources.Strings.Page, pageNumber, matchLine)))
                                            {
                                                matchedLines.Add(new MatchedLine 
                                                { 
                                                    Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Page, pageNumber, matchLine), 
                                                    SearchTerm = searchTerm, 
                                                    FileName = fileName, 
                                                    LineNumber = pageNumber, 
                                                    StartIndex = searchMatch.Index + Resources.Strings.Page.Length + 3 + pageNumber.ToString().Length, 
                                                    Length = searchMatch.Length 
                                                });
                                            }
                                        }
                                    }
                                }
                                catch (ArgumentException aex)
                                {
                                    throw new ArgumentException(aex.Message + " " + Resources.Strings.RegexFailureSearchCancelled);
                                }
                            }
                        }
                    });

                    if (this.RegexOptions.HasFlag(RegexOptions.Multiline))
                    {
                        // Get matches not found in above list, if using Regex, since the above code will not catch matches that span across "Run" content.
                        List<MatchedLine> allContentMatchedLines = new List<MatchedLine>();
                        foreach (string searchTerm in searchTerms)
                        {
                            if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            try
                            {
                                allContent = allContent.Trim();
                                MatchCollection matches = Regex.Matches(allContent, searchTerm, this.RegexOptions);           // Use this match for getting the locations of the match.

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches)
                                    {
                                        startIndex = match.Index >= IndexBoundary ? this.GetLocationOfFirstWord(allContent, match.Index - IndexBoundary) : 0;
                                        endIndex = (allContent.Length >= match.Index + match.Length + IndexBoundary) ? this.GetLocationOfLastWord(allContent, match.Index + match.Length + IndexBoundary) : allContent.Length;
                                        string matchLine = allContent.Substring(startIndex, endIndex - startIndex);
                                        Match searchMatch = Regex.Match(matchLine, searchTerm, this.RegexOptions);         // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                        allContentMatchedLines.Add(new MatchedLine 
                                        { 
                                            Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Page, pageNumber, matchLine), 
                                            SearchTerm = searchTerm, 
                                            FileName = fileName, 
                                            LineNumber = pageNumber, 
                                            StartIndex = searchMatch.Index + Resources.Strings.Page.Length + 3 + pageNumber.ToString().Length, 
                                            Length = searchMatch.Length 
                                        });
                                    }

                                    fileName = string.Empty;
                                }
                            }
                            catch (ArgumentException aex)
                            {
                                throw new ArgumentException(aex.Message + " " + Resources.Strings.RegexFailureSearchCancelled);
                            }
                        }

                        // Only add those lines that do not already exist in matches found. Do not like the search mechanism below, but it helps to search by excluding the "Page {0}:\t{1}" content.
                        matchedLines.AddRange(allContentMatchedLines.Where(acm => !matchedLines.Any(ml => acm.Length == ml.Length && acm.Content.Contains(ml.Content.Substring(Resources.Strings.Page.Length + 3 + ml.LineNumber.ToString().Length, ml.Content.Length - (Resources.Strings.Page.Length + 3 + ml.LineNumber.ToString().Length))))));
                    }

                    document.Close();
                }

                if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    matchedLines.Clear();
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? Resources.Strings.FileCorruptOrLockedByApp
                        : Resources.Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in PDF files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInPdf(string fileName, IEnumerable<string> searchTerms)
        {
            string pdfPage = string.Empty;
            int startIndex = 0;
            int endIndex = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            using (PdfReader reader = new PdfReader(fileName))
            {
                for (int pageCounter = 1; pageCounter <= reader.NumberOfPages; pageCounter++)
                {
                    if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    try
                    {
                        ////pdfPage = PdfTextExtractor.GetTextFromPage(reader, pageCounter);                                            // Shows the result with line breaks.
                        pdfPage = Regex.Replace(PdfTextExtractor.GetTextFromPage(reader, pageCounter), @"\r\n?|\n", " ").Trim();        // Shows the result in a single line, since line breaks are removed.

                        foreach (string searchTerm in searchTerms)
                        {
                            MatchCollection matches = Regex.Matches(pdfPage, searchTerm, this.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    startIndex = match.Index >= IndexBoundary ? match.Index - IndexBoundary : 0;
                                    endIndex = (pdfPage.Length >= match.Index + match.Length + IndexBoundary) ? match.Index + match.Length + IndexBoundary : pdfPage.Length;
                                    string matchLine = pdfPage.Substring(startIndex, endIndex - startIndex);
                                    Match searchMatch = Regex.Match(matchLine, searchTerm, this.RegexOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine 
                                    { 
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Page, pageCounter.ToString(), matchLine), 
                                        SearchTerm = searchTerm, 
                                        FileName = fileName, 
                                        LineNumber = 1, 
                                        StartIndex = searchMatch.Index + Resources.Strings.Page.Length + 3 + pageCounter.ToString().Length, 
                                        Length = searchMatch.Length 
                                    });
                                }
                            }
                        }
                    }
                    catch (ArgumentException aex)
                    {
                        throw new ArgumentException(aex.Message + " " + Resources.Strings.RegexFailureSearchCancelled);
                    }
                }

                reader.Close();
            }

            if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
            {
                matchedLines.Clear();
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in PPTX files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInPptx(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                using (PresentationDocument pptDocument = PresentationDocument.Open(fileName, false))
                {
                    string[] slideAllText = this.GetPresentationSlidesText(pptDocument.PresentationPart);
                    pptDocument.Close();

                    int startIndex = 0;
                    int endIndex = 0;

                    foreach (string searchTerm in searchTerms)
                    {
                        if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            break;
                        }

                        for (int slideCounter = 0; slideCounter < slideAllText.Length; slideCounter++)
                        {
                            if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            MatchCollection matches = Regex.Matches(slideAllText[slideCounter], searchTerm, this.RegexOptions);            // Use this match for getting the locations of the match.

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    startIndex = match.Index >= IndexBoundary ? match.Index - IndexBoundary : 0;
                                    endIndex = (slideAllText[slideCounter].Length >= match.Index + match.Length + IndexBoundary) ? match.Index + match.Length + IndexBoundary : slideAllText[slideCounter].Length;
                                    string matchLine = slideAllText[slideCounter].Substring(startIndex, endIndex - startIndex);

                                    while (matchLine.StartsWith("\r") || matchLine.StartsWith("\n"))
                                    {
                                        matchLine = matchLine.Substring(1, matchLine.Length - 1);                       // Remove lines starting with the newline character.
                                    }

                                    while ((matchLine.EndsWith("\r") || matchLine.EndsWith("\n")) && matchLine.Length > 2)
                                    {
                                        matchLine = matchLine.Substring(0, matchLine.Length - 1);                       // Remove lines ending with the newline character.
                                    }

                                    Match searchMatch = Regex.Match(matchLine, searchTerm, this.RegexOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine 
                                    { 
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Slide, (slideCounter + 1).ToString(), matchLine), 
                                        SearchTerm = searchTerm, 
                                        FileName = fileName, 
                                        LineNumber = 1, 
                                        StartIndex = searchMatch.Index + Resources.Strings.Slide.Length + 3 + (slideCounter + 1).ToString().Length, 
                                        Length = searchMatch.Length 
                                    });
                                }
                            }
                        }
                    }

                    if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        matchedLines.Clear();
                    }
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? Resources.Strings.FileCorruptOrLockedByApp
                        : Resources.Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in XLSX files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInXlsx(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            List<SpreadsheetCellDetail> excelCellDetails = new List<SpreadsheetCellDetail>();

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wkbkPart = document.WorkbookPart;
                    List<Sheet> sheets = wkbkPart.Workbook.Descendants<Sheet>().ToList();
                    string cellValue = string.Empty;
                    List<OpenXmlElement> sharedStringTable = wkbkPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable.ToList();        // Get it in memory for performance.
                    CellFormats cellFormats = wkbkPart.WorkbookStylesPart.Stylesheet.CellFormats;
                    List<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat> numberingFormats = wkbkPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null
                        ? wkbkPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>().ToList()
                        : null;
                    string cellFormatCodeUpper = string.Empty;

                    foreach (Sheet sheet in sheets)
                    {
                        if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            break;
                        }

                        if (sheet == null)
                        {
                            throw new ArgumentException(Resources.Strings.SheetNotFound);
                        }

                        WorksheetPart workSheetPart = wkbkPart.GetPartById(sheet.Id) as WorksheetPart;

                        if (workSheetPart != null)
                        {
                            foreach (Cell cell in ((WorksheetPart)wkbkPart.GetPartById(sheet.Id)).Worksheet.Descendants<Cell>())
                            {
                                if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    break;
                                }

                                if (cell != null && cell.CellReference != null && !string.IsNullOrWhiteSpace(cell.InnerText))
                                {
                                    cellValue = this.GetSpreadsheetCellValue(cell, sharedStringTable, cellFormats, numberingFormats);
                                    excelCellDetails.Add(new SpreadsheetCellDetail { CellContent = cellValue, CellReference = cell.CellReference.Value, SheetName = sheet.Name.Value });
                                }
                            }
                        }
                    }

                    document.Close();
                }

                foreach (string searchTerm in searchTerms)
                {
                    if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    try
                    {
                        foreach (SpreadsheetCellDetail ecd in excelCellDetails)
                        {
                            if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            MatchCollection matches = Regex.Matches(ecd.CellContent, searchTerm, this.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    matchedLines.Add(new MatchedLine 
                                    { 
                                        Content = string.Format("{0}\\{1}\t\t{2}", ecd.SheetName, ecd.CellReference, ecd.CellContent), 
                                        SearchTerm = searchTerm, 
                                        FileName = fileName, 
                                        LineNumber = 1, 
                                        StartIndex = match.Index + (ecd.SheetName.Length + ecd.CellReference.Length + 3), 
                                        Length = match.Length 
                                    });
                                }
                            }
                        }
                    }
                    catch (ArgumentException aex)
                    {
                        throw new ArgumentException(aex.Message + " " + Resources.Strings.RegexFailureSearchCancelled);
                    }
                }

                if (this.localMatcherObj.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    matchedLines.Clear();
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? Resources.Strings.FileCorruptOrLockedByApp
                        : Resources.Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in zipped files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInZip(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                SharpCompress.Archives.IArchive archive = null;

                if (fileName.ToUpper().EndsWith(".GZ") && SharpCompress.Archives.GZip.GZipArchive.IsGZipFile(fileName))
                {
                    archive = SharpCompress.Archives.GZip.GZipArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".RAR") && SharpCompress.Archives.Rar.RarArchive.IsRarFile(fileName))
                {
                    archive = SharpCompress.Archives.Rar.RarArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".7Z") && SharpCompress.Archives.SevenZip.SevenZipArchive.IsSevenZipFile(fileName))
                {
                    archive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".TAR") && SharpCompress.Archives.Tar.TarArchive.IsTarFile(fileName))
                {
                    archive = SharpCompress.Archives.Tar.TarArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".ZIP") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
                {
                    archive = SharpCompress.Archives.Zip.ZipArchive.Open(fileName);
                }

                if (archive != null)
                {
                    matchedLines = this.GetMatchedLinesInZipArchive(fileName, searchTerms, tempDirPath, archive);
                    archive.Dispose();
                }

                this.RemoveTempDirectory(tempDirPath);
                return matchedLines;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Returns a string array of the text in the presentation slides.
        /// </summary>
        /// <param name="presentationPart">The presentation part of the presentation document.</param>
        /// <returns>String array of the text in the presentation slides.</returns>
        private string[] GetPresentationSlidesText(PresentationPart presentationPart)
        {
            Presentation presentation = presentationPart.Presentation;
            List<SlidePart> slideParts = presentationPart.SlideParts.ToList();
            string[] retVal = new string[slideParts.Count()];
            string relationshipId = string.Empty;
            int tempSlideNumber = 0;

            foreach (SlidePart slidePart in slideParts)
            {
                var slide = presentation.SlideIdList.Where(s => ((SlideId)s).RelationshipId == presentationPart.GetIdOfPart(slidePart)).FirstOrDefault();
                int index = presentation.SlideIdList.ToList().IndexOf(slide);
                relationshipId = ((SlideId)presentation.SlideIdList.ChildElements[index]).RelationshipId;
                List<string> titles = new List<string>();
                List<string> content = new List<string>();
                List<string> notes = new List<string>();

                slidePart.Slide.Descendants<Shape>().ToList().ForEach(shape =>
                {
                    foreach (PlaceholderShape item in shape.Descendants<PlaceholderShape>().Where(i => i.Type != null))
                    {
                        if ((item.Type.ToString().ToUpper() == "CenteredTitle".ToUpper() || item.Type.ToString().ToUpper() == "SubTitle".ToUpper() || item.Type.ToString().ToUpper() == "Title".ToUpper())
                                && shape.TextBody != null && !string.IsNullOrWhiteSpace(shape.TextBody.InnerText))
                        {
                            titles.AddRange(shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text));
                        }
                    }
                });

                content = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text).ToList();
                content.RemoveAll(s => titles.Any(t => t == s));

                if (slidePart.NotesSlidePart != null && slidePart.NotesSlidePart.NotesSlide != null && slidePart.NotesSlidePart.NotesSlide.Descendants() != null && slidePart.NotesSlidePart.NotesSlide.Descendants().Count() > 0)
                {
                    notes = slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                        .Where(s =>
                        {
                            return !int.TryParse(s.Text, out tempSlideNumber);      // Remove the record as it contains the slide number.
                        }).Select(s => s.Text).ToList();
                }

                retVal[index] = string.Join(
                    Environment.NewLine, 
                    new string[] { string.Join(Environment.NewLine, titles.ToArray()), string.Join(string.Empty, content.ToArray()), string.Join(Environment.NewLine, notes.ToArray()) });
            }

            return retVal;
        }

        /// <summary>
        /// Gets the content of the spread sheet cell based on the data type of the cell.
        /// </summary>
        /// <param name="excelCell">The spread sheet cell object.</param>
        /// <param name="sharedStringTable">The table containing the shared strings.</param>
        /// <param name="cellFormats">The cell formats in use.</param>
        /// <param name="numberingFormats">The numbering formats in use.</param>
        /// <returns>The content of the spread sheet cell based on the data type of the cell.</returns>
        private string GetSpreadsheetCellValue(Cell excelCell, IEnumerable<OpenXmlElement> sharedStringTable, CellFormats cellFormats, IEnumerable<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat> numberingFormats)
        {
            string retVal = string.Empty;

            retVal = excelCell.InnerText;
            if (excelCell.DataType != null)
            {
                switch (excelCell.DataType.Value)
                {
                    case CellValues.SharedString:
                        if (sharedStringTable != null)
                        {
                            retVal = sharedStringTable.ElementAt(int.Parse(retVal)).InnerText;
                        }

                        break;
                    case CellValues.Boolean:
                        switch (retVal)
                        {
                            case "0":
                                retVal = "FALSE";
                                break;
                            default:
                                retVal = "TRUE";
                                break;
                        }

                        break;
                }
            }
            else
            {
                if (excelCell.StyleIndex != null)
                {
                    UInt32Value styleIndex = excelCell.StyleIndex;
                    CellFormat cellFormat = (CellFormat)cellFormats.ElementAt((int)styleIndex.Value);
                    double tempDouble;

                    // Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference - Section: 18.8.30 numFmt (Number Format)
                    if (cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.Value < 163 && retVal != "0")
                    {
                        if (double.TryParse(retVal.ToString(), out tempDouble))
                        {
                            if (tempDouble >= MinExcelDate && tempDouble <= MaxExcelDate)
                            {
                                try
                                {
                                    retVal = DateTime.FromOADate(double.Parse(retVal.ToString())).ToString();
                                }
                                catch (ArgumentException ae)
                                {
                                    // Throw only if it is not an OleAut date error.
                                    if (!ae.Message.Contains("Not a legal OleAut date"))
                                    {
                                        throw;
                                    }
                                }
                            }
                        }
                    }
                    else if (numberingFormats != null && cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.Value > 163)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.NumberingFormat cellFormatUsed = numberingFormats.FirstOrDefault(nf => nf.NumberFormatId == cellFormat.NumberFormatId.Value);

                        if (cellFormatUsed.FormatCode.Value.ToUpper().Contains("D") || cellFormatUsed.FormatCode.Value.ToUpper().Contains("M") || cellFormatUsed.FormatCode.Value.ToUpper().Contains("Y"))
                        {
                            if (double.TryParse(retVal.ToString(), out tempDouble))
                            {
                                retVal = DateTime.FromOADate(double.Parse(retVal.ToString())).ToString();
                            }
                        }
                    }
                }
            }

            return retVal;
        }

        /// <summary>
        /// Remove the temporarily created directory.
        /// </summary>
        /// <param name="tempDirPath">The temporarily created directory containing archive contents.</param>
        private void RemoveTempDirectory(string tempDirPath)
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

        /// <summary>
        /// Internal class to hold the cell information of an spread sheet document.
        /// </summary>
        protected internal class SpreadsheetCellDetail
        {
            /// <summary>
            /// Gets or sets the content of the cell.
            /// </summary>
            public string CellContent { get; set; }

            /// <summary>
            /// Gets or sets the cell reference (address).
            /// </summary>
            public string CellReference { get; set; }

            /// <summary>
            /// Gets or sets the name displayed for the sheet.
            /// </summary>
            public string SheetName { get; set; }
        }

        #endregion Private Methods
    }
}
