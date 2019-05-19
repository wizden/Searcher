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
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
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
        /// Private variable to hold the name of the temporary extraction directory.
        /// </summary>
        private const string TempExtractDirectoryName = "→_extract_←";

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        private const int IndexBoundary = 50;

        /// <summary>
        /// The maximum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MaxExcelDate = 2958465.99999;

        /// <summary>
        /// The minimum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MinExcelDate = -657434;

        /// <summary>
        /// The maximum length of a content to check. If longer than this, split the search output result to minimise the length of results content displayed.
        /// </summary>
        private const int MaxContentLengthCheck = 200;

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
        /// Gets or sets the cancellation token source object to cancel file search.
        /// </summary>
        public CancellationTokenSource CancellationTokenSource { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the search mode uses regex.
        /// </summary>
        public bool IsRegexSearch { get; set; }

        /// <summary>
        /// Gets or sets the regex options to use when searching.
        /// </summary>
        public RegexOptions RegexOptions { get; set; }

        #endregion Public Properties

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
                this.CancellationTokenSource = matcher.CancellationTokenSource;
                this.localMatcherObj = matcher;
            }

            List<MatchedLine> matchedLines = new List<MatchedLine>();

            switch (fileExtension)
            {
                case ".7Z":
                case ".GZ":
                case ".RAR":
                case ".TAR":
                case ".ZIP":
                    matchedLines = this.GetMatchesInZip(fileName, searchTerms);
                    break;
                case ".DOCX":
                    matchedLines = this.GetMatchesInDocx(fileName, searchTerms);
                    break;
                case ".PDF":
                    matchedLines = this.GetMatchesInPdf(fileName, searchTerms);
                    break;
                case ".PPTX":
                    matchedLines = this.GetMatchesInPptx(fileName, searchTerms);
                    break;
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

                        bce.Descendants().Where(r => r is DocumentFormat.OpenXml.Wordprocessing.Run && r.HasChildren).ToList().ForEach(r =>
                        {
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
                        });

                        string content = contentText.ToString();
                        allContent = content;

                        if (!string.IsNullOrEmpty(content))
                        {
                            foreach (string searchTerm in searchTerms)
                            {
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
                                            if (!matchedLines.Any(ml => ml.SearchTerm == searchTerm && ml.Content == string.Format("Page {0}:\t{1}", pageNumber, matchLine)))
                                            {
                                                matchedLines.Add(new MatchedLine { Content = string.Format("Page {0}:\t{1}", pageNumber, matchLine), SearchTerm = searchTerm, FileName = fileName, LineNumber = pageNumber, StartIndex = searchMatch.Index + 7 + pageNumber.ToString().Length, Length = searchMatch.Length });
                                            }
                                        }
                                    }
                                }
                                catch (ArgumentException aex)
                                {
                                    throw new ArgumentException(aex.Message + " If using Regex, try escaping using the \\ character. Regex failures - Search cancelled. Correct regular expression and retry.");
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
                                        allContentMatchedLines.Add(new MatchedLine { Content = string.Format("Page {0}:\t{1}", pageNumber, matchLine), SearchTerm = searchTerm, FileName = fileName, LineNumber = pageNumber, StartIndex = searchMatch.Index + 7 + pageNumber.ToString().Length, Length = searchMatch.Length });
                                    }

                                    fileName = string.Empty;
                                }
                            }
                            catch (ArgumentException aex)
                            {
                                throw new ArgumentException(aex.Message + " If using Regex, try escaping using the \\ character. Regex failures - Search cancelled. Correct regular expression and retry.");
                            }
                        }

                        // Only add those lines that do not already exist in matches found. Do not like the search mechanism below, but it helps to search by excluding the "Page {0}:\t{1}" content.
                        matchedLines.AddRange(allContentMatchedLines.Where(acm => !matchedLines.Any(ml => acm.Length == ml.Length && acm.Content.Contains(ml.Content.Substring(7 + ml.LineNumber.ToString().Length, ml.Content.Length - (7 + ml.LineNumber.ToString().Length))))));
                    }

                    document.Close();
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? "The file is either corrupt or open in another application."
                        : "The file is either corrupt or protected.";
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
                                    matchedLines.Add(new MatchedLine { Content = string.Format("Page {0}:\t{1}", pageCounter.ToString(), matchLine), SearchTerm = searchTerm, FileName = fileName, LineNumber = 1, StartIndex = searchMatch.Index + 7 + pageCounter.ToString().Length, Length = searchMatch.Length });
                                }
                            }
                        }
                    }
                    catch (ArgumentException aex)
                    {
                        throw new ArgumentException(aex.Message + " If using Regex, try escaping using the \\ character. Regex failures - Search cancelled. Correct regular expression and retry.");
                    }
                }

                reader.Close();
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
                        for (int slideCounter = 0; slideCounter < slideAllText.Length; slideCounter++)
                        {
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
                                    matchedLines.Add(new MatchedLine { Content = string.Format("Slide {0}:\t{1}", (slideCounter + 1).ToString(), matchLine), SearchTerm = searchTerm, FileName = fileName, LineNumber = 1, StartIndex = searchMatch.Index + 8 + (slideCounter + 1).ToString().Length, Length = searchMatch.Length });
                                }
                            }
                        }
                    }
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? "The file is either corrupt or open in another application."
                        : "The file is either corrupt or protected.";
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
                using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fileName, false))
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
                        if (sheet == null)
                        {
                            throw new ArgumentException("Sheet Not Found");
                        }

                        WorksheetPart workSheetPart = (WorksheetPart)wkbkPart.GetPartById(sheet.Id);

                        foreach (Cell cell in ((WorksheetPart)wkbkPart.GetPartById(sheet.Id)).Worksheet.Descendants<Cell>())
                        {
                            if (cell != null && !string.IsNullOrWhiteSpace(cell.InnerText))
                            {
                                cellValue = this.GetSpreadsheetCellValue(cell, sharedStringTable, cellFormats, numberingFormats);
                                excelCellDetails.Add(new SpreadsheetCellDetail { CellContent = cellValue, CellReference = cell.CellReference.Value, SheetName = sheet.Name.Value });
                            }
                        }
                    }

                    document.Close();
                }

                foreach (string searchTerm in searchTerms)
                {
                    try
                    {
                        excelCellDetails.ForEach(ecd =>
                        {
                            MatchCollection matches = Regex.Matches(ecd.CellContent, searchTerm, this.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    matchedLines.Add(new MatchedLine { Content = string.Format("{0}\\{1}\t\t{2}", ecd.SheetName, ecd.CellReference, ecd.CellContent), SearchTerm = searchTerm, FileName = fileName, LineNumber = 1, StartIndex = match.Index + (ecd.SheetName.Length + ecd.CellReference.Length + 3), Length = match.Length });
                                }
                            }
                        });
                    }
                    catch (ArgumentException aex)
                    {
                        throw new ArgumentException(aex.Message + " If using Regex, try escaping using the \\ character. Regex failures - Search cancelled. Correct regular expression and retry.");
                    }
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? "The file is either corrupt or open in another application."
                        : "The file is either corrupt or protected.";
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

                return matchedLines;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                // Clean up to delete the temporarily created directory.
                Directory.Delete(tempDirPath, true);
            }
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
                        try
                        {
                            reader.WriteEntryToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                            string fullFilePath = System.IO.Path.Combine(tempDirPath, reader.Entry.Key);
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
                            throw new PathTooLongException(string.Format("Error accessing entry {0} in archive {1} - {2}", reader.Entry.Key, fileName, ptlex.Message));
                        }
                    }
                }
            }
            catch (ArgumentNullException ane)
            {
                if (ane.Message.Contains("String reference not set to an instance of a String."))
                {
                    throw new NotSupportedException(string.Format("Unable to access {0}. The file may be encrypted.", fileName));
                }
                else
                {
                    throw;
                }
            }
            return matchedLines;
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
