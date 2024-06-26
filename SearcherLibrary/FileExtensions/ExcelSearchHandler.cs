﻿// <copyright file="ExcelSearchHandler.cs" company="dennjose">
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

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Class to search XLSX/XLSM files.
    /// </summary>
    public class ExcelSearchHandler : FileSearchHandler
    {
        #region Private Fields

        /// <summary>
        /// The maximum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MaxExcelDate = 2958465.99999;

        /// <summary>
        /// The minimum date value in excel (see https://support.office.com/en-us/article/move-data-from-excel-to-access-90c35a40-bcc3-46d9-aa7f-4106f78850b4).
        /// </summary>
        private const double MinExcelDate = -657434;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => [".XLSX", ".XLSM"];

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .XLSX/.XLSM files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            int matchCounter = 0;
            List<MatchedLine> matchedLines = [];
            List<SpreadsheetCellDetail> excelCellDetails = [];

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart? wkbkPart = document.WorkbookPart;

                    if (wkbkPart != null)
                    {
                        List<Sheet> sheets = wkbkPart.Workbook.Descendants<Sheet>().ToList();
                        string cellValue = string.Empty;
                        // Get it in memory for performance.
                        List<OpenXmlElement> sharedStringTable = wkbkPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable?.ToList()
                            ?? [];
                        CellFormats cellFormats = wkbkPart.WorkbookStylesPart?.Stylesheet.CellFormats ?? new CellFormats();
                        List<NumberingFormat> numberingFormats = wkbkPart.WorkbookStylesPart?.Stylesheet.NumberingFormats?.Elements<NumberingFormat>().ToList()
                            ?? [];
                        string cellFormatCodeUpper = string.Empty;
                        int counter = 0;

                        foreach (Sheet sheet in sheets)
                        {
                            if (sheet == null)
                            {
                                throw new ArgumentException(Resources.Strings.SheetNotFound);
                            }

                            var sheetIdValue = sheet.Id?.Value ?? string.Empty;
                            if (wkbkPart.GetPartById(sheetIdValue) is WorksheetPart workSheetPart)
                            {
                                foreach (Cell cell in ((WorksheetPart)wkbkPart.GetPartById(sheetIdValue)).Worksheet.Descendants<Cell>())
                                {
                                    counter++;

                                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                                    {
                                        break;
                                    }

                                    if (cell != null && cell.CellReference != null && !string.IsNullOrWhiteSpace(cell.InnerText))
                                    {
                                        cellValue = GetSpreadsheetCellValue(cell, sharedStringTable, cellFormats, numberingFormats);
                                        excelCellDetails.Add(new SpreadsheetCellDetail { CellContent = cellValue, CellReference = cell.CellReference.Value ?? string.Empty, SheetName = sheet.Name?.Value ?? string.Empty });
                                    }

                                    if (counter % 10000 == 0)
                                    {
                                        counter = 0;
                                        System.Threading.Thread.Sleep(100);     // Large dataset. Give the UI 100 milliseconds to refresh itself.
                                    }
                                }
                            }
                        }
                    }
                }

                foreach (string searchTerm in searchTerms)
                {
                    try
                    {
                        foreach (SpreadsheetCellDetail ecd in excelCellDetails)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            MatchCollection matches = Regex.Matches(ecd.CellContent, searchTerm, matcher.RegularExpressionOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches.Cast<Match>())
                                {
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
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

                if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    matchedLines.Clear();
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$") ? Resources.Strings.FileCorruptOrLockedByApp : Resources.Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Gets the content of the spread sheet cell based on the data type of the cell.
        /// </summary>
        /// <param name="excelCell">The spread sheet cell object.</param>
        /// <param name="sharedStringTable">The table containing the shared strings.</param>
        /// <param name="cellFormats">The cell formats in use.</param>
        /// <param name="numberingFormats">The numbering formats in use.</param>
        /// <returns>The content of the spread sheet cell based on the data type of the cell.</returns>
        private static string GetSpreadsheetCellValue(Cell excelCell, IEnumerable<OpenXmlElement> sharedStringTable, CellFormats cellFormats, IEnumerable<NumberingFormat> numberingFormats)
        {
            string retVal = string.Empty;

            retVal = excelCell.InnerText;
            if (excelCell.DataType != null)
            {
                if (excelCell.DataType.Value == CellValues.SharedString)
                {
                    if (sharedStringTable != null)
                    {
                        retVal = sharedStringTable.ElementAt(int.Parse(retVal)).InnerText;
                    }
                }
                else if (excelCell.DataType.Value == CellValues.Boolean)
                {
                    retVal = retVal switch
                    {
                        "0" => "FALSE",
                        _ => "TRUE",
                    };
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
                        NumberingFormat? cellFormatUsed = numberingFormats.FirstOrDefault(nf => nf.NumberFormatId?.Value == cellFormat.NumberFormatId.Value);

                        if (cellFormatUsed != null && cellFormatUsed.FormatCode != null && !string.IsNullOrWhiteSpace(cellFormatUsed.FormatCode.Value))
                        {
                            if (cellFormatUsed.FormatCode.Value.ToUpper().Contains('D')
                                || cellFormatUsed.FormatCode.Value.ToUpper().Contains('M')
                                || cellFormatUsed.FormatCode.Value.ToUpper().Contains('Y'))
                            {
                                if (double.TryParse(retVal.ToString(), out tempDouble))
                                {
                                    retVal = DateTime.FromOADate(double.Parse(retVal.ToString())).ToString();
                                }
                            }
                        }
                    }
                }
            }

            return retVal;
        }

        #endregion Private Methods

        #region Protected Internal Classes

        /// <summary>
        /// Internal class to hold the cell information of an spread sheet document.
        /// </summary>
        protected internal class SpreadsheetCellDetail
        {
            #region Internal Properties

            /// <summary>
            /// Gets or sets the content of the cell.
            /// </summary>
            internal string CellContent { get; set; } = string.Empty;

            /// <summary>
            /// Gets or sets the cell reference (address).
            /// </summary>
            internal string CellReference { get; set; } = string.Empty;

            /// <summary>
            /// Gets or sets the name displayed for the sheet.
            /// </summary>
            internal string SheetName { get; set; } = string.Empty;

            #endregion Internal Properties
        }

        #endregion Protected Internal Classes
    }
}
