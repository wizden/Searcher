// <copyright file="ExcelSearchHandler.cs" company="dennjose">
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
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

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
        public static new List<string> Extensions => new List<string> { ".XLSX", ".XLSM" };

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
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            List<SpreadsheetCellDetail> excelCellDetails = new List<SpreadsheetCellDetail>();

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wkbkPart = document.WorkbookPart;
                    List<Sheet> sheets = wkbkPart.Workbook.Descendants<Sheet>().ToList();
                    string cellValue = string.Empty;
                    List<OpenXmlElement> sharedStringTable = wkbkPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable?.ToList();        // Get it in memory for performance.
                    CellFormats cellFormats = wkbkPart.WorkbookStylesPart.Stylesheet.CellFormats;
                    List<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat> numberingFormats = wkbkPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null
                        ? wkbkPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>().ToList()
                        : null;
                    string cellFormatCodeUpper = string.Empty;
                    int counter = 0;

                    foreach (Sheet sheet in sheets)
                    {
                        if (sheet == null)
                        {
                            throw new ArgumentException(Resources.Strings.SheetNotFound);
                        }

                        WorksheetPart workSheetPart = wkbkPart.GetPartById(sheet.Id) as WorksheetPart;

                        if (workSheetPart != null)
                        {
                            foreach (Cell cell in ((WorksheetPart)wkbkPart.GetPartById(sheet.Id)).Worksheet.Descendants<Cell>())
                            {
                                counter++;

                                if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    break;
                                }

                                if (cell != null && cell.CellReference != null && !string.IsNullOrWhiteSpace(cell.InnerText))
                                {
                                    cellValue = this.GetSpreadsheetCellValue(cell, sharedStringTable, cellFormats, numberingFormats);
                                    excelCellDetails.Add(new SpreadsheetCellDetail { CellContent = cellValue, CellReference = cell.CellReference.Value, SheetName = sheet.Name.Value });
                                }

                                if (counter % 10000 == 0)
                                {
                                    counter = 0;
                                    System.Threading.Thread.Sleep(100);     // Large dataset. Give the UI 100 milliseconds to refresh itself.
                                }
                            }
                        }
                    }

                    document.Close();
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

                            MatchCollection matches = Regex.Matches(ecd.CellContent, searchTerm, this.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
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
            internal string CellContent { get; set; }

            /// <summary>
            /// Gets or sets the cell reference (address).
            /// </summary>
            internal string CellReference { get; set; }

            /// <summary>
            /// Gets or sets the name displayed for the sheet.
            /// </summary>
            internal string SheetName { get; set; }

            #endregion Internal Properties
        }

        #endregion Protected Internal Classes
    }
}
