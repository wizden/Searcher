// <copyright file="PdfSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using iTextSharp.text.pdf;
    using iTextSharp.text.pdf.parser;
    using SearcherLibrary.Resources;

    /// <summary>
    /// Class to search PDF files.
    /// </summary>
    public class PdfSearchHandler : FileSearchHandler
    {
        #region Internal Fields

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        internal const int MaxIndexBoundary = 50;

        #endregion Internal Fields

        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => new List<string> { ".PDF" };

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
            var matchCounter = 0;
            var matchedLines = new List<MatchedLine>();

            using (var reader = new PdfReader(fileName))
            {
                for (var pageCounter = 1; pageCounter <= reader.NumberOfPages; pageCounter++)
                {
                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    try
                    {
                        ////pdfPage = PdfTextExtractor.GetTextFromPage(reader, pageCounter);                                            // Shows the result with line breaks.
                        var pdfPage = Regex.Replace(PdfTextExtractor.GetTextFromPage(reader, pageCounter), @"\r\n?|\n", " ").Trim();

                        foreach (var searchTerm in searchTerms)
                        {
                            var matches = Regex.Matches(pdfPage, searchTerm, matcher.RegexOptions); // Use this match for getting the locations of the match.

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    var startIndex = match.Index >= MaxIndexBoundary ? match.Index - MaxIndexBoundary : 0;
                                    var endIndex = pdfPage.Length >= match.Index + match.Length + MaxIndexBoundary ? match.Index + match.Length + MaxIndexBoundary : pdfPage.Length;
                                    var matchLine = pdfPage.Substring(startIndex, endIndex - startIndex);
                                    var searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegexOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
                                        Content = string.Format("{0} {1}:\t{2}", Strings.Page, pageCounter.ToString(), matchLine),
                                        SearchTerm = searchTerm,
                                        FileName = fileName,
                                        LineNumber = 1,
                                        StartIndex = searchMatch.Index + Strings.Page.Length + 3 + pageCounter.ToString().Length,
                                        Length = searchMatch.Length
                                    });
                                }
                            }
                        }
                    }
                    catch (ArgumentException aex)
                    {
                        throw new ArgumentException(aex.Message + " " + Strings.RegexFailureSearchCancelled);
                    }

                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        matchedLines.Clear();
                    }
                }

                reader.Close();
            }

            return matchedLines;
        }

        #endregion Public Methods
    }
}
