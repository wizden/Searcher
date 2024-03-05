// <copyright file="PdfSearchHandler.cs" company="dennjose">
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
    using System.Text;
    using System.Text.RegularExpressions;
    using SearcherLibrary.Resources;
    using UglyToad.PdfPig;
    using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

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
        public static new List<string> Extensions => new() { ".PDF" };

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

            using (var document = PdfDocument.Open(fileName))
            {
                StringBuilder documentAllContent = new();
                int pageCounter = 0;

                foreach (var page in document.GetPages())
                {
                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    try
                    {
                        pageCounter++;
                        var pdfPage = Regex.Replace(ContentOrderTextExtractor.GetText(page), @"\r\n?|\n", " ").Trim();

                        if (matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                        {
                            documentAllContent.AppendLine(pdfPage);
                        }

                        if (!matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                        {

                            foreach (var searchTerm in searchTerms)
                            {
                                var matches = Regex.Matches(pdfPage, searchTerm, matcher.RegularExpressionOptions); // Use this match for getting the locations of the match.

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches.Cast<Match>())
                                    {
                                        var startIndex = match.Index >= MaxIndexBoundary ? match.Index - MaxIndexBoundary : 0;
                                        var endIndex = pdfPage.Length >= match.Index + match.Length + MaxIndexBoundary ? match.Index + match.Length + MaxIndexBoundary : pdfPage.Length;
                                        var matchLine = pdfPage[startIndex..endIndex];
                                        var searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegularExpressionOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
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

                if (matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                {
                    matchedLines = matcher.GetMatch(new string[] { string.Join(string.Empty, documentAllContent.ToString()) }, searchTerms, Strings.Page);
                }
            }

            return matchedLines;
        }

        #endregion Public Methods
    }
}
