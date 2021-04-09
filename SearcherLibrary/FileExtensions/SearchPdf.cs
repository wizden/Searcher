// <copyright file="SearchPdf.cs" company="dennjose">
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

    /// <summary>
    /// Class to search PDF files.
    /// </summary>
    internal class SearchPdf : SearchOtherExtensions
    {
        /// <summary>
        /// Search for matches in PDF files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInPdf(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            string pdfPage = string.Empty;
            int startIndex = 0;
            int endIndex = 0;
            int matchCounter = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            using (PdfReader reader = new PdfReader(fileName))
            {
                for (int pageCounter = 1; pageCounter <= reader.NumberOfPages; pageCounter++)
                {
                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    try
                    {
                        ////pdfPage = PdfTextExtractor.GetTextFromPage(reader, pageCounter);                                            // Shows the result with line breaks.
                        pdfPage = Regex.Replace(PdfTextExtractor.GetTextFromPage(reader, pageCounter), @"\r\n?|\n", " ").Trim();        // Shows the result in a single line, since line breaks are removed.

                        foreach (string searchTerm in searchTerms)
                        {
                            MatchCollection matches = Regex.Matches(pdfPage, searchTerm, matcher.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    startIndex = match.Index >= SearchOtherExtensions.IndexBoundary ? match.Index - SearchOtherExtensions.IndexBoundary : 0;
                                    endIndex = (pdfPage.Length >= match.Index + match.Length + SearchOtherExtensions.IndexBoundary) ? match.Index + match.Length + SearchOtherExtensions.IndexBoundary : pdfPage.Length;
                                    string matchLine = pdfPage.Substring(startIndex, endIndex - startIndex);
                                    Match searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegexOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
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

            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
            {
                matchedLines.Clear();
            }

            return matchedLines;
        }
    }
}
