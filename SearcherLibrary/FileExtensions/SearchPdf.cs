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
                        string pdfPage = Regex.Replace(PdfTextExtractor.GetTextFromPage(reader, pageCounter), @"\r\n?|\n", " ").Trim();

                        foreach (string searchTerm in searchTerms)
                        {
                            MatchCollection matches = Regex.Matches(pdfPage, searchTerm, matcher.RegexOptions);            // Use this match for getting the locations of the match.
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    int startIndex = match.Index >= IndexBoundary ? match.Index - IndexBoundary : 0;
                                    int endIndex = (pdfPage.Length >= match.Index + match.Length + IndexBoundary) ? match.Index + match.Length + IndexBoundary : pdfPage.Length;
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

                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        matchedLines.Clear();
                    }
                }

                reader.Close();
            }

            return matchedLines;
        }
    }
}
