using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using SearcherLibrary.Resources;

namespace SearcherLibrary.FileExtensions
{
    public class PdfSearchHandler : FileSearchHandler
    {

        /// <summary>
        ///     Handles files with the .pdf extension.
        /// </summary>
        public new static List<String> Extensions => new List<String> {".PDF"};

        /// <summary>
        ///     The number of characters to display before and after the matched content index.
        /// </summary>
        internal const Int32 MaxIndexBoundary = 50;

        public override List<MatchedLine> Search(String fileName, IEnumerable<String> searchTerms, Matcher matcher)
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
                                    var searchMatch = Regex.Match(matchLine,
                                                                  searchTerm,
                                                                  matcher.
                                                                      RegexOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine
                                                     {
                                                         MatchId    = matchCounter++,
                                                         Content    = String.Format("{0} {1}:\t{2}", Strings.Page, pageCounter.ToString(), matchLine),
                                                         SearchTerm = searchTerm,
                                                         FileName   = fileName,
                                                         LineNumber = 1,
                                                         StartIndex = searchMatch.Index + Strings.Page.Length + 3 + pageCounter.ToString().Length,
                                                         Length     = searchMatch.Length
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

    }
}
