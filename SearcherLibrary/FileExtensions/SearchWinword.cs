// <copyright file="SearchWinword.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// Class to search Word files.
    /// </summary>
    internal class SearchWinword : SearchOtherExtensions
    {
        /// <summary>
        /// The maximum length of a content to check. If longer than this, split the search output result to minimise the length of results content displayed.
        /// </summary>
        private const int MaxContentLengthCheck = 200;

        /// <summary>
        /// Search for matches in DOCX files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInDocx(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            int pageNumber = 1;
            int startIndex = 0;
            int endIndex = 0;
            int matchCounter = 0;
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
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
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
                                if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    break;
                                }

                                try
                                {
                                    content = content.Trim();
                                    MatchCollection matches = Regex.Matches(content, searchTerm, matcher.RegexOptions);           // Use this match for getting the locations of the match.

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
                                            Match searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegexOptions);         // Use this match for the result highlight, based on additional characters being selected before and after the match.

                                            // Only add, if it does not already exist (Not sure how the body elements manage to bring back the same content again.
                                            if (!matchedLines.Any(ml => ml.SearchTerm == searchTerm && ml.Content == string.Format("{0} {1}:\t{2}", Resources.Strings.Page, pageNumber, matchLine)))
                                            {
                                                matchedLines.Add(new MatchedLine
                                                {
                                                    MatchId = matchCounter++,
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

                    if (matcher.RegexOptions.HasFlag(RegexOptions.Multiline))
                    {
                        // Get matches not found in above list, if using Regex, since the above code will not catch matches that span across "Run" content.
                        List<MatchedLine> allContentMatchedLines = new List<MatchedLine>();
                        foreach (string searchTerm in searchTerms)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            try
                            {
                                allContent = allContent.Trim();
                                MatchCollection matches = Regex.Matches(allContent, searchTerm, matcher.RegexOptions);           // Use this match for getting the locations of the match.

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches)
                                    {
                                        startIndex = match.Index >= SearchOtherExtensions.IndexBoundary ? this.GetLocationOfFirstWord(allContent, match.Index - SearchOtherExtensions.IndexBoundary) : 0;
                                        endIndex = (allContent.Length >= match.Index + match.Length + SearchOtherExtensions.IndexBoundary) ? this.GetLocationOfLastWord(allContent, match.Index + match.Length + SearchOtherExtensions.IndexBoundary) : allContent.Length;
                                        string matchLine = allContent.Substring(startIndex, endIndex - startIndex);
                                        Match searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegexOptions);         // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                        allContentMatchedLines.Add(new MatchedLine
                                        {
                                            MatchId = matchCounter++,
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
    }
}
