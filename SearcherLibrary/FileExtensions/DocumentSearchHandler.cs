// <copyright file="DocumentSearchHandler.cs" company="dennjose">
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
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using SearcherLibrary.Resources;

    /// <summary>
    /// Class to search OpenDocument files.
    /// </summary>
    public class DocumentSearchHandler : FileSearchHandler
    {
        #region Private Fields

        /// <summary>
        /// The maximum length of a content to check. If longer than this, split the search output result to minimise the length of results content displayed.
        /// </summary>
        private const int MaxContentLengthCheck = 200;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Handle files with the .DOCX/.DOCM extension.
        /// </summary>
        public static new List<string> Extensions => new() { ".DOCX", ".DOCM" };

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .DOCX/.DOCM files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            var pageNumber = 1;
            var startIndex = 0;
            var endIndex = 0;
            var matchCounter = 0;
            var matchedLines = new List<MatchedLine>();

            try
            {
                using (var document = WordprocessingDocument.Open(fileName, false))
                {
                    var allContentStringBuilder = new StringBuilder();
                    var body = document.MainDocumentPart?.Document.Body;
                    body?.Descendants().Where(bce => bce is Paragraph && bce.HasChildren).ToList().ForEach(bce =>
                                 {
                                     var contentText = new StringBuilder(); // Set content for each paragraph and detect page breaks (dependant on word processing application).

                                     foreach (var r in bce.Descendants().Where(r => r is Run && r.HasChildren))
                                     {
                                         if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                                         {
                                             break;
                                         }

                                         r.Descendants().ToList().ForEach(rce =>
                                                   {
                                                       if (rce is Text)
                                                       {
                                                           contentText.Append(rce.InnerText);
                                                       }

                                                       if (rce is LastRenderedPageBreak)
                                                       {
                                                           pageNumber++;
                                                       }
                                                   });
                                     }

                                     var content = contentText.ToString();
                                     allContentStringBuilder.AppendLine(content);

                                     if (!matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                                     {
                                         if (!string.IsNullOrEmpty(content))
                                         {
                                             foreach (var searchTerm in searchTerms)
                                             {
                                                 if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                                                 {
                                                     break;
                                                 }

                                                 try
                                                 {
                                                     content = content.Trim();
                                                     var matches = Regex.Matches(content, searchTerm, matcher.RegularExpressionOptions); // Use this match for getting the locations of the match.

                                                     if (matches.Count > 0)
                                                     {
                                                         foreach (Match match in matches.Cast<Match>())
                                                         {
                                                             if (content.Length > MaxContentLengthCheck)
                                                             {
                                                                 startIndex = match.Index >= IndexBoundary ? GetLocationOfFirstWord(content, match.Index - IndexBoundary) : 0;
                                                                 endIndex = content.Length >= match.Index + match.Length + IndexBoundary
                                                                                ? GetLocationOfLastWord(content, match.Index + match.Length + IndexBoundary)
                                                                                : content.Length;
                                                             }
                                                             else
                                                             {
                                                                 startIndex = 0;
                                                                 endIndex = content.Length;
                                                             }

                                                             var matchLine = content[startIndex..endIndex];
                                                             var searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegularExpressionOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.

                                                             // Only add, if it does not already exist (Not sure how the body elements manage to bring back the same content again.
                                                             if (!matchedLines.Any(ml => ml.SearchTerm == searchTerm &&
                                                                                         ml.Content == string.Format("{0} {1}:\t{2}", Strings.Page, pageNumber, matchLine)))
                                                             {
                                                                 matchedLines.Add(new MatchedLine
                                                                 {
                                                                     MatchId = matchCounter++,
                                                                     Content = string.Format("{0} {1}:\t{2}", Strings.Page, pageNumber, matchLine),
                                                                     SearchTerm = searchTerm,
                                                                     FileName = fileName,
                                                                     LineNumber = pageNumber,
                                                                     StartIndex = searchMatch.Index + Strings.Page.Length + 3 + pageNumber.ToString().Length,
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
                                             }
                                         }
                                     }
                                 });

                    if (matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                    {
                        // Get matches not found in above list, if using Regex, since the above code will not catch matches that span across "Run" content.
                        var allContentMatchedLines = new List<MatchedLine>();

                        foreach (var searchTerm in searchTerms)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            try
                            {
                                var allContent = allContentStringBuilder.ToString().Trim();
                                var matches = Regex.Matches(allContent, searchTerm, matcher.RegularExpressionOptions); // Use this match for getting the locations of the match.

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches.Cast<Match>())
                                    {
                                        startIndex = match.Index;
                                        endIndex = match.Index + match.Length;
                                        var matchLine = allContent[startIndex..endIndex];
                                        var searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegularExpressionOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                        allContentMatchedLines.Add(new MatchedLine
                                        {
                                            MatchId = matchCounter++,
                                            Content = string.Format("{0} {1}:\t{2}", Strings.Page, pageNumber, matchLine),
                                            SearchTerm = searchTerm,
                                            FileName = fileName,
                                            LineNumber = pageNumber,
                                            StartIndex = searchMatch.Index + Strings.Page.Length + 3 + pageNumber.ToString().Length,
                                            Length = searchMatch.Length
                                        });
                                    }

                                    fileName = string.Empty;
                                }
                            }
                            catch (ArgumentException aex)
                            {
                                throw new ArgumentException(aex.Message + " " + Strings.RegexFailureSearchCancelled);
                            }
                        }

                        // Only add those lines that do not already exist in matches found. Do not like the search mechanism below, but it helps to search by excluding the "Page {0}:\t{1}" content.
                        matchedLines.AddRange(allContentMatchedLines.Where(acm => !matchedLines.Any(ml => acm.Length == ml.Length 
                            && acm.Content.Contains(ml.Content[(Strings.Page.Length + 3 + ml.LineNumber.ToString().Length)..]))));
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
                    var error = fileName.Contains("~$")
                                    ? Strings.FileCorruptOrLockedByApp
                                    : Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }

                throw;
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
        private static int GetLocationOfFirstWord(string content, int startIndex)
        {
            var retVal = startIndex;

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
        private static int GetLocationOfLastWord(string content, int endIndex)
        {
            var retVal = endIndex;

            if (retVal > 0)
            {
                while (retVal < content.Length - 1 && content[retVal] != ' ')
                {
                    retVal++;
                }
            }

            return retVal;
        }

        #endregion Private Methods
    }
}
