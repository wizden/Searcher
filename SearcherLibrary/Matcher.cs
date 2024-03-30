// <copyright file="Matcher.cs" company="dennjose">
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
    using System.Globalization;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;

    /// <summary>
    /// Class to search files for matches.
    /// </summary>
    /// <remarks>
    /// Initialises a new instance of the <see cref="Matcher"/> class.
    /// </remarks>
    /// <param name="matchWholeWord">Is the search to be performed on the word as a whole.</param>
    /// <param name="isRegexSearch">Is the search to be performed as a regular expression search.</param>
    /// <param name="allMatchesInFile">Should the search report only those files that match all search terms.</param>
    /// <param name="cancellationTokenSource">The cancellation object.</param>
    /// <param name="regexOptions">The regular expressions options object (e.g. Compiled, Multiline, IgnoreCase).</param>
    public class Matcher(bool matchWholeWord = false, bool isRegexSearch = false, bool allMatchesInFile = false, CancellationTokenSource? cancellationTokenSource = null, RegexOptions regexOptions = System.Text.RegularExpressions.RegexOptions.None)
    {
        #region Private Fields

        /// <summary>
        /// Private store for limiting display of long strings.
        /// </summary>
        private const int MaxStringLengthCheck = 2000;

        /// <summary>
        /// Private store for setting the end index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const int MaxStringLengthDisplayIndexEnd = 200;

        /// <summary>
        /// Private store for setting the start index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const int MaxStringLengthDisplayIndexStart = 100;

        #endregion Private Fields
        #region Public Constructors

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether the search must contain all of the search terms.
        /// </summary>
        public bool AllMatchesInFile { get; set; } = allMatchesInFile;

        /// <summary>
        /// Gets or sets the cancellation token source object to cancel file search.
        /// </summary>
        public CancellationTokenSource CancellationTokenSource { get; set; } = cancellationTokenSource ?? new CancellationTokenSource();

        /// <summary>
        /// Gets or sets the culture to determine the resource for language.
        /// </summary>
        public static CultureInfo CultureInfo
        {
            get
            {
                return Thread.CurrentThread.CurrentCulture;
            }

            set
            {
                Thread.CurrentThread.CurrentCulture = value;
                Thread.CurrentThread.CurrentUICulture = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the search mode uses regex.
        /// </summary>
        public bool IsRegexSearch { get; set; } = isRegexSearch;

        /// <summary>
        /// Gets or sets a value indicating whether the search is for the whole word.
        /// </summary>
        public bool MatchWholeWord { get; set; } = matchWholeWord;

        /// <summary>
        /// Gets or sets the regex options to use when searching.
        /// </summary>
        public RegexOptions RegularExpressionOptions { get; set; } = regexOptions;

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in the content based on the search terms.
        /// </summary>
        /// <param name="content">The content in which to search.</param>
        /// <param name="searchTerms">The terms to search in the content.</param>
        /// <returns>List of matched lines.</returns>
        /// <exception cref="ArgumentException">Generated when regular expression failures occur.</exception>
        public List<MatchedLine> GetMatch(IEnumerable<string> content, IEnumerable<string> searchTerms, string locationType = "Line")
        {
            List<MatchedLine> matchedLines;

            if (this.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
            {
                matchedLines = this.GetMatchesForMultilineRegex(content, searchTerms, locationType);
            }
            else
            {
                matchedLines = this.SearchASCIIContent(content, searchTerms);
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Get the line number in the file based on the index position.
        /// </summary>
        /// <param name="fileContent">The file content.</param>
        /// <param name="indexPosition">The index position.</param>
        /// <returns>The line number for the index in the file.</returns>
        private static int GetLineNumberFromIndex(string fileContent, int indexPosition)
        {
            int retVal = 1;

            for (int i = 0; i <= indexPosition - 1; i++)
            {
                if (fileContent[i] == '\n')
                {
                    retVal++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Find multi line matches using regular expressions.
        /// </summary>
        /// <param name="content">The content in which to search.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>List of matches found in content based on search terms.</returns>
        private List<MatchedLine> GetMatchesForMultilineRegex(IEnumerable<string> content, IEnumerable<string> searchTerms, string locationType = "Line")
        {
            List<MatchedLine> matchedLines = [];
            string allContent = string.Join(Environment.NewLine, content);
            foreach (string searchTerm in searchTerms)
            {
                if (this.CancellationTokenSource.IsCancellationRequested)
                {
                    break;
                }

                string termToSearch = searchTerm.Replace(".*", "(.|\n)*");        // Convert the .* to ((.|\n)*) for multiline regex.
                MatchCollection matches = Regex.Matches(allContent, termToSearch, this.RegularExpressionOptions);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches.Cast<Match>())
                    {
                        if (allContent.Length >= MaxStringLengthCheck)
                        {
                            // If the lines are exceesively long, handle accordingly.
                            int lineToDisplayStart = GetLineNumberFromIndex(allContent, match.Index);

                            // 7 - Based on length of "Line {0}:\t{1}".
                            //                         12345   6 7
                            matchedLines.Add(new MatchedLine
                            {
                                Content = string.Format("{0} {1}:\t{2}", locationType, lineToDisplayStart, match.Value),
                                SearchTerm = termToSearch,
                                LineNumber = lineToDisplayStart,
                                StartIndex = lineToDisplayStart.ToString().Length + (locationType.Length + 3),  // locationName.Length + 3 => e.g. "Line" + " :\t"
                                Length = match.Length
                            });
                        }
                        else
                        {
                            int lineNumberStart = GetLineNumberFromIndex(allContent, match.Index);

                            matchedLines.Add(new MatchedLine
                            {
                                Content = string.Format("{0} {1}:\t{2}", locationType, lineNumberStart.ToString(), match.Value),
                                SearchTerm = termToSearch,
                                LineNumber = lineNumberStart + 1,
                                StartIndex = (locationType.Length + 3) + lineNumberStart.ToString().Length,  // locationName.Length + 3 => e.g. "Line" + " :\t"
                                Length = match.Length
                            });
                        }
                    }
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search the contents of the IEnumerable string in the ASCII for matches.
        /// </summary>
        /// <param name="contents">The contents of the file or string</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>List of matches in the file or string contents based on the search terms.</returns>
        private List<MatchedLine> SearchASCIIContent(IEnumerable<string> contents, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = [];
            int matchCounter = 0;
            int lineCounter = 0;
            int lineToDisplayStart;
            int lineToDisplayEnd;
            Match tempMatchObj;

            // Get length of line keyword based on length in language.
            // 7 - Based on length of "Line {0}:\t{1}".
            //                         12345   6 7
            int lengthOfLineKeywordPlus3 = Resources.Strings.Line.Length + 3;

            foreach (string line in contents)
            {
                lineCounter++;
                string searchLine = line.Trim();

                if (this.CancellationTokenSource.IsCancellationRequested)
                {
                    break;
                }

                foreach (string searchTerm in searchTerms)
                {
                    try
                    {
                        MatchCollection matches = Regex.Matches(searchLine, searchTerm, this.RegularExpressionOptions);

                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches.Cast<Match>())
                            {
                                if (searchLine.Length >= MaxStringLengthCheck)
                                {
                                    // If the lines are exceesively long, handle accordingly.
                                    tempMatchObj = match;
                                    lineToDisplayStart = match.Index >= MaxStringLengthDisplayIndexStart ? match.Index - MaxStringLengthDisplayIndexStart : match.Index;
                                    lineToDisplayEnd = searchLine.Length - (match.Index + match.Length) >= MaxStringLengthDisplayIndexEnd ? MaxStringLengthDisplayIndexEnd : searchLine.Length - (match.Index + match.Length);
                                    string tempSearchLine = searchLine.Substring(lineToDisplayStart, lineToDisplayEnd);
                                    tempMatchObj = Regex.Match(tempSearchLine, searchTerm, this.RegularExpressionOptions);

                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineCounter, tempSearchLine),
                                        SearchTerm = searchTerm,
                                        LineNumber = lineCounter,
                                        StartIndex = tempMatchObj.Index + lengthOfLineKeywordPlus3 + lineCounter.ToString().Length,
                                        Length = tempMatchObj.Length
                                    });
                                }
                                else
                                {
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineCounter, searchLine),
                                        SearchTerm = searchTerm,
                                        LineNumber = lineCounter,
                                        StartIndex = match.Index + lengthOfLineKeywordPlus3 + lineCounter.ToString().Length,
                                        Length = match.Length
                                    });
                                }
                            }
                        }
                    }
                    catch (ArgumentException aex)
                    {
                        this.CancellationTokenSource.Cancel();
                        throw new ArgumentException(aex.Message + Resources.Strings.RegexFailureSearchCancelled);
                    }
                }
            }

            return matchedLines;
        }

        #endregion Private Methods
    }
}
