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
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;

    /// <summary>
    /// Class to search files for matches.
    /// </summary>
    public class Matcher
    {
        #region Private Fields

        /// <summary>
        /// Private store for the maximum length in the search text.
        /// </summary>
        private const int MaxSearchTextLength = 100;

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

        /// <summary>
        /// Private field to store class that allows searching NON-ASCII files.
        /// </summary>
        private SearchOtherExtensions otherExtensions = new SearchOtherExtensions();

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="Matcher"/> class.
        /// </summary>
        /// <param name="matchWholeWord">Is the search to be performed on the word as a whole.</param>
        /// <param name="isRegexSearch">Is the search to be performed as a regular expression search.</param>
        /// <param name="isMultiLineRegex">Is the regular expression search to be performed across multiple lines.</param>
        /// <param name="allMatchesInFile">Should the search report only those files that match all search terms.</param>
        /// <param name="cancellationTokenSource">The cancellation object.</param>
        /// <param name="regexOptions">The regular expressions options object (e.g. Compiled, Multiline, IgnoreCase).</param>
        public Matcher(bool matchWholeWord = false, bool isRegexSearch = false, bool isMultiLineRegex = false, bool allMatchesInFile = false, CancellationTokenSource cancellationTokenSource = null, RegexOptions regexOptions = System.Text.RegularExpressions.RegexOptions.None)
        {
            this.MatchWholeWord = matchWholeWord;
            this.IsRegexSearch = isRegexSearch;
            this.IsMultiLineRegex = isMultiLineRegex;
            this.AllMatchesInFile = allMatchesInFile;
            this.RegexOptions = regexOptions;
            this.CancellationTokenSource = cancellationTokenSource ?? new CancellationTokenSource();
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether the search must contain all of the search terms.
        /// </summary>
        public bool AllMatchesInFile { get; set; }

        /// <summary>
        /// Gets or sets the cancellation token source object to cancel file search.
        /// </summary>
        public CancellationTokenSource CancellationTokenSource { get; set; }

        /// <summary>
        /// Gets or sets the culture to determine the resource for language.
        /// </summary>
        public CultureInfo CultureInfo 
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
        /// Gets or sets a value indicating whether the search is across multiple lines using Regex.
        /// </summary>
        public bool IsMultiLineRegex { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the search mode uses regex.
        /// </summary>
        public bool IsRegexSearch { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the search is for the whole word.
        /// </summary>
        public bool MatchWholeWord { get; set; }

        /// <summary>
        /// Gets or sets the regex options to use when searching.
        /// </summary>
        public RegexOptions RegexOptions { get; set; }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in the file based on the search terms.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search in the file.</param>
        /// <returns>List of matched lines.</returns>
        /// <exception cref="ArgumentException">Generated when regular expression failures occur.</exception>
        /// /// <exception cref="IOException">Generated on issues occurring when dealing with the files being processed.</exception>
        public List<MatchedLine> GetMatch(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines;

            try
            {
                if (this.IsNonAsciiSearch(fileName))
                {
                    matchedLines = this.FileSearchInNonASCII(fileName, searchTerms); 
                }
                else
                {
                    matchedLines = this.GetMatch(File.ReadAllLines(fileName), searchTerms);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Error accessing file {0}. {1}", fileName, ex.Message));
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in the content based on the search terms.
        /// </summary>
        /// <param name="content">The content in which to search.</param>
        /// <param name="searchTerms">The terms to search in the content.</param>
        /// <returns>List of matched lines.</returns>
        /// <exception cref="ArgumentException">Generated when regular expression failures occur.</exception>
        public List<MatchedLine> GetMatch(IEnumerable<string> content, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            if (this.IsMultiLineRegex)
            {
                matchedLines = this.GetMatchesForMultilineRegex(content, searchTerms);
            }
            else
            {
                matchedLines = this.SearchASCIIContent(content, searchTerms);
            }

            this.ClearInvalidResults(matchedLines, searchTerms);
            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Remove any results that are invalid based on the search type being ANY or ALL.
        /// </summary>
        /// <param name="matchedLines">The list of matched lines.</param>
        /// <param name="searchTerms">The list of terms that were searched.</param>
        private void ClearInvalidResults(List<MatchedLine> matchedLines, IEnumerable<string> searchTerms)
        {
            bool canShowResult = true;

            if (this.AllMatchesInFile)
            {
                List<string> test = matchedLines.Select(ml => ml.SearchTerm.ToUpper()).Distinct().ToList();
                canShowResult = matchedLines.Select(ml => ml.SearchTerm.ToUpper()).Distinct().Count() == searchTerms.Count();
            }

            if (!canShowResult)
            {
                matchedLines.Clear();
            }
        }

        /// <summary>
        /// Search the NON ASCII file for the terms and get the matches.
        /// </summary>
        /// <param name="fileName">The file to search.</param>
        /// <param name="searchTerms">The list of terms to search.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> FileSearchInNonASCII(string fileName, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> retVal = this.otherExtensions.SearchFileForMatch(fileName, searchTerms, this);
            this.ClearInvalidResults(retVal, searchTerms);
            return retVal;
        }

        /// <summary>
        /// Get the line number in the file based on the index position.
        /// </summary>
        /// <param name="fileContent">The file content.</param>
        /// <param name="indexPosition">The index position.</param>
        /// <returns>The line number for the index in the file.</returns>
        private int GetLineNumberFromIndex(string fileContent, int indexPosition)
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
        private List<MatchedLine> GetMatchesForMultilineRegex(IEnumerable<string> content, IEnumerable<string> searchTerms)
        {
            int lineToDisplayStart = 0;
            string tempSearchLine = string.Empty;

            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string allContent = string.Join(Environment.NewLine, content);
            foreach (string searchTerm in searchTerms)
            {
                if (this.CancellationTokenSource.IsCancellationRequested)
                {
                    break;
                }

                string termToSearch = searchTerm.Replace(".*", "(.|\n)*");        // Convert the .* to ((.|\n)*) for multiline regex.
                MatchCollection matches = Regex.Matches(allContent, termToSearch, this.RegexOptions);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        if (allContent.Length >= MaxStringLengthCheck)
                        {
                            // If the lines are exceesively long, handle accordingly.
                            lineToDisplayStart = this.GetLineNumberFromIndex(allContent, match.Index);

                            // 7 - Based on length of "Line {0}:\t{1}".
                            //                         12345   6 7
                            matchedLines.Add(new MatchedLine
                            {
                                Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineToDisplayStart, match.Value),
                                SearchTerm = termToSearch,
                                LineNumber = lineToDisplayStart,
                                StartIndex = lineToDisplayStart.ToString().Length + 7,
                                Length = match.Length
                            });
                        }
                        else
                        {
                            int lineNumberStart = this.GetLineNumberFromIndex(allContent, match.Index);

                            matchedLines.Add(new MatchedLine
                            {
                                Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineNumberStart.ToString(), match.Value),
                                SearchTerm = termToSearch,
                                LineNumber = lineNumberStart + 1,
                                StartIndex = 7 + lineNumberStart.ToString().Length,
                                Length = match.Length
                            });
                        }
                    }
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Search files that are not in the ASCII format.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>Boolean indicating whether the file searched is not in ASCII format.</returns>
        private bool IsNonAsciiSearch(string fileName)
        {
            bool retVal = false;
            string fileExtension = Path.GetExtension(fileName).ToUpper();

            // Determine exceptional names to search first.
            if (fileExtension.ToUpper() == ".7Z".ToUpper())
            {
                retVal = true;
            }
            else
            {
                retVal = Enum.GetNames(typeof(OtherExtensions)).Any(s => fileExtension.Contains(s.ToUpper()));
            }

            return retVal;
        }

        /// <summary>
        /// Search the contents of the IEnumerable string in the ASCII for matches.
        /// </summary>
        /// <param name="contents">The contents of the file or string</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <returns>List of matches in the file or string contents based on the search terms.</returns>
        private List<MatchedLine> SearchASCIIContent(IEnumerable<string> contents, IEnumerable<string> searchTerms)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            int lineCounter = 0;
            int lineToDisplayStart = 0;
            int lineToDisplayEnd = 0;
            Match tempMatchObj;
            string searchLine = string.Empty;
            string tempSearchLine = string.Empty;

            foreach (string line in contents)
            {
                lineCounter++;
                searchLine = line.Trim();

                if (this.CancellationTokenSource.IsCancellationRequested)
                {
                    break;
                }

                foreach (string searchTerm in searchTerms)
                {
                    try
                    {
                        MatchCollection matches = Regex.Matches(searchLine, searchTerm, this.RegexOptions);

                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                if (searchLine.Length >= MaxStringLengthCheck)
                                {
                                    // If the lines are exceesively long, handle accordingly.
                                    tempMatchObj = match;
                                    lineToDisplayStart = match.Index >= MaxStringLengthDisplayIndexStart ? match.Index - MaxStringLengthDisplayIndexStart : match.Index;
                                    lineToDisplayEnd = searchLine.Length - (match.Index + match.Length) >= MaxStringLengthDisplayIndexEnd ? MaxStringLengthDisplayIndexEnd : searchLine.Length - (match.Index + match.Length);
                                    tempSearchLine = searchLine.Substring(lineToDisplayStart, lineToDisplayEnd);
                                    tempMatchObj = Regex.Match(tempSearchLine, searchTerm, this.RegexOptions);

                                    // 7 - Based on length of "Line {0}:\t{1}".
                                    //                         12345   6 7
                                    matchedLines.Add(new MatchedLine
                                    {
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineCounter, tempSearchLine),
                                        SearchTerm = searchTerm,
                                        LineNumber = lineCounter,
                                        StartIndex = tempMatchObj.Index + 7 + lineCounter.ToString().Length,
                                        Length = tempMatchObj.Length
                                    });
                                }
                                else
                                {
                                    matchedLines.Add(new MatchedLine
                                    {
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Line, lineCounter, searchLine),
                                        SearchTerm = searchTerm,
                                        LineNumber = lineCounter,
                                        StartIndex = match.Index + 7 + lineCounter.ToString().Length,
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
