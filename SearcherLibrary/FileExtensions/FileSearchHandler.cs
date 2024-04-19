// <copyright file="FileSearchHandler.cs" company="dennjose">
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

    using SearcherLibrary.Resources;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Class to search files that do not have any specific handler.
    /// </summary>
    public class FileSearchHandler : IFileSearchHandler
    {
        #region Internal Fields

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        internal const int IndexBoundary = 50;

        /// <summary>
        /// Private variable to hold the name of the temporary extraction directory.
        /// </summary>
        internal const string TempExtractDirectoryName = "→_extract_←";

        #endregion Internal Fields

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

        /// <summary>
        /// List of characters not allowed for the file system.
        /// </summary>
        private static List<char>? disallowedCharactersByOperatingSystem;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets the list of extensions that can be processed by this handler. Empty is default and handles all non-specified extensions.
        /// </summary>
        public static List<string> Extensions => [];

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets the list of characters not allowed by the operating system.
        /// </summary>
        internal static List<char> DisallowedCharactersByOperatingSystem
        {
            get
            {
                if (disallowedCharactersByOperatingSystem == null)
                {
                    disallowedCharactersByOperatingSystem = [];

                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        // Based on: https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file. Excluding "\" and "/" as these get used for IEntry paths.
                        disallowedCharactersByOperatingSystem = ['<', '>', ':', '|', '?', '*'];
                    }
                }

                return disallowedCharactersByOperatingSystem;
            }
        }

        #endregion Internal Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in files not being handled by other extensions.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public virtual List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            var matchedLines = new List<MatchedLine>();
            var matchCounter = 0;
            var lineCounter = 0;
            Match tempMatchObj;

            // Get length of line keyword based on length in language.
            // 7 - Based on length of "Line {0}:\t{1}".
            //                         12345   6 7
            var lengthOfLineKeywordPlus3 = Strings.Line.Length + 3;

            if (matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
            {
                return GetMultiLineRegexMatches(fileName, searchTerms, matcher);
            }
            else
            {
                foreach (var line in File.ReadAllLines(fileName))
                {
                    lineCounter++;
                    string? searchLine = line.Trim();

                    if (matcher.CancellationTokenSource.IsCancellationRequested)
                    {
                        break;
                    }

                    foreach (var searchTerm in searchTerms)
                    {
                        try
                        {
                            var matches = Regex.Matches(searchLine, searchTerm, matcher.RegularExpressionOptions);

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches.Cast<Match>())
                                {
                                    if (searchLine.Length >= MaxStringLengthCheck)
                                    {
                                        // If the lines are exceesively long, handle accordingly.
                                        tempMatchObj = match;
                                        int lineToDisplayStart = match.Index >= MaxStringLengthDisplayIndexStart ? match.Index - MaxStringLengthDisplayIndexStart : match.Index;
                                        int lineToDisplayEnd = searchLine.Length - (match.Index + match.Length) >= MaxStringLengthDisplayIndexEnd
                                                   ? MaxStringLengthDisplayIndexEnd
                                                   : searchLine.Length - (match.Index + match.Length);
                                        string? tempSearchLine = searchLine.Substring(lineToDisplayStart, lineToDisplayEnd);
                                        tempMatchObj = Regex.Match(tempSearchLine, searchTerm, matcher.RegularExpressionOptions);

                                        matchedLines.Add(new MatchedLine
                                        {
                                            MatchId = matchCounter++,
                                            Content = string.Format("{0} {1}:\t{2}", Strings.Line, lineCounter, tempSearchLine),
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
                                            Content = string.Format("{0} {1}:\t{2}", Strings.Line, lineCounter, searchLine),
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
                            matcher.CancellationTokenSource.Cancel();
                            throw new ArgumentException(aex.Message + Strings.RegexFailureSearchCancelled);
                        }
                    }
                }
            }

            return matchedLines;
        }

        private static List<MatchedLine> GetMultiLineRegexMatches(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];
            foreach (var searchTerm in searchTerms)
            {
                string allText = File.ReadAllText(fileName);
                var matches = Regex.Matches(allText, searchTerm, matcher.RegularExpressionOptions);

                if (matches.Count > 0)
                {
                    int matchCounter = 0;
                    var lengthOfLineKeywordPlus3 = Strings.Line.Length + 3;

                    foreach (Match match in matches.Cast<Match>())
                    {
                        int lineNumber = allText[..match.Index].Count(nl => nl == '\n') + 1;

                        matchedLines.Add(new MatchedLine
                        {
                            MatchId = matchCounter++,
                            Content = string.Format("{0} {1}:\t{2}", Strings.Line, lineNumber, match.Value),
                            SearchTerm = searchTerm,
                            LineNumber = lineNumber,
                            StartIndex = lengthOfLineKeywordPlus3 + lineNumber.ToString().Length,
                            Length = match.Length
                        });
                    }
                }
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Remove the temporarily created directory.
        /// </summary>
        /// <param name="tempDirPath">The temporarily created directory containing archive contents.</param>
        internal static void RemoveTempDirectory(string tempDirPath)
        {
            IOException? fileAccessException;
            int counter = 0;    // If unable to delete after 10 attempts, get out instead of staying stuck.

            do
            {
                fileAccessException = null;

                try
                {
                    // Clean up to delete the temporarily created directory. If files are still in use, an exception will be thrown.
                    Directory.Delete(tempDirPath, true);
                }
                catch (IOException ioe)
                {
                    fileAccessException = ioe;
                    Thread.Sleep(counter);     // Sleep for tiny increments of time, while waiting for file to be released (max 45 milliseconds).
                    counter++;
                }
            } while (fileAccessException != null
                && fileAccessException.Message.Contains("The process cannot access the file", StringComparison.CurrentCultureIgnoreCase) && counter < 10);
        }

        #endregion Internal Methods
    }
}
