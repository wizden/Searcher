using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using SearcherLibrary.Resources;

namespace SearcherLibrary.FileExtensions
{
    public class FileSearchHandler : IFileSearchHandler
    {

        /*
         * TODO:
         * SearchArchive.cs
         * SearchExcel.cs
         * SearchOdp.cs
         * SearchOds.cs
         * SearchOdt.cs
         * SearchOutlook.cs
         * If all works, remove old mechanism
         * Formatting code tasks.
        */

        /// <summary>
        ///     Gets or sets the list of extensions that can be processed by this handler. Empty is default and handles all
        ///     non-specified extensions.
        /// </summary>
        public static List<String> Extensions => new List<String>();

        /// <summary>
        ///     The number of characters to display before and after the matched content index.
        /// </summary>
        internal const Int32 IndexBoundary = 50;

        /// <summary>
        ///     Private store for limiting display of long strings.
        /// </summary>
        private const Int32 MaxStringLengthCheck = 2000;

        /// <summary>
        ///     Private store for setting the end index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const Int32 MaxStringLengthDisplayIndexEnd = 200;

        /// <summary>
        ///     Private store for setting the start index for strings where the length exceeds MaxStringLengthCheck.
        /// </summary>
        private const Int32 MaxStringLengthDisplayIndexStart = 100;

        /// <summary>
        ///     Gets or sets the regex options to use when searching.
        /// </summary>
        public RegexOptions RegexOptions { get; set; }

        public virtual List<MatchedLine> Search(String fileName, IEnumerable<String> searchTerms, Matcher matcher)
        {
            var   matchedLines       = new List<MatchedLine>();
            var   matchCounter       = 0;
            var   lineCounter        = 0;
            var   lineToDisplayStart = 0;
            var   lineToDisplayEnd   = 0;
            Match tempMatchObj;
            var   searchLine     = String.Empty;
            var   tempSearchLine = String.Empty;

            // Get length of line keyword based on length in language.
            // 7 - Based on length of "Line {0}:\t{1}".
            //                         12345   6 7
            var lengthOfLineKeywordPlus3 = Strings.Line.Length + 3;

            foreach (var line in File.ReadAllLines(fileName))
            {
                lineCounter++;
                searchLine = line.Trim();

                if (matcher.CancellationTokenSource.IsCancellationRequested)
                {
                    break;
                }

                foreach (var searchTerm in searchTerms)
                {
                    try
                    {
                        var matches = Regex.Matches(searchLine, searchTerm, RegexOptions);

                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                if (searchLine.Length >= MaxStringLengthCheck)
                                {
                                    // If the lines are exceesively long, handle accordingly.
                                    tempMatchObj       = match;
                                    lineToDisplayStart = match.Index >= MaxStringLengthDisplayIndexStart ? match.Index - MaxStringLengthDisplayIndexStart : match.Index;
                                    lineToDisplayEnd = searchLine.Length - (match.Index + match.Length) >= MaxStringLengthDisplayIndexEnd
                                                           ? MaxStringLengthDisplayIndexEnd
                                                           : searchLine.Length - (match.Index + match.Length);
                                    tempSearchLine = searchLine.Substring(lineToDisplayStart, lineToDisplayEnd);
                                    tempMatchObj   = Regex.Match(tempSearchLine, searchTerm, RegexOptions);

                                    matchedLines.Add(new MatchedLine
                                                     {
                                                         MatchId    = matchCounter++,
                                                         Content    = String.Format("{0} {1}:\t{2}", Strings.Line, lineCounter, tempSearchLine),
                                                         SearchTerm = searchTerm,
                                                         LineNumber = lineCounter,
                                                         StartIndex = tempMatchObj.Index + lengthOfLineKeywordPlus3 + lineCounter.ToString().Length,
                                                         Length     = tempMatchObj.Length
                                                     });
                                }
                                else
                                {
                                    matchedLines.Add(new MatchedLine
                                                     {
                                                         MatchId    = matchCounter++,
                                                         Content    = String.Format("{0} {1}:\t{2}", Strings.Line, lineCounter, searchLine),
                                                         SearchTerm = searchTerm,
                                                         LineNumber = lineCounter,
                                                         StartIndex = match.Index + lengthOfLineKeywordPlus3 + lineCounter.ToString().Length,
                                                         Length     = match.Length
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

            return matchedLines;
        }

    }
}
