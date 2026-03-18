// <copyright file="OdpSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

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
using SharpCompress.Archives;
using SharpCompress.Common;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SearcherLibrary.FileExtensions
{
    /// <summary>
    /// Class to search ODP files.
    /// </summary>
    public class OdpSearchHandler : FileSearchHandler
    {
        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => [".ODP"];

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .ODP files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];
            int matchCounter = 0;
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                IArchive? archive = null;

                if (fileName.ToUpper().EndsWith(".ODP") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
                {
                    archive = ArchiveFactory.OpenArchive(fileName);
                }

                if (archive != null)
                {
                    foreach (IArchiveEntry entry in archive.Entries)
                    {
                        if (!entry.IsDirectory && entry.Key == "content.xml")
                        {
                            // Ignore symbolic links as these are captured by the original target.
                            if (string.IsNullOrWhiteSpace(entry.LinkTarget) && !entry.Key.Any(c => DisallowedCharactersByOperatingSystem.Any(dc => dc == c)))
                            {
                                try
                                {
                                    entry.WriteToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                                    string fullFilePath = System.IO.Path.Combine(tempDirPath, entry.Key.Replace(@"/", @"\"));
                                    StringBuilder presentationAllText = new();

                                    foreach (XElement element in XDocument.Load(fullFilePath, LoadOptions.None).Descendants().Where(d => d.Name.LocalName == "page"))
                                    {
                                        string slideName = element.Attributes().Where(sn => sn.Name.LocalName == "name").Select(sn => sn.Value).FirstOrDefault() ?? string.Empty;

                                        // Search based on keyword "Slide", not the resources translation.
                                        if (int.TryParse(slideName.Replace("Slide", string.Empty), out int slideNumber))
                                        {
                                            string slideAllText = string.Join(Environment.NewLine, element.Descendants().Where(sc => sc.Name.LocalName == "span").Select(sc => sc.Value));

                                            if (!matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                                            {
                                                foreach (string searchTerm in searchTerms)
                                                {
                                                    MatchCollection matches = Regex.Matches(slideAllText, searchTerm, matcher.RegularExpressionOptions);            // Use this match for getting the locations of the match.

                                                    if (matches.Count > 0)
                                                    {
                                                        foreach (Match match in matches.Cast<Match>())
                                                        {
                                                            int startIndex = match.Index >= IndexBoundary ? match.Index - IndexBoundary : 0;
                                                            int endIndex = (slideAllText.Length >= match.Index + match.Length + IndexBoundary) ? match.Index + match.Length + IndexBoundary : slideAllText.Length;
                                                            string matchLine = slideAllText[startIndex..endIndex];

                                                            while (matchLine.StartsWith('\r') || matchLine.StartsWith('\n'))
                                                            {
                                                                matchLine = matchLine[1..];         // Remove lines starting with the newline character.
                                                            }

                                                            while ((matchLine.EndsWith('\r') || matchLine.EndsWith('\n')) && matchLine.Length > 2)
                                                            {
                                                                matchLine = matchLine[..^1];       // Remove lines ending with the newline character.
                                                            }

                                                            Match searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegularExpressionOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                                            matchedLines.Add(new MatchedLine
                                                            {
                                                                MatchId = matchCounter++,
                                                                Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Slide, slideNumber.ToString(), matchLine),
                                                                SearchTerm = searchTerm,
                                                                FileName = fileName,
                                                                LineNumber = 1,
                                                                StartIndex = searchMatch.Index + Resources.Strings.Slide.Length + 3 + slideNumber.ToString().Length,
                                                                Length = searchMatch.Length
                                                            });
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                presentationAllText.AppendLine(slideAllText);
                                            }
                                        }
                                    }

                                    if (matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                                    {
                                        matchedLines = matcher.GetMatch([string.Join(Environment.NewLine, presentationAllText.ToString())], searchTerms, Strings.Slide);
                                    }
                                }
                                catch (PathTooLongException ptlex)
                                {
                                    throw new PathTooLongException(string.Format("{0} {1} {2} {3} - {4}", Resources.Strings.ErrorAccessingEntry, entry.Key, Resources.Strings.InArchive, fileName, ptlex.Message));
                                }
                            }
                        }
                    }

                    if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        matchedLines.Clear();
                    }

                    archive.Dispose();
                }
            }
            finally
            {
                RemoveTempDirectory(tempDirPath);
            }

            return matchedLines;
        }

        #endregion Public Methods
    }
}
