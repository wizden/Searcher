// <copyright file="SearchOdp.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;
    using SharpCompress.Common;
    using SharpCompress.Readers;

    /// <summary>
    /// Class to search OpenDocument Presentation file.
    /// </summary>
    internal class SearchOdp : SearchOtherExtensions
    {
        /// <summary>
        /// Search for matches in ODP file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInOdp(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            int matchCounter = 0;
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                SharpCompress.Archives.IArchive archive = null;

                if (fileName.ToUpper().EndsWith(".ODP") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
                {
                    archive = SharpCompress.Archives.Zip.ZipArchive.Open(fileName);
                }

                if (archive != null)
                {
                    IReader reader = archive.ExtractAllEntries();
                    while (reader.MoveToNextEntry())
                    {
                        if (!reader.Entry.IsDirectory && reader.Entry.Key == "content.xml")
                        {
                            // Ignore symbolic links as these are captured by the original target.
                            if (string.IsNullOrWhiteSpace(reader.Entry.LinkTarget) && !reader.Entry.Key.Any(c => DisallowedCharactersByOperatingSystem.Any(dc => dc == c)))
                            {
                                try
                                {
                                    reader.WriteEntryToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                                    string fullFilePath = System.IO.Path.Combine(tempDirPath, reader.Entry.Key.Replace(@"/", @"\"));

                                    foreach (XElement element in XDocument.Load(fullFilePath, LoadOptions.None).Descendants().Where(d => d.Name.LocalName == "page"))
                                    {
                                        string slideName = element.Attributes().Where(sn => sn.Name.LocalName == "name").Select(sn => sn.Value).FirstOrDefault();
                                        int slideNumber;

                                        // Search based on keyword "Slide", not the resources translation.
                                        if (int.TryParse(slideName.Replace("Slide", string.Empty), out slideNumber))
                                        {
                                            string slideAllText = string.Join(Environment.NewLine, element.Descendants().Where(sc => sc.Name.LocalName == "span").Select(sc => sc.Value));

                                            foreach (string searchTerm in searchTerms)
                                            {
                                                MatchCollection matches = Regex.Matches(slideAllText, searchTerm, matcher.RegexOptions);            // Use this match for getting the locations of the match.

                                                if (matches.Count > 0)
                                                {
                                                    foreach (Match match in matches)
                                                    {
                                                        int startIndex = match.Index >= IndexBoundary ? match.Index - IndexBoundary : 0;
                                                        int endIndex = (slideAllText.Length >= match.Index + match.Length + IndexBoundary) ? match.Index + match.Length + IndexBoundary : slideAllText.Length;
                                                        string matchLine = slideAllText.Substring(startIndex, endIndex - startIndex);

                                                        while (matchLine.StartsWith("\r") || matchLine.StartsWith("\n"))
                                                        {
                                                            matchLine = matchLine.Substring(1, matchLine.Length - 1);                       // Remove lines starting with the newline character.
                                                        }

                                                        while ((matchLine.EndsWith("\r") || matchLine.EndsWith("\n")) && matchLine.Length > 2)
                                                        {
                                                            matchLine = matchLine.Substring(0, matchLine.Length - 1);                       // Remove lines ending with the newline character.
                                                        }

                                                        Match searchMatch = Regex.Match(matchLine, searchTerm, this.RegexOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
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
                                    }
                                }
                                catch (PathTooLongException ptlex)
                                {
                                    throw new PathTooLongException(string.Format("{0} {1} {2} {3} - {4}", Resources.Strings.ErrorAccessingEntry, reader.Entry.Key, Resources.Strings.InArchive, fileName, ptlex.Message));
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                this.RemoveTempDirectory(tempDirPath);
            }

            return matchedLines;
        }
    }
}
