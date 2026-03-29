// <copyright file="OdtSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

using SharpCompress.Archives;
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

using SharpCompress.Common;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SearcherLibrary.FileExtensions
{
    /// <summary>
    /// Class to search OpenDocument Text file.
    /// </summary>
    public class OdtSearchHandler : FileSearchHandler
    {
        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => [".ODT"];

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .ODT files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                IArchive? archive = null;

                if (fileName.ToUpper().EndsWith(".ODT") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
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
                                    IEnumerable<string> content = XDocument.Load(fullFilePath, LoadOptions.None).Descendants().Where(d => d.Name.LocalName == "p").Select(d => d.Value);

                                    if (!matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                                    {
                                        matchedLines = matcher.GetMatch(content, searchTerms);
                                    }
                                    else
                                    {
                                        matchedLines = matcher.GetMatch([string.Join(string.Empty, content)], searchTerms);
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
