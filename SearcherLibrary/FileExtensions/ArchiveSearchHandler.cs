﻿// <copyright file="ArchiveSearchHandler.cs" company="dennjose">
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

    using SharpCompress.Common;
    using SharpCompress.Readers;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;

    /// <summary>
    /// Class to search archive files.
    /// </summary>
    public class ArchiveSearchHandler : FileSearchHandler
    {
        #region Public Properties

        /// <summary>
        /// Handles files with the .PDF extension.
        /// </summary>
        public static new List<string> Extensions => [".7Z", ".GZ", ".RAR", ".TAR", ".ZIP"];

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in archive files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];
            string tempDirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);
            Directory.CreateDirectory(tempDirPath);
            SharpCompress.Archives.IArchive? archive = null;

            try
            {

                if (fileName.ToUpper().EndsWith(".GZ") && SharpCompress.Archives.GZip.GZipArchive.IsGZipFile(fileName))
                {
                    archive = SharpCompress.Archives.GZip.GZipArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".RAR") && SharpCompress.Archives.Rar.RarArchive.IsRarFile(fileName))
                {
                    archive = SharpCompress.Archives.Rar.RarArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".7Z") && SharpCompress.Archives.SevenZip.SevenZipArchive.IsSevenZipFile(fileName))
                {
                    archive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".TAR") && SharpCompress.Archives.Tar.TarArchive.IsTarFile(fileName))
                {
                    archive = SharpCompress.Archives.Tar.TarArchive.Open(fileName);
                }
                else if (fileName.ToUpper().EndsWith(".ZIP") && SharpCompress.Archives.Zip.ZipArchive.IsZipFile(fileName))
                {
                    archive = SharpCompress.Archives.Zip.ZipArchive.Open(fileName);
                }

                if (archive != null)
                {
                    matchedLines = GetMatchedLinesInZipArchive(fileName, searchTerms, tempDirPath, archive, matcher);
                    archive.Dispose();
                }
            }
            finally
            {
                // Clean up temp directory if any errors occur.
                RemoveTempDirectory(tempDirPath);
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Decompress a GZIP file stream.
        /// </summary>
        /// <param name="fileName">The GZIP file name.</param>
        /// <param name="searchTerms">The terms to search for.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>List of matched lines for GZIP file contents.</returns>
        private static List<MatchedLine> DecompressGZipStream(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];
            string newFileName = string.Empty;

            FileInfo fileToDecompress = new(fileName);
            using (FileStream originalFileStream = fileToDecompress.OpenRead())
            {
                string currentFileName = fileToDecompress.FullName;
                newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);

                using FileStream decompressedFileStream = File.Create(newFileName);
                using GZipStream decompressionStream = new(originalFileStream, CompressionMode.Decompress);
                decompressionStream.CopyTo(decompressedFileStream);
            }

            if (!string.IsNullOrWhiteSpace(newFileName))
            {
                matchedLines.AddRange(FileSearchHandlerFactory.Search(newFileName, searchTerms, matcher));
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in zipped archive files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="tempDirPath">The temporary extract directory.</param>
        /// <param name="archive">The archive to be searched.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private static List<MatchedLine> GetMatchedLinesInZipArchive(string fileName, IEnumerable<string> searchTerms, string tempDirPath, SharpCompress.Archives.IArchive archive, Matcher matcher)
        {
            List<MatchedLine> matchedLines = [];

            try
            {
                IReader reader = archive.ExtractAllEntries();
                while (reader.MoveToNextEntry())
                {
                    if (!reader.Entry.IsDirectory)
                    {
                        // Ignore symbolic links as these are captured by the original target.
                        if (string.IsNullOrWhiteSpace(reader.Entry.LinkTarget) && !reader.Entry.Key.Any(c => DisallowedCharactersByOperatingSystem.Any(dc => dc == c)))
                        {
                            try
                            {
                                reader.WriteEntryToDirectory(tempDirPath, new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                                string fullFilePath = Path.Combine(tempDirPath, reader.Entry.Key.Replace(@"/", @"\"));
                                matchedLines!.AddRange(FileSearchHandlerFactory.Search(fullFilePath, searchTerms, matcher));

                                if (matchedLines != null && matchedLines.Count > 0)
                                {
                                    // Want the exact path of the file - without the .extract part.
                                    string dirNameToDisplay = fullFilePath.Replace(TempExtractDirectoryName, string.Empty);
                                    matchedLines.Where(ml => string.IsNullOrEmpty(ml.FileName) || ml.FileName.Contains(TempExtractDirectoryName)).ToList()
                                        .ForEach(ml => ml.FileName = dirNameToDisplay);
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
                    matchedLines!.Clear();
                }
            }
            catch (ArgumentNullException ane)
            {
                if (ane.Message.Contains("Value cannot be null") && fileName.EndsWith(".gz", StringComparison.OrdinalIgnoreCase))
                {
                    matchedLines = DecompressGZipStream(fileName, searchTerms, matcher);
                }
                else if (ane.Message.Contains("String reference not set to an instance of a String."))
                {
                    throw new NotSupportedException(string.Format("{0} {1}. {2}", Resources.Strings.ErrorAccessingFile, fileName, Resources.Strings.FileEncrypted));
                }
                else
                {
                    throw;
                }
            }

            return matchedLines!;
        }

        #endregion Private Methods
    }
}
