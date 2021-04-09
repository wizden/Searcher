// <copyright file="SearchArchive.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using SharpCompress.Common;
    using SharpCompress.Readers;

    /// <summary>
    /// Class to search archive file.
    /// </summary>
    internal class SearchArchive : SearchOtherExtensions
    {
        /// <summary>
        /// Search for matches in zipped files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInZip(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string tempDirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName + TempExtractDirectoryName);

            try
            {
                Directory.CreateDirectory(tempDirPath);
                SharpCompress.Archives.IArchive archive = null;

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
                    matchedLines = this.GetMatchedLinesInZipArchive(fileName, searchTerms, tempDirPath, archive, matcher);
                    archive.Dispose();
                }

                this.RemoveTempDirectory(tempDirPath);
                return matchedLines;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Decompress a GZIP file stream.
        /// </summary>
        /// <param name="fileName">The GZIP file name.</param>
        /// <param name="searchTerms">The terms to search for.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>List of matched lines for GZIP file contents.</returns>
        private List<MatchedLine> DecompressGZipStream(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            string newFileName = string.Empty;

            FileInfo fileToDecompress = new FileInfo(fileName);
            using (FileStream originalFileStream = fileToDecompress.OpenRead())
            {
                string currentFileName = fileToDecompress.FullName;
                newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);

                using (FileStream decompressedFileStream = File.Create(newFileName))
                {
                    using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(newFileName))
            {
                matchedLines.AddRange(matcher.GetMatch(newFileName, searchTerms));
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
        private List<MatchedLine> GetMatchedLinesInZipArchive(string fileName, IEnumerable<string> searchTerms, string tempDirPath, SharpCompress.Archives.IArchive archive, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

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
                                string fullFilePath = System.IO.Path.Combine(tempDirPath, reader.Entry.Key.Replace(@"/", @"\"));
                                matchedLines.AddRange(matcher.GetMatch(fullFilePath, searchTerms));

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
                    matchedLines.Clear();
                }
            }
            catch (ArgumentNullException ane)
            {
                if (ane.Message.Contains("Value cannot be null") && fileName.EndsWith(".gz", StringComparison.OrdinalIgnoreCase))
                {
                    matchedLines = this.DecompressGZipStream(fileName, searchTerms, matcher);
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

            return matchedLines;
        }
    }
}
