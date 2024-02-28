using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class ArchiveTests
    {
        #region Private Fields

        private static string rootDirectory = "FilesToTest";

        #endregion Private Fields

        #region Public Methods

        /// <summary>
        /// Static method to get list of files to test. Will need update as new file type handlers for search are added.
        /// </summary>
        /// <returns>List of file names as an object array expected by <see cref="MemberDataAttribute"/>.</returns>
        public static List<object[]> GetFileNames()
        {
            string rootPath = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location).Parent.Parent.Parent.FullName, rootDirectory);

            return new List<object[]>()
            {
                new object[]{  Path.Combine(rootPath, "7z.7z") },
                new object[]{  Path.Combine(rootPath, "Gz.gz") },
                ////new object[]{  Path.Combine(rootPath, "Rar.rar") },
                new object[]{  Path.Combine(rootPath, "Tar.tar") },
                new object[]{  Path.Combine(rootPath, "Zip.zip") },
            };
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_CaseInsensitive_Matches22(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(22, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_CaseSensitive_Matches13(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(12, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseInsensitive_Matches10(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });
            Assert.Equal(10, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseInsensitive_Multiline_Matches114(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            Assert.Equal(114, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseSensitive_Matches10(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });
            Assert.Equal(10, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_TwoWords_CaseInsensitive_Matches32(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(32, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_TwoWords_CaseInsensitive_Matches20(string filePath)
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(20, matchedLines.Count);
        }

        #endregion Public Methods
    }
}
