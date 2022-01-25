using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class EmailTests
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
            string rootPath = Path.Combine(Directory.GetParent(new Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath).FullName, rootDirectory);

            return new List<object[]>()
            {
                new object[]{  Path.Combine(rootPath, "Eml.eml") },
                new object[]{  Path.Combine(rootPath, "Msg.msg") },
                new object[]{  Path.Combine(rootPath, "Oft.oft") }
            };
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_CaseInsensitive_MatchesTwo(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.Equal(2, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_CaseSensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.None });
            Assert.Single(matchedLines);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseInsensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_Regex_CaseSensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.Singleline });
            Assert.Single(matchedLines);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
        }

        [Theory]
        [MemberData(nameof(GetFileNames))]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = System.Text.RegularExpressions.RegexOptions.None });
            Assert.Equal(2, matchedLines.Count);
        }

        #endregion Public Methods
    }
}
