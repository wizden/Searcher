using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class TextFileTests
    {
        #region Private Fields

        private static string rootDirectory = "FilesToTest";
        string filePath = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location).Parent.Parent.Parent.FullName, rootDirectory, "Txt.txt");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesTwo()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_LongText_CaseInsensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "aa" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Exclude_Result_ALL_NotFound()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quicker" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase, AllMatchesInFile = true });
            Assert.Empty(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_Multiline_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The(.|\n)*?fox" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_RegexFailure()
        {
            Assert.Throws<Exception>(() =>
            {
                FileSearchHandlerFactory.Search(filePath, new string[] { @"The [ fox" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            });
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            Assert.Equal(79, matchedLines.Count);
            Assert.StartsWith("Line 1", matchedLines[0].Content);
            Assert.StartsWith("Line 2", matchedLines[1].Content);
            Assert.StartsWith("Line 2", matchedLines[2].Content);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_Cancellation_Succeeds()
        {
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "aa" }, 
                new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase, 
                    CancellationTokenSource = cts });
            Assert.Empty(matchedLines);
        }

        #endregion Public Methods
    }
}
