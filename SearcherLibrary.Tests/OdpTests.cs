using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class OdpTests
    {
        #region Private Fields

        private static string rootDirectory = "FilesToTest";
        string filePath = Path.Combine(Directory.GetParent(new Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath).FullName, rootDirectory, "Odp.odp");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesThree()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesTwo()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
            Assert.StartsWith("Slide 1", matchedLines[0].Content);
            Assert.StartsWith("Slide 2", matchedLines[1].Content);
            Assert.StartsWith("Slide 3", matchedLines[2].Content);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesFour()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(4, matchedLines.Count);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_WithCancellation()
        {
            var test = File.ReadAllText(filePath);
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            cancellationTokenSource.Cancel();
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { CancellationTokenSource = cancellationTokenSource });
            Assert.Empty(matchedLines);
        }

        #endregion Public Methods
    }
}
