﻿namespace SearcherLibrary.Tests
{
    using System.Text.RegularExpressions;

    public class OdpTests
    {
        #region Private Fields

        private readonly string filePath = TestHelpers.GetFilePathForTestFile("Odp.odp");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesThree()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["The"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["th.*qu"], new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["e(.|\n)*?o"], new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(3, matchedLines.Count);
            Assert.StartsWith("Slide 1", matchedLines[0].Content);
            Assert.StartsWith("Slide 2", matchedLines[1].Content);
            Assert.StartsWith("Slide 3", matchedLines[2].Content);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["Th.*qu"], new Matcher { RegularExpressionOptions = RegexOptions.Singleline });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesFour()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(4, matchedLines.Count);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_WithCancellation()
        {
            // Arrange
            CancellationTokenSource cancellationTokenSource = new();
            cancellationTokenSource.Cancel();

            // Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the"], new Matcher { CancellationTokenSource = cancellationTokenSource });

            // Assert
            Assert.Empty(matchedLines);
        }

        #endregion Public Methods
    }
}
