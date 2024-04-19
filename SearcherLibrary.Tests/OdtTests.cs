namespace SearcherLibrary.Tests
{
    using System.Text.RegularExpressions;

    public class OdtTests
    {
        #region Private Fields

        private readonly string filePath = TestHelpers.GetFilePathForTestFile("Odt.odt");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["The"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Single(matchedLines);
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
        public void SearchText_Regex_CaseInsensitive_Multiline_Matches80()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["e(.|\n)*?o"], new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(80, matchedLines.Count);
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
        public void SearchText_LongText_CaseInsensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["aa"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        //////[Fact]
        //////public void ShouldThrowExceptionOn_FileLockReleaseFailure()
        //////{
        //////    var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
        //////    Assert.Equal(2, matchedLines.Count);
        //////}


        #endregion Public Methods
    }
}
