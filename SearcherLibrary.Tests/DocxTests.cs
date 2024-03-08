namespace SearcherLibrary.Tests
{
    using System.Text.RegularExpressions;

    public class DocxTests
    {
        #region Private Fields

        private readonly string filePath = TestHelpers.GetFilePathForTestFile("Docx.docx");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });

            // Assert
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(3, matchedLines.Count);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(2, matchedLines.Count);
        }

        #endregion Public Methods
    }
}
