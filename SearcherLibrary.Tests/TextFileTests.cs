using System.Reflection;
using System.Text.RegularExpressions;

namespace SearcherLibrary.Tests
{
    public class TextFileTests
    {
        #region Private Fields

        private static readonly string rootDirectory = "FilesToTest";
        private readonly string filePath = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.Parent!.Parent!.Parent!.FullName, rootDirectory, "Txt.txt");

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
        public void SearchText_LongText_CaseInsensitive_MatchesOne()
        {
            // Arrange / Act
			var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "aa" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            
			// Assert
			Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Exclude_Result_ALL_NotFound()
        {
            // Arrange / Act
			var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quicker" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase, AllMatchesInFile = true });
            
			// Assert
			Assert.Empty(matchedLines);
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
        public void SearchText_Regex_CaseSensitive_Multiline_MatchesOne()
        {
            // Arrange / Act
			var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The(.|\n)*?fox" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            
			// Assert
			Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_RegexFailure()
        {
        	// Assert
			Assert.Throws<Exception>(() =>
            {
                FileSearchHandlerFactory.Search(filePath, new string[] { @"The [ fox" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            });
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            // Arrange / Act
			var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            
			// Assert
			Assert.Equal(79, matchedLines.Count);
            Assert.StartsWith("Line 1", matchedLines[0].Content);
            Assert.StartsWith("Line 2", matchedLines[1].Content);
            Assert.StartsWith("Line 2", matchedLines[2].Content);
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

        [Fact]
        public void SearchText_Cancellation_Succeeds()
        {
            // Arrange
            CancellationTokenSource cts = new();
            cts.Cancel();

            // Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "aa" }, 
                new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase, 
                    CancellationTokenSource = cts });
            
			// Assert
			Assert.Empty(matchedLines);
        }

        #endregion Public Methods
    }
}
