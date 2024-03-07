using System.Reflection;
using System.Text.RegularExpressions;

namespace SearcherLibrary.Tests
{
    public class OdtTests
    {
        #region Private Fields

        private readonly static string rootDirectory = "FilesToTest";
        private readonly string filePath = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.Parent!.Parent!.Parent!.FullName, rootDirectory, "Odt.odt");

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
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_Matches80()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });
            Assert.Equal(80, matchedLines.Count);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });
            Assert.Single(matchedLines);
        }

        [Fact]
        public void SearchText_LongText_CaseInsensitive_MatchesOne()
        {
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "aa" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
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

        //////[Fact]
        //////public void ShouldThrowExceptionOn_FileLockReleaseFailure()
        //////{
        //////    var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
        //////    Assert.Equal(2, matchedLines.Count);
        //////}


        #endregion Public Methods
    }
}
