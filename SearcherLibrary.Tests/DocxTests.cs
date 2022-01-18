using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class DocxTests
    {
        #region Private Fields

        private static string rootDirectory = "FilesToTest";
        string filePath = Path.Combine(Directory.GetParent(new Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath).FullName, rootDirectory, "Docx.docx");

        #endregion Private Fields

        #region Public Methods

        [Fact]
        public void SearchText_CaseInsensitive_MatchesTwo()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.True(matchedLines.Count == 2);
        }

        [Fact]
        public void SearchText_CaseSensitive_MatchesOne()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.None });
            Assert.True(matchedLines.Count == 1);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_MatchesOne()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.True(matchedLines.Count == 1);
        }

        [Fact]
        public void SearchText_Regex_CaseInsensitive_Multiline_MatchesThree()
        {
            // TODO: Determine if there is a way to give multi-line regex with sensible return value to client.
            //var test = File.ReadAllText(filePath);
            //var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            //Assert.True(matchedLines.Count == 3);
        }

        [Fact]
        public void SearchText_Regex_CaseSensitive_MatchesOne()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.Singleline });
            Assert.True(matchedLines.Count == 1);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase });
            Assert.True(matchedLines.Count == 3);
        }

        [Fact]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo()
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegexOptions = System.Text.RegularExpressions.RegexOptions.None });
            Assert.True(matchedLines.Count == 2);
        }

        #endregion Public Methods
    }
}
