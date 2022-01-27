using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Xunit;

namespace SearcherLibrary.Tests
{
    public class EmailTests
    {
        public class SearchText_Regex_CaseInsensitive_Multiline_DataGenerator : TheoryData<string, int>
        {
            public SearchText_Regex_CaseInsensitive_Multiline_DataGenerator()
            {
                string rootPath = Path.Combine(Directory.GetParent(new Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath).FullName, rootDirectory);
                Add(Path.Combine(rootPath, "Eml.eml"), 11);
                Add(Path.Combine(rootPath, "Msg.msg"), 11);
                Add(Path.Combine(rootPath, "Oft.oft"), 7);
            }
        }
        
        public class SearchText_DataGenerator : TheoryData<string>
        {
            public SearchText_DataGenerator()
            {
                string rootPath = Path.Combine(Directory.GetParent(new Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath).FullName, rootDirectory);
                Add(Path.Combine(rootPath, "Eml.eml"));
                Add(Path.Combine(rootPath, "Msg.msg"));
                Add(Path.Combine(rootPath, "Oft.oft"));
            }
        }

        #region Private Fields

        private static string rootDirectory = "FilesToTest";

        #endregion Private Fields

        #region Public Methods

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_CaseInsensitive_MatchesTwo(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the" }, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(2, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_CaseSensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "The" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Single(matchedLines);
        }

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_Regex_CaseInsensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });
            Assert.Single(matchedLines);
        }

        [Theory]
        [ClassData(typeof(SearchText_Regex_CaseInsensitive_Multiline_DataGenerator))]
        public void SearchText_Regex_CaseInsensitive_Multiline_Matches11(string filePath, int matchCount)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "e(.|\n)*?o" }, new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });

            try
            {
                Assert.Equal(matchCount, matchedLines.Count);
            }
            catch (Xunit.Sdk.EqualException ee)
            {
                throw new Xunit.Sdk.XunitException($"Failed for {filePath}. {ee}");
            }
            
        }

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_Regex_CaseSensitive_MatchesOne(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "Th.*qu" }, new Matcher { RegularExpressionOptions = RegexOptions.Singleline });
            Assert.Single(matchedLines);
        }

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_TwoWords_CaseInsensitive_MatchesThree(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick"}, new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });
            Assert.Equal(3, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(SearchText_DataGenerator))]
        public void SearchText_TwoWords_CaseInsensitive_MatchesTwo(string filePath)
        {
            var test = File.ReadAllText(filePath);
            var matchedLines = FileSearchHandlerFactory.Search(filePath, new string[] { "the", "quick" }, new Matcher { RegularExpressionOptions = RegexOptions.None });
            Assert.Equal(2, matchedLines.Count);
        }

        #endregion Public Methods
    }
}
