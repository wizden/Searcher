namespace SearcherLibrary.Tests
{
    using System.Text.RegularExpressions;

    public class ArchiveTests
    {
        #region Public Methods

        /// <summary>
        /// Class to generate test data that gets list of files to test. Will need update as new file type handlers for search are added.
        /// </summary>
        public class FileNamesTestData : TheoryData<string>
        {
            public FileNamesTestData()
            {
                Add(TestHelpers.GetFilePathForTestFile("7z.7z"));
                Add(TestHelpers.GetFilePathForTestFile("Gz.gz"));
                ////Add(TestHelpers.GetFilePathForExtension("Rar.rar"));
                Add(TestHelpers.GetFilePathForTestFile("Tar.tar"));
                Add(TestHelpers.GetFilePathForTestFile("Zip.zip"));
            }
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_CaseInsensitive_Matches22(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(22, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_CaseSensitive_Matches13(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["The"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(12, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_Regex_CaseInsensitive_Matches10(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["th.*qu"], new Matcher { RegularExpressionOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(10, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_Regex_CaseInsensitive_Multiline_Matches114(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["e(.|\n)*?o"], new Matcher { RegularExpressionOptions = RegexOptions.Multiline | RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(114, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_Regex_CaseSensitive_Matches10(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["Th.*qu"], new Matcher { RegularExpressionOptions = RegexOptions.Singleline });

            // Assert
            Assert.Equal(10, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_TwoWords_CaseInsensitive_Matches32(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.IgnoreCase });

            // Assert
            Assert.Equal(32, matchedLines.Count);
        }

        [Theory]
        [ClassData(typeof(FileNamesTestData))]
        public void SearchText_TwoWords_CaseInsensitive_Matches20(string filePath)
        {
            // Arrange / Act
            var matchedLines = FileSearchHandlerFactory.Search(filePath, ["the", "quick"], new Matcher { RegularExpressionOptions = RegexOptions.None });

            // Assert
            Assert.Equal(20, matchedLines.Count);
        }

        #endregion Public Methods
    }
}
