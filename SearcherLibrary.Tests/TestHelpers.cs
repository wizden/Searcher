namespace SearcherLibrary.Tests
{
    using System.Reflection;

    internal class TestHelpers
    {
        /// <summary>
        /// The name of the directory where the test files are stored.
        /// </summary>
        private static readonly string rootFilesToTestDirectory = "FilesToTest";

        /// <summary>
        /// Get the full path to the test file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>The full file path and name for the test file.</returns>
        protected internal static string GetFilePathForTestFile(string fileName)
        {
            return Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.Parent!.Parent!.Parent!.FullName, rootFilesToTestDirectory, fileName); ;
        }
    }
}
