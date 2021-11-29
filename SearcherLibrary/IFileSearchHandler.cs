// <copyright file="IFileSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Interface to search for text in files.
    /// </summary>
    public interface IFileSearchHandler
    {
        #region Public Methods

        /// <summary>
        /// Search for matches in file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher);

        #endregion Public Methods
    }
}
