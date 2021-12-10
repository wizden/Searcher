// <copyright file="FileSearchHandlerFactory.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary
{
    /*
     * Searcher - Utility to search file content
     * Copyright (C) 2018  Dennis Joseph
     * 
     * This file is part of Searcher.

     * Searcher is free software: you can redistribute it and/or modify
     * it under the terms of the GNU General Public License as published by
     * the Free Software Foundation, either version 3 of the License, or
     * (at your option) any later version.
     * 
     * Searcher is distributed in the hope that it will be useful,
     * but WITHOUT ANY WARRANTY; without even the implied warranty of
     * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     * GNU General Public License for more details.
     * 
     * You should have received a copy of the GNU General Public License
     * along with Searcher.  If not, see <https://www.gnu.org/licenses/>.
     */

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using Resources;
    using FileExtensions;

    /// <summary>
    /// Factory to determine which handler can process the search on a file.
    /// </summary>
    public class FileSearchHandlerFactory
    {
        #region Private Fields

        /// <summary>
        /// Private store for the list of available search handlers.
        /// </summary>
        private static readonly Lazy<Dictionary<string, IFileSearchHandler>> LazySearchHandlers = new Lazy<Dictionary<string, IFileSearchHandler>>(
            () =>
        {
            var retVal = new Dictionary<string, IFileSearchHandler>();

            foreach (var fshType in Assembly.GetAssembly(typeof(FileSearchHandler)).GetTypes().Where(t => t.IsSubclassOf(typeof(FileSearchHandler))))
            {
                foreach (var fileType in fshType?.GetProperties().Where(p => p.Name == "Extensions" && p.PropertyType == typeof(List<string>))
                   .Select(p => p.GetValue(null, null)).Cast<List<string>>().FirstOrDefault())
                {
                    retVal.Add(fileType.ToUpper(), (IFileSearchHandler)Activator.CreateInstance(fshType));
                }
            }

            return retVal;
        },
         LazyThreadSafetyMode.ExecutionAndPublication);

        #endregion Private Fields

        #region Private Constructors

        /// <summary>
        /// Prevents a default instance of the <see cref="FileSearchHandlerFactory" /> class from being created.
        /// </summary>
        private FileSearchHandlerFactory()
        {
        }

        #endregion Private Constructors

        #region Public Methods

        /// <summary>
        /// Search for content in the specified file based on the search criteria.
        /// </summary>
        /// <param name="fileName">The file name for which the content is searched for.</param>
        /// <param name="searchTerms">The terms to search in the file.</param>
        /// <param name="matcher">Criteria for matching terms in file.</param>
        /// <returns>List of matches that are found in the file.</returns>
        public static List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            var handler = GetSearchHandler(fileName);
			
			try
            {
                return handler.Search(fileName, searchTerms, matcher);
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("{0} {1}. {2}", Strings.ErrorAccessingFile, fileName, ex.Message));
            }
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Determine the search handler for the file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>The <see cref="IFileSearchHandler" /> search handler to search file content.</returns>
        private static IFileSearchHandler GetSearchHandler(string fileName)
        {
            var searchHandler = LazySearchHandlers.Value.FirstOrDefault(sh => sh.Key == Path.GetExtension(fileName).ToUpper()).Value;

            if (searchHandler != null)
            {
                return searchHandler;
            }

            return new FileSearchHandler();
        }

        #endregion Private Methods
    }
}