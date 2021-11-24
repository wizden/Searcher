// <copyright file="FileSearchHandlerFactory.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using SearcherLibrary.FileExtensions;

namespace SearcherLibrary
{
    /// <summary>
    ///     Factory to determine which handler can process the search on a file.
    /// </summary>
    public class FileSearchHandlerFactory
    {

        /// <summary>
        ///     Private store for the list of available search handlers.
        /// </summary>
        private static readonly Lazy<Dictionary<String, IFileSearchHandler>> _lazySearchHandlers = new Lazy<Dictionary<String, IFileSearchHandler>>(() =>
        {
            Console.WriteLine("Thread id: " + Thread.CurrentThread.ManagedThreadId);
            var retVal = new Dictionary<String, IFileSearchHandler>();

            foreach (var fshType in Assembly.GetAssembly(typeof(FileSearchHandler)).GetTypes().Where(t => t.IsSubclassOf(typeof(FileSearchHandler))))
            {
                foreach (var fileType in fshType.GetProperties().
                                                 Where(p => p.Name == "Extensions" && p.PropertyType == typeof(List<String>)).
                                                 Select(p => p.GetValue(null, null)).
                                                 Cast<List<String>>().
                                                 FirstOrDefault())
                {
                    retVal.Add(fileType.ToUpper(), (IFileSearchHandler) Activator.CreateInstance(fshType));
                }
            }

            return retVal;
        });

        private FileSearchHandlerFactory() { }

        public static List<MatchedLine> Search(String fileName, IEnumerable<String> searchTerms, Matcher matcher)
        {
            var handler = GetSearchHandler(fileName);
            return handler.Search(fileName, searchTerms, matcher);
        }

        /// <summary>
        ///     Determine the search handler for the file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>The <see cref="IFileSearchHandler" /> search handler to search file content.</returns>
        private static IFileSearchHandler GetSearchHandler(String fileName)
        {
            var searchHandler = _lazySearchHandlers.Value.FirstOrDefault(sh => sh.Key == Path.GetExtension(fileName).ToUpper()).Value;

            if (searchHandler != null)
            {
                return searchHandler;
            }

            return new FileSearchHandler();
        }

    }
}
