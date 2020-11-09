// <copyright file="FileSearchResult.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Documents;

    /// <summary>
    /// Store the search result and start/end times of file search.
    /// </summary>
    public class FileSearchResult
    {
        /// <summary>
        /// Gets or sets the full path for the file.
        /// </summary>
        public string FileNamePath { get; set; }

        /// <summary>
        /// Gets or sets the search result matches.
        /// </summary>
        public List<Inline> SearchMatches { get; set; } = new List<Inline>();

        /// <summary>
        /// Gets or sets the search start date/time ticks.
        /// </summary>
        public long SearchStartDateTimeTicks { get; set; }
    }
}
