// <copyright file="FileSearchResult.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
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
