// <copyright file="MatchedLine.cs" company="dennjose">
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
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Class to capture a matched line.
    /// </summary>
    public class MatchedLine
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the content of the line.
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets the length of the match instance.
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// Gets or sets the line number at which a match exists.
        /// </summary>
        public int LineNumber { get; set; }

        /// <summary>
        /// Gets or sets the search term for which the match exists.
        /// </summary>
        public string SearchTerm { get; set; }

        /// <summary>
        /// Gets or sets the start index at which a match exists.
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the file has been processed for display.
        /// </summary>
        public bool DisplayProcessed { get; set; }

        #endregion Public Properties
    }
}
