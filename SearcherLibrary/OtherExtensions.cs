// <copyright file="OtherExtensions.cs" company="dennjose">
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

    /// <summary>
    /// List of extensions that can be searched.
    /// </summary>
    public enum OtherExtensions
    {
        /// <summary>
        /// Files with .7z extension. Needs to be checked manually as an identifier cannot have a number as the first character.
        /// </summary>
        SEVENZIP,

        /// <summary>
        /// Files with .DOCM extension.
        /// </summary>
        DOCM,

        /// <summary>
        /// Files with .DOCX extension.
        /// </summary>
        DOCX,

        /// <summary>
        /// Files with .EML extension.
        /// </summary>
        EML,

        /// <summary>
        /// Files with .GZ (GZIP) extension.
        /// </summary>
        GZ,

        /// <summary>
        /// Files with .MSG extension for outlook mail files.
        /// </summary>
        MSG,

        /// <summary>
        /// Files with .ODP extension for Open Document Presentation.
        /// </summary>
        ODP,

        /// <summary>
        /// Files with .ODS extension for Open Document Spreadsheet.
        /// </summary>
        ODS,

        /// <summary>
        /// Files with .ODT extension for Open Document Text.
        /// </summary>
        ODT,

        /// <summary>
        /// Files with .OFT extension for outlook template files.
        /// </summary>
        OFT,
        
        /// <summary>
        /// Files with .PDF extension.
        /// </summary>
        PDF,

        /// <summary>
        /// Files with .RAR extension.
        /// </summary>
        RAR,

        /// <summary>
        /// Files with .PPTM extension.
        /// </summary>
        PPTM,

        /// <summary>
        /// Files with .PPTX extension.
        /// </summary>
        PPTX,

        /// <summary>
        /// Files with .TAR extension.
        /// </summary>
        TAR,

        /// <summary>
        /// Files with .XLSM extension.
        /// </summary>
        XLSM,

        /// <summary>
        /// Files with .XLSX extension.
        /// </summary>
        XLSX,

        /// <summary>
        /// Files with .ZIP extension.
        /// </summary>
        ZIP
    }
}
