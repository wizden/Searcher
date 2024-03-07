// <copyright file="PresentationSearchHandler.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
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
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using SearcherLibrary.Resources;

    /// <summary>
    /// Class to search presentation files files.
    /// </summary>
    public class PresentationSearchHandler : FileSearchHandler
    {
        #region Public Properties

        /// <summary>
        /// Handles files with the .PPTX/.PPTM extension.
        /// </summary>
        public static new List<string> Extensions => new() { ".PPTX", ".PPTM" };

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Search for matches in .PPTX/.PPTM files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            var matchCounter = 0;
            var matchedLines = new List<MatchedLine>();

            try
            {
                using var pptDocument = PresentationDocument.Open(fileName, false);
                var presentationPart = pptDocument.PresentationPart;

                if (presentationPart != null)
                {
                    var slideAllText = GetPresentationSlidesText(presentationPart);
                    var startIndex = 0;
                    var endIndex = 0;

                    if (!matcher.RegularExpressionOptions.HasFlag(RegexOptions.Multiline))
                    {
                        for (var slideCounter = 0; slideCounter < slideAllText.Length; slideCounter++)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            foreach (var searchTerm in searchTerms)
                            {
                                var matches = Regex.Matches(slideAllText[slideCounter], searchTerm, matcher.RegularExpressionOptions); // Use this match for getting the locations of the match.

                                if (matches.Count > 0)
                                {
                                    foreach (Match match in matches.Cast<Match>())
                                    {
                                        startIndex = match.Index >= FileSearchHandler.IndexBoundary ? match.Index - FileSearchHandler.IndexBoundary : 0;
                                        endIndex = slideAllText[slideCounter].Length >= match.Index + match.Length + FileSearchHandler.IndexBoundary
                                                       ? match.Index + match.Length + FileSearchHandler.IndexBoundary
                                                       : slideAllText[slideCounter].Length;
                                        var matchLine = slideAllText[slideCounter][startIndex..endIndex];

                                        while (matchLine.StartsWith("\r") || matchLine.StartsWith("\n"))
                                        {
                                            matchLine = matchLine[1..]; // Remove lines starting with the newline character.
                                        }

                                        while ((matchLine.EndsWith("\r") || matchLine.EndsWith("\n")) && matchLine.Length > 2)
                                        {
                                            matchLine = matchLine[..^1]; // Remove lines ending with the newline character.
                                        }

                                        var searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegularExpressionOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                        matchedLines.Add(new MatchedLine
                                        {
                                            MatchId = matchCounter++,
                                            Content = string.Format("{0} {1}:\t{2}", Strings.Slide, (slideCounter + 1).ToString(), matchLine),
                                            SearchTerm = searchTerm,
                                            FileName = fileName,
                                            LineNumber = 1,
                                            StartIndex = searchMatch.Index + Strings.Slide.Length + 3 + (slideCounter + 1).ToString().Length,
                                            Length = searchMatch.Length
                                        });
                                    }
                                }
                            }
                        }

                        if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            matchedLines.Clear();
                        }
                    }
                    else
                    {
                        matchedLines = matcher.GetMatch(new string[] { string.Join(Environment.NewLine, slideAllText) }, searchTerms, Strings.Slide);
                    }
                }
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    var error = fileName.Contains("~$")
                                    ? Strings.FileCorruptOrLockedByApp
                                    : Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }

                throw;
            }

            return matchedLines;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Returns a string array of the text in the presentation slides.
        /// </summary>
        /// <param name="presentationPart">The presentation part of the presentation document.</param>
        /// <returns>string array of the text in the presentation slides.</returns>
        private static string[] GetPresentationSlidesText(PresentationPart presentationPart)
        {
            var presentation    = presentationPart.Presentation;
            var slideParts      = presentationPart.SlideParts.ToList();
            var retVal          = new string[slideParts.Count];
            var relationshipId  = string.Empty;
            var tempSlideNumber = 0;

            if (presentation != null && presentation.SlideIdList != null)
            {
                foreach (var slidePart in slideParts)
                {
                    var slide = presentation.SlideIdList.Where(s => ((SlideId)s).RelationshipId == presentationPart.GetIdOfPart(slidePart)).FirstOrDefault();

                    if (slide != null)
                    {
                        var index = presentation.SlideIdList.ToList().IndexOf(slide);
                        relationshipId = ((SlideId)presentation.SlideIdList.ChildElements[index]).RelationshipId;
                        var titles = new List<string>();
                        var content = new List<string>();
                        var notes = new List<string>();

                        slidePart.Slide.Descendants<Shape>().ToList().ForEach(shape =>
                                          {
                                              foreach (var item in shape.Descendants<PlaceholderShape>().Where(i => i.Type != null))
                                              {
                                                  if ((item.Type!.ToString()!.ToUpper() == "CenteredTitle".ToUpper() ||
                                                       item.Type.ToString()!.ToUpper() == "SubTitle".ToUpper() ||
                                                       item.Type.ToString()!.ToUpper() == "Title".ToUpper()) &&
                                                      shape.TextBody != null &&
                                                      !string.IsNullOrWhiteSpace(shape.TextBody.InnerText))
                                                  {
                                                      titles.AddRange(shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text));
                                                  }
                                              }
                                          });

                        content = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text).ToList();
                        content.RemoveAll(s => titles.Any(t => t == s));

                        if (slidePart.NotesSlidePart != null &&
                            slidePart.NotesSlidePart.NotesSlide != null &&
                            slidePart.NotesSlidePart.NotesSlide.Descendants() != null &&
                            slidePart.NotesSlidePart.NotesSlide.Descendants().Any())
                        {
                            notes = slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Where(s =>
                                                    {
                                                        return !int.TryParse(s.Text, out tempSlideNumber); // Remove the record as it contains the slide number.
                                                    }).Select(s => s.Text).ToList();
                        }

                        retVal[index] = string.Join(string.Empty, string.Join(Environment.NewLine, titles.ToArray()), string.Join(string.Empty, content.ToArray()), string.Join(Environment.NewLine, notes.ToArray()));
                    }
                }
            }
            return retVal;
        }

        #endregion Private Methods
    }
}
