// <copyright file="SearchPowerpoint.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;

    /// <summary>
    /// Class to search Power point files.
    /// </summary>
    internal class SearchPowerpoint : SearchOtherExtensions
    {
        /// <summary>
        /// Search for matches in PPTX files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInPptx(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            int matchCounter = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                using (PresentationDocument pptDocument = PresentationDocument.Open(fileName, false))
                {
                    string[] slideAllText = this.GetPresentationSlidesText(pptDocument.PresentationPart);
                    pptDocument.Close();

                    int startIndex = 0;
                    int endIndex = 0;

                    for (int slideCounter = 0; slideCounter < slideAllText.Length; slideCounter++)
                    {
                        if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            break;
                        }

                        foreach (string searchTerm in searchTerms)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            MatchCollection matches = Regex.Matches(slideAllText[slideCounter], searchTerm, matcher.RegexOptions);            // Use this match for getting the locations of the match.

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    startIndex = match.Index >= SearchOtherExtensions.IndexBoundary ? match.Index - SearchOtherExtensions.IndexBoundary : 0;
                                    endIndex = (slideAllText[slideCounter].Length >= match.Index + match.Length + SearchOtherExtensions.IndexBoundary) ? match.Index + match.Length + SearchOtherExtensions.IndexBoundary : slideAllText[slideCounter].Length;
                                    string matchLine = slideAllText[slideCounter].Substring(startIndex, endIndex - startIndex);

                                    while (matchLine.StartsWith("\r") || matchLine.StartsWith("\n"))
                                    {
                                        matchLine = matchLine.Substring(1, matchLine.Length - 1);                       // Remove lines starting with the newline character.
                                    }

                                    while ((matchLine.EndsWith("\r") || matchLine.EndsWith("\n")) && matchLine.Length > 2)
                                    {
                                        matchLine = matchLine.Substring(0, matchLine.Length - 1);                       // Remove lines ending with the newline character.
                                    }

                                    Match searchMatch = Regex.Match(matchLine, searchTerm, matcher.RegexOptions);          // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine
                                    {
                                        MatchId = matchCounter++,
                                        Content = string.Format("{0} {1}:\t{2}", Resources.Strings.Slide, (slideCounter + 1).ToString(), matchLine),
                                        SearchTerm = searchTerm,
                                        FileName = fileName,
                                        LineNumber = 1,
                                        StartIndex = searchMatch.Index + Resources.Strings.Slide.Length + 3 + (slideCounter + 1).ToString().Length,
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
            }
            catch (FileFormatException ffx)
            {
                if (ffx.Message == "File contains corrupted data.")
                {
                    string error = fileName.Contains("~$")
                        ? Resources.Strings.FileCorruptOrLockedByApp
                        : Resources.Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }
                else
                {
                    throw;
                }
            }

            return matchedLines;
        }

        /// <summary>
        /// Returns a string array of the text in the presentation slides.
        /// </summary>
        /// <param name="presentationPart">The presentation part of the presentation document.</param>
        /// <returns>String array of the text in the presentation slides.</returns>
        private string[] GetPresentationSlidesText(PresentationPart presentationPart)
        {
            Presentation presentation = presentationPart.Presentation;
            List<SlidePart> slideParts = presentationPart.SlideParts.ToList();
            string[] retVal = new string[slideParts.Count()];
            string relationshipId = string.Empty;
            int tempSlideNumber = 0;

            foreach (SlidePart slidePart in slideParts)
            {
                var slide = presentation.SlideIdList.Where(s => ((SlideId)s).RelationshipId == presentationPart.GetIdOfPart(slidePart)).FirstOrDefault();
                int index = presentation.SlideIdList.ToList().IndexOf(slide);
                relationshipId = ((SlideId)presentation.SlideIdList.ChildElements[index]).RelationshipId;
                List<string> titles = new List<string>();
                List<string> content = new List<string>();
                List<string> notes = new List<string>();

                slidePart.Slide.Descendants<Shape>().ToList().ForEach(shape =>
                {
                    foreach (PlaceholderShape item in shape.Descendants<PlaceholderShape>().Where(i => i.Type != null))
                    {
                        if ((item.Type.ToString().ToUpper() == "CenteredTitle".ToUpper() || item.Type.ToString().ToUpper() == "SubTitle".ToUpper() || item.Type.ToString().ToUpper() == "Title".ToUpper())
                                && shape.TextBody != null && !string.IsNullOrWhiteSpace(shape.TextBody.InnerText))
                        {
                            titles.AddRange(shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text));
                        }
                    }
                });

                content = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text).ToList();
                content.RemoveAll(s => titles.Any(t => t == s));

                if (slidePart.NotesSlidePart != null && slidePart.NotesSlidePart.NotesSlide != null && slidePart.NotesSlidePart.NotesSlide.Descendants() != null && slidePart.NotesSlidePart.NotesSlide.Descendants().Count() > 0)
                {
                    notes = slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                        .Where(s =>
                        {
                            return !int.TryParse(s.Text, out tempSlideNumber);      // Remove the record as it contains the slide number.
                        }).Select(s => s.Text).ToList();
                }

                retVal[index] = string.Join(
                    Environment.NewLine,
                    new string[] { string.Join(Environment.NewLine, titles.ToArray()), string.Join(string.Empty, content.ToArray()), string.Join(Environment.NewLine, notes.ToArray()) });
            }

            return retVal;
        }
    }
}
