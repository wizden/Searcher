using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SearcherLibrary.Resources;

namespace SearcherLibrary.FileExtensions
{
    public class PresentationSearchHandler : FileSearchHandler
    {

        /// <summary>
        ///     Handles files with the .pdf extension.
        /// </summary>
        public new static List<String> Extensions => new List<String> {".PPTX", ".PPTM"};

        public override List<MatchedLine> Search(String fileName, IEnumerable<String> searchTerms, Matcher matcher)
        {
            var matchCounter = 0;
            var matchedLines = new List<MatchedLine>();

            try
            {
                using (var pptDocument = PresentationDocument.Open(fileName, false))
                {
                    var slideAllText = GetPresentationSlidesText(pptDocument.PresentationPart);
                    pptDocument.Close();

                    var startIndex = 0;
                    var endIndex   = 0;

                    for (var slideCounter = 0; slideCounter < slideAllText.Length; slideCounter++)
                    {
                        if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            break;
                        }

                        foreach (var searchTerm in searchTerms)
                        {
                            if (matcher.CancellationTokenSource.Token.IsCancellationRequested)
                            {
                                break;
                            }

                            var matches = Regex.Matches(slideAllText[slideCounter], searchTerm, matcher.RegexOptions); // Use this match for getting the locations of the match.

                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    startIndex = match.Index >= SearchOtherExtensions.IndexBoundary ? match.Index - SearchOtherExtensions.IndexBoundary : 0;
                                    endIndex = slideAllText[slideCounter].Length >= match.Index + match.Length + SearchOtherExtensions.IndexBoundary
                                                   ? match.Index + match.Length + SearchOtherExtensions.IndexBoundary
                                                   : slideAllText[slideCounter].Length;
                                    var matchLine = slideAllText[slideCounter].Substring(startIndex, endIndex - startIndex);

                                    while (matchLine.StartsWith("\r") || matchLine.StartsWith("\n"))
                                    {
                                        matchLine = matchLine.Substring(1, matchLine.Length - 1); // Remove lines starting with the newline character.
                                    }

                                    while ((matchLine.EndsWith("\r") || matchLine.EndsWith("\n")) && matchLine.Length > 2)
                                    {
                                        matchLine = matchLine.Substring(0, matchLine.Length - 1); // Remove lines ending with the newline character.
                                    }

                                    var searchMatch = Regex.Match(matchLine,
                                                                  searchTerm,
                                                                  matcher.
                                                                      RegexOptions); // Use this match for the result highlight, based on additional characters being selected before and after the match.
                                    matchedLines.Add(new MatchedLine
                                                     {
                                                         MatchId    = matchCounter++,
                                                         Content    = String.Format("{0} {1}:\t{2}", Strings.Slide, (slideCounter + 1).ToString(), matchLine),
                                                         SearchTerm = searchTerm,
                                                         FileName   = fileName,
                                                         LineNumber = 1,
                                                         StartIndex = searchMatch.Index + Strings.Slide.Length + 3 + (slideCounter + 1).ToString().Length,
                                                         Length     = searchMatch.Length
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
                    var error = fileName.Contains("~$")
                                    ? Strings.FileCorruptOrLockedByApp
                                    : Strings.FileCorruptOrProtected;
                    throw new IOException(error, ffx);
                }

                throw;
            }

            return matchedLines;
        }

        /// <summary>
        ///     Returns a string array of the text in the presentation slides.
        /// </summary>
        /// <param name="presentationPart">The presentation part of the presentation document.</param>
        /// <returns>String array of the text in the presentation slides.</returns>
        private String[] GetPresentationSlidesText(PresentationPart presentationPart)
        {
            var presentation    = presentationPart.Presentation;
            var slideParts      = presentationPart.SlideParts.ToList();
            var retVal          = new String[slideParts.Count()];
            var relationshipId  = String.Empty;
            var tempSlideNumber = 0;

            foreach (var slidePart in slideParts)
            {
                var slide = presentation.SlideIdList.Where(s => ((SlideId) s).RelationshipId == presentationPart.GetIdOfPart(slidePart)).FirstOrDefault();
                var index = presentation.SlideIdList.ToList().IndexOf(slide);
                relationshipId = ((SlideId) presentation.SlideIdList.ChildElements[index]).RelationshipId;
                var titles  = new List<String>();
                var content = new List<String>();
                var notes   = new List<String>();

                slidePart.Slide.Descendants<Shape>().
                          ToList().
                          ForEach(shape =>
                                  {
                                      foreach (var item in shape.Descendants<PlaceholderShape>().Where(i => i.Type != null))
                                      {
                                          if ((item.Type.ToString().ToUpper() == "CenteredTitle".ToUpper() ||
                                               item.Type.ToString().ToUpper() == "SubTitle".ToUpper()      ||
                                               item.Type.ToString().ToUpper() == "Title".ToUpper()) &&
                                              shape.TextBody != null                                &&
                                              !String.IsNullOrWhiteSpace(shape.TextBody.InnerText))
                                          {
                                              titles.AddRange(shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text));
                                          }
                                      }
                                  });

                content = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(s => s.Text).ToList();
                content.RemoveAll(s => titles.Any(t => t == s));

                if (slidePart.NotesSlidePart                                  != null &&
                    slidePart.NotesSlidePart.NotesSlide                       != null &&
                    slidePart.NotesSlidePart.NotesSlide.Descendants()         != null &&
                    slidePart.NotesSlidePart.NotesSlide.Descendants().Count() > 0)
                {
                    notes = slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().
                                      Where(s =>
                                            {
                                                return !Int32.TryParse(s.Text, out tempSlideNumber); // Remove the record as it contains the slide number.
                                            }).
                                      Select(s => s.Text).
                                      ToList();
                }

                retVal[index] = String.Join(Environment.NewLine,
                                            String.Join(Environment.NewLine, titles.ToArray()),
                                            String.Join(String.Empty,        content.ToArray()),
                                            String.Join(Environment.NewLine, notes.ToArray()));
            }

            return retVal;
        }

    }
}
