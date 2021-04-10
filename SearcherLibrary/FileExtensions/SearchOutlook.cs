// <copyright file="SearchOutlook.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace SearcherLibrary.FileExtensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;

    /// <summary>
    /// Class to search Outlook file.
    /// </summary>
    internal class SearchOutlook
    {
        /// <summary>
        /// Search for matches in ODP file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        internal List<MatchedLine> GetMatchesInOutlook(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                using (var msg = new MsgReader.Outlook.Storage.Message(fileName))
                {
                    string recipients = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false);
                    string recipientsCc = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false);

                    if (!string.IsNullOrWhiteSpace(recipientsCc))
                    {
                        recipients += Environment.NewLine + recipientsCc;
                    }

                    string headerInfo = string.Join(Environment.NewLine, new string[] { msg.Sender.Raw, msg.SentOn.GetValueOrDefault().ToString(), recipients, msg.Subject });

                    matchedLines.AddRange(this.GetMatchesForHeader(fileName, searchTerms, matcher, headerInfo));
                    matchedLines.AddRange(this.GetMatchesForSubject(fileName, searchTerms, matcher, msg.BodyText));
                }
            }
            catch (Exception)
            {
                throw;
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in ODP file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <param name="headerInfo">The content of the mail header.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesForHeader(string fileName, IEnumerable<string> searchTerms, Matcher matcher, string headerInfo)
        {
            int matchCounter = 0;
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            foreach (string searchTerm in searchTerms)
            {
                MatchCollection matches = Regex.Matches(headerInfo, searchTerm, matcher.RegexOptions);            // Use this match for getting the locations of the match.
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        matchedLines.Add(new MatchedLine
                        {
                            MatchId = matchCounter++,
                            Content = string.Format("{0}:\t{1}", Resources.Strings.Header, headerInfo),
                            SearchTerm = searchTerm,
                            FileName = fileName,
                            LineNumber = 1,
                            StartIndex = match.Index + Resources.Strings.Header.Length + 2,
                            Length = match.Length
                        });
                    }
                }
            }            

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in ODP file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <param name="body">The content of the mail body.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesForSubject(string fileName, IEnumerable<string> searchTerms, Matcher matcher, string body)
        {
            List<MatchedLine> matchedLines = matcher.GetMatch(body.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries), searchTerms);
            return matchedLines;
        }
    }
}
