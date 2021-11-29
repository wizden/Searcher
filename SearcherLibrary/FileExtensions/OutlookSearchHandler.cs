﻿// <copyright file="OutlookSearchHandler.cs" company="dennjose">
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
    using MsgReader.Mime;
    using MsgReader.Mime.Header;
    using MsgReader.Outlook;

    /// <summary>
    /// Class to search Outlook files.
    /// </summary>
    public class OutlookSearchHandler : FileSearchHandler
    {
        #region Internal Fields

        /// <summary>
        /// The number of characters to display before and after the matched content index.
        /// </summary>
        internal const int MaxIndexBoundary = 50;

        #endregion Internal Fields

        #region Public Properties

        /// <summary>
        /// Handles files with the .EML/.MSG/.OFT extension.
        /// </summary>
        public static new List<string> Extensions => new List<string> { ".EML", ".MSG", ".OFT" };

        #endregion Public Properties

        #region Internal Methods

        /// <summary>
        /// Search for matches in Outlook file.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        public override List<MatchedLine> Search(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            try
            {
                string fileExtension = Path.GetExtension(fileName).ToUpper();

                switch (fileExtension)
                {
                    case ".OFT":
                    case ".MSG":
                        matchedLines = this.GetMatchesInMsgOftFiles(fileName, searchTerms, matcher);
                        break;
                    case ".EML":
                        matchedLines = this.GetMatchesInEmlFiles(fileName, searchTerms, matcher);
                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }

            return matchedLines;
        }

        #endregion Internal Methods

        #region Private Methods

        /// <summary>
        /// Search for matches in the in .EML files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInEmlFiles(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();
            Message email = Message.Load(new FileInfo(fileName));
            string recipients = string.Join(", ", email.Headers.To.Select(t => GetEmailForDisplay(t)));
            string recipientsCc = string.Join(", ", email.Headers.Cc.Select(cc => GetEmailForDisplay(cc)));
            string sender = (email.Headers.Sender ?? email.Headers.From).DisplayName + " " + (email.Headers.Sender ?? email.Headers.From).Address;
            string dateSent = email.Headers.DateSent == DateTime.MinValue ? string.Empty : email.Headers.Date;
            string headerInfo = string.Join(", ", new string[] { sender, dateSent, recipients, recipientsCc, email.Headers.Subject });
            string body = email.TextBody?.GetBodyAsText();

            if (string.IsNullOrWhiteSpace(body))
            {
                body = email.HtmlBody?.GetBodyAsText();
            }

            matchedLines.AddRange(GetMatchesForHeader());
            matchedLines.AddRange(GetMatchesForBody());

            string GetEmailForDisplay(RfcMailAddress mailAddress)
            {
                if (string.IsNullOrWhiteSpace(mailAddress.DisplayName))
                {
                    return mailAddress.Address;
                }
                else
                {
                    return mailAddress.DisplayName == mailAddress.Address
                        ? mailAddress.Address
                        : mailAddress.DisplayName + " " + mailAddress.Address;
                }
            }

            // Local function to get matches in the header part of mail.
            List<MatchedLine> GetMatchesForHeader()
            {
                int matchCounter = 0;

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

            // Local function to get matches in the body of mail.
            List<MatchedLine> GetMatchesForBody()
            {
                return matcher.GetMatch(body.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries), searchTerms);
            }

            return matchedLines;
        }

        /// <summary>
        /// Search for matches in the in .MSG or .OFT files.
        /// </summary>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="searchTerms">The terms to search.</param>
        /// <param name="matcher">The matcher object to determine search criteria.</param>
        /// <returns>The matched lines containing the search terms.</returns>
        private List<MatchedLine> GetMatchesInMsgOftFiles(string fileName, IEnumerable<string> searchTerms, Matcher matcher)
        {
            List<MatchedLine> matchedLines = new List<MatchedLine>();

            using (var msg = new Storage.Message(fileName))
            {
                string recipients = msg.GetEmailRecipients(RecipientType.To, false, false);
                string recipientsCc = msg.GetEmailRecipients(RecipientType.Cc, false, false);
                string headerInfo = string.Join(", ", new string[] { msg.Sender.Raw, msg.SentOn.GetValueOrDefault().ToString(), recipients, recipientsCc, msg.Subject });

                matchedLines.AddRange(GetMatchesForHeader());
                matchedLines.AddRange(GetMatchesForBody());

                // Local function to get matches in the header part of mail.
                List<MatchedLine> GetMatchesForHeader()
                {
                    int matchCounter = 0;

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

                // Local function to get matches in the body of mail.
                List<MatchedLine> GetMatchesForBody()
                {
                    return matcher.GetMatch(msg.BodyText.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries), searchTerms);
                }
            }

            return matchedLines;
        }

        #endregion Private Methods
    }
}