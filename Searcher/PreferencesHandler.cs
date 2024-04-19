// <copyright file="PreferencesHandler.cs" company="dennjose">
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
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Xml.Linq;

    /// <summary>
    /// Class to handle preferences for the executable.
    /// </summary>
    public static class PreferencesHandler
    {
        #region Fields

        /// <summary>
        /// Private store for the main search window.
        /// </summary>
        private static SearchWindow? mainSearchWindow;

        /// <summary>
        /// Private store for the preferences file.
        /// </summary>
        private static XDocument? preferencesFile;

        /// <summary>
        /// Gets the location of the preference file.
        /// </summary>
        public static string PreferenceFilePath
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SearcherPreferences.xml");
            }
        }

        public static string NoPreferenceFilePath
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NoPrefs");
            }
        }

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets the preferences file.
        /// </summary>
        public static XDocument PreferencesFile
        {
            get
            {
                if (preferencesFile == null)
                {
                    if (File.Exists(PreferenceFilePath))
                    {
                        preferencesFile = XDocument.Load(PreferenceFilePath);
                    }
                    else
                    {
                        preferencesFile = CreatePreferencesFile();
                    }
                }

                return preferencesFile;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Saves current search preferences for future use.
        /// </summary>
        /// <param name="itemsToAdd">The list of items to add.</param>
        /// <param name="preferenceElement">The search preference value.</param>
        public static void AddItemsToPreferences(IEnumerable<string> itemsToAdd, string preferenceElement)
        {
            PreferencesFile.Descendants(preferenceElement).Descendants().Remove();

            foreach (string item in itemsToAdd)
            {
                var itemToAddFromXml = PreferencesFile.Descendants(preferenceElement).FirstOrDefault();
                itemToAddFromXml?.Add(new XElement("Value", item));
            }
        }

        /// <summary>
        /// Check if any missing xml elements are there from the default settings. If yes, create the missing elements. Generally done for upgrade. Not ideal, but avoids XSD.
        /// </summary>
        /// <param name="fileName">The preferences file name.</param>
        public static void CheckPreferencesFile(string fileName)
        {
            XDocument? prefFile;

            try
            {
                prefFile = XDocument.Load(PreferenceFilePath);
            }
            catch (IOException ioe)
            {
                throw new IOException(string.Format("{0} {1}. {2}", Application.Current.Resources["PreferencesFileLoadFailure"].ToString(), fileName, ioe.Message));
            }

            if (PreferencesFile != null)
            {
                IOrderedEnumerable<string> defaultNodes = CreatePreferencesFile().Descendants().Select(t => t.Name.LocalName).OrderBy(t => t);
                IOrderedEnumerable<string> prefFileNodes = prefFile.Descendants().Select(t => t.Name.LocalName).OrderBy(t => t);
                List<string> missingNodes = defaultNodes.Except(prefFileNodes).ToList();

                if (missingNodes.Count > 0)
                {
                    foreach (string node in missingNodes)
                    {
                        if (prefFile.Root != null)
                        {
                            if (node == "MatchWholeWord")
                            {
                                prefFile.Root.Add(new XElement("MatchWholeWord", false));
                            }
                            else if (node == "MatchCase")
                            {
                                prefFile.Root.Add(new XElement("MatchCase", false));
                            }
                            else if (node == "MaxDropDownItems")
                            {
                                prefFile.Root.Add(new XElement("MaxDropDownItems", 10));
                            }
                            else if (node == "SearchSubfolders")
                            {
                                prefFile.Root.Add(new XElement("SearchSubfolders", true));
                            }
                            else if (node == "HighlightResults")
                            {
                                prefFile.Root.Add(new XElement("HighlightResults", true));
                            }
                            else if (node == "MinFileCreateSearchDate")
                            {
                                prefFile.Root.Add(new XElement("MinFileCreateSearchDate", DateTime.MinValue));
                            }
                            else if (node == "MaxFileCreateSearchDate")
                            {
                                prefFile.Root.Add(new XElement("MaxFileCreateSearchDate", DateTime.MaxValue));
                            }
                            else if (node == "Culture")
                            {
                                prefFile.Root.Add(new XElement("Culture", System.Globalization.CultureInfo.CurrentUICulture.Name));
                            }
                            else if (node == "ShowExecutionTime")
                            {
                                prefFile.Root.Add(new XElement("ShowExecutionTime", false));
                            }
                            else if (node == "SeparatorCharacter")
                            {
                                prefFile.Root.Add(new XElement("SeparatorCharacter", ";"));
                            }
                            else if (node == "SearchContentMode")
                            {
                                prefFile.Root.Add(new XElement("SearchContentMode", "Any"));
                            }
                            else if (node == "BackGroundColour")
                            {
                                prefFile.Root.Add(new XElement("BackGroundColour", "#FFFFFF"));
                            }
                            else if (node == "HighlightResultsColour")
                            {
                                prefFile.Root.Add(new XElement("HighlightResultsColour", "#FFDAB9"));
                            }
                            else if (node == "CustomEditor")
                            {
                                prefFile.Root.Add(new XElement("CustomEditor", string.Empty));
                            }
                            else if (node == "CheckForUpdates")
                            {
                                prefFile.Root.Add(new XElement("CheckForUpdates", true));
                            }
                            else if (node == "LastUpdateCheckDate")
                            {
                                prefFile.Root.Add(new XElement("LastUpdateCheckDate", DateTime.Today.AddMonths(-1).ToShortDateString()));
                            }
                            else if (node == "WindowHeight")
                            {
                                prefFile.Root.Add(new XElement("WindowHeight", mainSearchWindow?.MinHeight ?? 200));
                            }
                            else if (node == "WindowLeft")
                            {
                                prefFile.Root.Add(new XElement("WindowLeft", mainSearchWindow?.Left ?? 200));
                            }
                            else if (node == "WindowTop")
                            {
                                prefFile.Root.Add(new XElement("WindowTop", mainSearchWindow?.Top ?? 200));
                            }
                            else if (node == "WindowWidth")
                            {
                                prefFile.Root.Add(new XElement("WindowWidth", mainSearchWindow?.MinWidth ?? 200));
                            }
                            else if (node == "PopupWindowHeight")
                            {
                                prefFile.Root.Add(new XElement("PopupWindowHeight", 300));
                            }
                            else if (node == "PopupWindowWidth")
                            {
                                prefFile.Root.Add(new XElement("PopupWindowWidth", 500));
                            }
                            else if (node == "PopupWindowTimeoutSeconds")
                            {
                                prefFile.Root.Add(new XElement("PopupWindowTimeoutSeconds", 4));
                            }
                            else if (node == "SearchDirectories")
                            {
                                prefFile.Root.Add(new XElement("SearchDirectories", Array.Empty<XElement>()));
                            }
                            else if (node == "SearchContents")
                            {
                                prefFile.Root.Add(new XElement("SearchContents", Array.Empty<XElement>()));
                            }
                            else if (node == "SearchFilters")
                            {
                                prefFile.Root.Add(new XElement("SearchFilters", Array.Empty<XElement>()));
                            }
                            else if (node == "ShowFileMatchCount")
                            {
                                prefFile.Root.Add(new XElement("ShowFileMatchCount", true));
                            }
                            else if (node == "FilesToAlwaysExcludeFromSearch")
                            {
                                prefFile.Root.Add(new XElement("FilesToAlwaysExcludeFromSearch", Array.Empty<XElement>()));
                            }
                            else if (node == "DirectoriesToAlwaysExcludeFromSearch")
                            {
                                prefFile.Root.Add(new XElement("DirectoriesToAlwaysExcludeFromSearch", Array.Empty<XElement>()));
                            }
                        }
                    }

                    prefFile.Save(fileName);
                }

                preferencesFile = prefFile;
            }
        }

        /// <summary>
        /// Get the XElement nodes for the preference element.
        /// </summary>
        /// <param name="preference">Name of the preference element.</param>
        /// <returns>XElement nodes for the preference element.</returns>
        public static IEnumerable<XElement> GetPreference(string preference)
        {
            return PreferencesFile.Descendants(preference);
        }

        /// <summary>
        /// Get the value of the preference element.
        /// </summary>
        /// <param name="preference">Name of the preference element.</param>
        /// <returns>Value of the preference element.</returns>
        public static string GetPreferenceValue(string preference)
        {
            return PreferencesFile.Descendants(preference).FirstOrDefault()?.Value ?? string.Empty;
        }

        /// <summary>
        /// Save the preferences file.
        /// </summary>
        public static void SavePreferences()
        {
            PreferencesFile.Save(PreferenceFilePath);
        }

        /// <summary>
        /// Sets the main search window object for setting window specific preferences.
        /// </summary>
        /// <param name="searchWindow">The main search window object.</param>
        public static void SetMainSearchWindow(SearchWindow searchWindow)
        {
            mainSearchWindow = searchWindow;
        }

        /// <summary>
        /// Set the value for the preference element.
        /// </summary>
        /// <param name="preference">Name of the preference element.</param>
        /// <param name="value">Value of the preference element.</param>
        public static void SetPreferenceValue(string preference, string value)
        {
            if (PreferencesHandler.PreferencesFile != null)
            {
                XElement? preferenceElement = PreferencesFile.Descendants(preference).FirstOrDefault();

                if (preferenceElement != null)
                {
                    preferenceElement.Value = value;
                }
            }
        }

        /// <summary>
        /// Create a preference file.
        /// </summary>
        /// <returns>The search preference file.</returns>
        private static XDocument CreatePreferencesFile()
        {
            // Not bothering with XSD, as this is a one-off config operation, and not used for data exchange with other systems.
            XElement[] initialPreferences =
            [
                new ("MatchWholeWord", false),
                new ("MatchCase", false),
                new ("MaxDropDownItems", 10),
                new ("SearchSubfolders", true),
                new ("HighlightResults", true),
                new ("MinFileCreateSearchDate", string.Empty),
                new ("MaxFileCreateSearchDate", string.Empty),
                new ("SearchContentMode", "Any"),
                new ("ShowExecutionTime", false),
                new ("SeparatorCharacter", ";"),
                new ("BackGroundColour", "#FFFFFF"),
                new ("HighlightResultsColour", "#FFDAB9"),        // Brushes.PeachPuff
                new ("CustomEditor", string.Empty),
                new ("CheckForUpdates", true),
                new ("LastUpdateCheckDate", DateTime.Today.AddMonths(-1).ToShortDateString()),
                new ("WindowHeight", mainSearchWindow?.MinHeight ?? 200),
                new ("WindowLeft", mainSearchWindow?.Left ?? 200),
                new ("WindowTop", mainSearchWindow?.Top ?? 200),
                new ("WindowWidth", mainSearchWindow?.MinWidth ?? 200),
                new ("PopupWindowHeight", 300),
                new ("PopupWindowWidth", 500),
                new ("PopupWindowTimeoutSeconds", 4),
                new ("Culture", System.Globalization.CultureInfo.CurrentUICulture.Name),
                new ("SearchDirectories", Array.Empty<XElement>()),
                new ("SearchContents", Array.Empty<XElement>()),
                new ("SearchFilters", Array.Empty<XElement>()),
                new ("ShowFileMatchCount", true),
                new ("FilesToAlwaysExcludeFromSearch", Array.Empty<XElement>()),
                new("DirectoriesToAlwaysExcludeFromSearch", Array.Empty<XElement>())
            ];

            XDocument retVal = XDocument.Parse(new XElement("SearcherPreferences", initialPreferences).ToString(), LoadOptions.None);
            return retVal;
        }

        #endregion Methods
    }
}
