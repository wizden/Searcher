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
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Xml.Linq;

    public static class PreferencesHandler
    {
        /// <summary>
        /// Private store for the main search window.
        /// </summary>
        private static SearchWindow mainSearchWindow;

        /// <summary>
        /// Private store for the preferences file.
        /// </summary>
        private static XDocument preferencesFile;

        /// <summary>
        /// Private store for the location of the preference file.
        /// </summary>
        public static string PreferenceFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SearcherPreferences.xml");

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
                        LoadPreferences();
                    }
                    else
                    {
                        preferencesFile = CreatePreferencesFile();
                    }
                }

                return preferencesFile;
            }
        }

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
                PreferencesFile.Descendants(preferenceElement).FirstOrDefault().Add(new XElement("Value", item));
            }
        }

        /// <summary>
        /// Check if any missing xml elements are there from the default settings. If yes, create the missing elements. Generally done for upgrade. Not ideal, but avoids XSD.
        /// </summary>
        public static void CheckPreferencesFile(string fileName)
        {
            XDocument prefFile = null;

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
                            prefFile.Root.Add(new XElement("WindowHeight", (mainSearchWindow?.MinHeight ?? 200)));
                        }
                        else if (node == "WindowWidth")
                        {
                            prefFile.Root.Add(new XElement("WindowWidth", (mainSearchWindow?.MinWidth ?? 200)));
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
                            prefFile.Root.Add(new XElement("SearchDirectories", new XElement[] { null }));
                        }
                        else if (node == "SearchContents")
                        {
                            prefFile.Root.Add(new XElement("SearchContents", new XElement[] { null }));
                        }
                        else if (node == "SearchFilters")
                        {
                            prefFile.Root.Add(new XElement("SearchFilters", new XElement[] { null }));
                        }
                        else if (node == "ShowFileMatchCount")
                        {
                            prefFile.Root.Add(new XElement("ShowFileMatchCount", true));
                        }
                        else if (node == "FilesToAlwaysExcludeFromSearch")
                        {
                            prefFile.Root.Add(new XElement("FilesToAlwaysExcludeFromSearch", new XElement[] { null }));
                        }
                        else if (node == "DirectoriesToAlwaysExcludeFromSearch")
                        {
                            prefFile.Root.Add(new XElement("DirectoriesToAlwaysExcludeFromSearch", new XElement[] { null }));
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
            try
            {
                return PreferencesFile.Descendants(preference);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get the value of the preference element.
        /// </summary>
        /// <param name="preference">Name of the preference element.</param>
        /// <returns>Value of the preference element.</returns>
        public static string GetPreferenceValue(string preference)
        {
            try
            {
                return PreferencesFile.Descendants(preference).FirstOrDefault().Value;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Load the preferences file.
        /// </summary>
        public static void LoadPreferences()
        {
            preferencesFile = XDocument.Load(PreferenceFilePath);
        }

        /// <summary>
        /// Save the preferences file.
        /// </summary>
        public static void SavePreferences()
        {
            PreferencesFile.Save(PreferenceFilePath);
        }

        /// <summary>
        /// Set the value for the preference element.
        /// </summary>
        /// <param name="preference">Name of the preference element.</param>
        /// <param name="value">Value of the preference element.</param>
        public static void SetPreferenceValue(string preference, string value)
        {
            try
            {
                PreferencesFile.Descendants(preference).FirstOrDefault().Value = value;
            }
            catch (Exception)
            {
                throw;
            }
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
        /// Create a preference file.
        /// </summary>
        /// <returns>The search preference file.</returns>
        private static XDocument CreatePreferencesFile()
        {
            // Not bothering with XSD, as this is a one-off config operation, and not used for data exchange with other systems.
            XElement[] initialPreferences = new XElement[]
            {
                new XElement("MatchWholeWord", false),
                new XElement("MatchCase", false),
                new XElement("MaxDropDownItems", 10),
                new XElement("SearchSubfolders", true),
                new XElement("HighlightResults", true),
                new XElement("MinFileCreateSearchDate", string.Empty),
                new XElement("MaxFileCreateSearchDate", string.Empty),
                new XElement("SearchContentMode", "Any"),
                new XElement("ShowExecutionTime", false),
                new XElement("SeparatorCharacter", ";"),
                new XElement("BackGroundColour", "#FFFFFF"),
                new XElement("HighlightResultsColour", "#FFDAB9"),        // Brushes.PeachPuff
                new XElement("CustomEditor", string.Empty),
                new XElement("CheckForUpdates", true),
                new XElement("LastUpdateCheckDate", DateTime.Today.AddMonths(-1).ToShortDateString()),
                new XElement("WindowHeight", (mainSearchWindow?.MinHeight ?? 200)),
                new XElement("WindowWidth", (mainSearchWindow?.MinWidth ?? 200)),
                new XElement("PopupWindowHeight", 300),
                new XElement("PopupWindowWidth", 500),
                new XElement("PopupWindowTimeoutSeconds", 4),
                new XElement("Culture", System.Globalization.CultureInfo.CurrentUICulture.Name),
                new XElement("SearchDirectories", new XElement[] { null }),
                new XElement("SearchContents", new XElement[] { null }),
                new XElement("SearchFilters", new XElement[] { null }),
                new XElement("ShowFileMatchCount", true),
                new XElement("FilesToAlwaysExcludeFromSearch", new XElement[] { null }),
                new XElement("DirectoriesToAlwaysExcludeFromSearch", new XElement[] { null })
            };

            XDocument retVal = XDocument.Parse(new XElement("SearcherPreferences", initialPreferences).ToString(), LoadOptions.None);
            return retVal;
        }
    }
}
