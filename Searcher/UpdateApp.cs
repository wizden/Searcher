// <copyright file="UpdateApp.cs" company="dennjose">
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

    using Newtonsoft.Json.Linq;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using System.Xml.Linq;

    /// <summary>
    /// Class to handle updates for the application.
    /// </summary>
    public class UpdateApp
    {
        #region Private Fields

        /// <summary>
        /// The date when an attempt was made to check for updates.
        /// </summary>
        private DateTime lastUpdateCheckDate = DateTime.MinValue;

        /// <summary>
        /// Private store for the default file name when downloading the latest release.
        /// </summary>
        private string latestReleaseDefaultFileName = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Searcher_new.zip");

        /// <summary>
        /// The preference file for the application.
        /// </summary>
        private XDocument preferenceFile;

        /// <summary>
        /// Private store for the site that contains the latest update.
        /// </summary>
        private string siteWithLatestUpdate = string.Empty;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="UpdateApp"/> class.
        /// </summary>
        /// <param name="preferenceFile">The preferences file for the application.</param>
        public UpdateApp(XDocument preferenceFile)
        {
            this.preferenceFile = preferenceFile;

            if (this.preferenceFile.Descendants("LastUpdateCheckDate") != null && this.preferenceFile.Descendants("LastUpdateCheckDate").FirstOrDefault() != null)
            {
                DateTime.TryParse(this.preferenceFile.Descendants("LastUpdateCheckDate").FirstOrDefault()?.Value, out lastUpdateCheckDate);
            }
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Check if a newer version has already been downloaded.
        /// </summary>
        /// <returns>A newer version exists locally.</returns>
        public bool UpdatedAppExistsLocally()
        {
            // Delete any previous version.
            System.IO.File.Delete(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Searcher.exe.old"));
            return Common.ApplicationUpdateExists;
        }

        /// <summary>
        /// Check for updated version and download if found.
        /// </summary>
        /// <returns>String containing next date to check for updates.</returns>
        public async Task<string> CheckAndDownloadUpdatedVersionAsync()
        {
            bool newerVersionExists = await this.NewerVersionExistsAsync();
            string nextCheckDate = DateTime.Today.ToString("yyyy-MM-dd");

            if (newerVersionExists)
            {
                if (System.IO.File.Exists(this.latestReleaseDefaultFileName))
                {
                    string newProgPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg");

                    try
                    {
                        System.IO.Compression.ZipFile.ExtractToDirectory(this.latestReleaseDefaultFileName, newProgPath);
                    }
                    catch (Exception ex)
                    {
                        if (ex is IOException || ex is InvalidDataException)
                        {
                            // Nothing to handle. Failed to extract, so either file is invalid (re-download) or the extract was reattempted (remove zip file). In either case, remove downloaded zip file.
                        }
                        else
                        {
                            // Throw only if the exception is not of type IOException or InvalidDataException.
                            throw;
                        }
                    }

                    System.IO.File.Delete(this.latestReleaseDefaultFileName);
                }
            }
            else
            {
                if (this.lastUpdateCheckDate.AddMonths(1) > DateTime.Today)
                {
                    nextCheckDate = this.lastUpdateCheckDate.ToString("yyyy-MM-dd");
                }
            }

            return nextCheckDate;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Determine if a new release exists online.
        /// </summary>
        /// <returns>Boolean indicating whether a new release exists online.</returns>
        private async Task<bool> NewerVersionExistsAsync()
        {
            bool retVal = false;
            if (preferenceFile.Descendants("CheckForUpdates").FirstOrDefault()?.Value.ToUpper() == true.ToString().ToUpper())
            {
                // Check for updates monthly. Why bother the user more frequently. Can look to make this configurable in the future.
                if (this.lastUpdateCheckDate.AddMonths(1) < DateTime.Today)
                {
                    retVal = await this.NewReleaseFoundOnlineAsync();
                }
            }

            return retVal;
        }

        /// <summary>
        /// Get the name of the updated application download file.
        /// </summary>
        /// <param name="downloadUrl">The url path for the latest executable.</param>
        /// <returns>The name of the updated application download file.</returns>
        private async Task<string> GetDownloadedUpdateFilenameAsync(string downloadUrl)
        {
            string retVal = string.Empty;

            retVal = await Task.Run<string>(async () =>
            {
                string newFileName = string.Empty;

                if (!string.IsNullOrEmpty(this.siteWithLatestUpdate))
                {
                    if (!string.IsNullOrEmpty(downloadUrl))
                    {
                        try
                        {
                            if (File.Exists(this.latestReleaseDefaultFileName))
                            {
                                File.Delete(this.latestReleaseDefaultFileName);
                            }

                            using (HttpClient client = new HttpClient())
                            {
                                newFileName = this.latestReleaseDefaultFileName;
                                var response = await client.GetAsync(new Uri(downloadUrl), HttpCompletionOption.ResponseContentRead,
                                    new CancellationTokenSource(new TimeSpan(0, 5, 0)).Token);
                                
                                using (var fs = new FileStream(newFileName, FileMode.CreateNew))
                                {
                                    await response.Content.CopyToAsync(fs);
                                }
                            }
                        }
                        catch
                        {
                            // If the exception occurs after newFileName is set, set newFileName to empty string.
                            newFileName = string.Empty;
                        }
                    }
                }

                return newFileName;
            });

            return retVal;
        }

        /// <summary>
        /// Gets the download path of the latest release in GitHub.
        /// </summary>
        /// <returns>The download path of the latest release in GitHub.</returns>
        private async Task<string> GetLatestReleaseNetFrameworkDownloadPathInGitHubAsync()
        {
            string downloadUrl = string.Empty;
            Octokit.GitHubClient client = new Octokit.GitHubClient(new Octokit.ProductHeaderValue("searcher"));
            Octokit.Release latestRelease = await client.Repository.Release.Get("wizden", "searcher", "v1.0.46");

            if (latestRelease != null && latestRelease.Assets != null && latestRelease.Assets.Count > 0)
            {
                downloadUrl = latestRelease.Assets[0].BrowserDownloadUrl;
            }

            return downloadUrl;
        }

        /// <summary>
        /// Determine if a new application version exists.
        /// </summary>
        /// <returns>Boolean indicating whether a new application version exists.</returns>
        private async Task<bool> NewReleaseFoundOnlineAsync()
        {
            bool retVal = false;

            if (System.IO.Directory.Exists(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NewProg")))
            {
                retVal = true;
            }
            else
            {
                string downloadedFile = string.Empty;
                string downloadUrl = string.Empty;

                if (await this.NewReleaseExistsInSourceForgeAsync())
                {
                    this.siteWithLatestUpdate = "SourceForge";
                    downloadUrl = "https://sourceforge.net/projects/searcher/files/latest/download";
                    downloadedFile = await this.GetDownloadedUpdateFilenameAsync(downloadUrl);
                    retVal = !string.IsNullOrEmpty(downloadedFile);
                }

                if (!retVal && await this.NewReleaseExistsInGitHubAsync())
                {
                    this.siteWithLatestUpdate = "GitHub";
                    downloadUrl = await this.GetLatestReleaseNetFrameworkDownloadPathInGitHubAsync();
                    downloadedFile = await this.GetDownloadedUpdateFilenameAsync(downloadUrl);
                    retVal = !string.IsNullOrEmpty(downloadedFile);
                }

                if (!retVal && await this.NewReleaseExistsInGitHubAsync_NETCore())
                {
                    this.siteWithLatestUpdate = "GitHub";
                    bool hasRuntime = this.IsNETWindowsDesktopRuntimeInstalled();
                    downloadUrl = await this.GetGitHubDownloadLinkAsync_NETCore(hasRuntime);
                    downloadedFile = await this.GetDownloadedUpdateFilenameAsync(downloadUrl);
                    retVal = !string.IsNullOrEmpty(downloadedFile);
                }
            }

            return retVal;
        }

        /// <summary>
        /// Check if a new version exists on GitHub.
        /// </summary>
        /// <returns>Boolean indicating whether a new version exists.</returns>
        private async Task<bool> NewReleaseExistsInGitHubAsync()
        {
            bool retVal = false;

            retVal = await Task.Run<bool>(async () =>
            {
                bool newReleaseExists = false;
                string latestReleaseDownloadUrl = string.Empty;

                try
                {
                    latestReleaseDownloadUrl = await this.GetLatestReleaseNetFrameworkDownloadPathInGitHubAsync();
                }
                catch
                {
                    // Cannot do much. Failed to access/retrieve data from the website.
                }

                if (!string.IsNullOrEmpty(latestReleaseDownloadUrl))
                {
                    string searchValue = "/Searcher_v";
                    string strSiteVersion = latestReleaseDownloadUrl.Substring(latestReleaseDownloadUrl.IndexOf(searchValue) + searchValue.Length).Replace(".zip", string.Empty);
                    Version appVersion = new Version(Common.VersionNumber);
                    Version? siteVersion;

                    if (Version.TryParse(strSiteVersion, out siteVersion))
                    {
                        newReleaseExists = siteVersion > appVersion;
                    }
                }

                return newReleaseExists;
            });

            return retVal;
        }

        /// <summary>
        /// Check if a new version exists on GitHub for the NETCore version.
        /// </summary>
        /// <returns>Boolean indicating whether a new version exists.</returns>
        private async Task<bool> NewReleaseExistsInGitHubAsync_NETCore()
        {
            bool retVal = false;

            retVal = await Task.Run<bool>(async () =>
            {
                bool newReleaseExists = false;
                string latestReleaseDownloadUrl = string.Empty;

                try
                {
                    bool hasRuntime = this.IsNETWindowsDesktopRuntimeInstalled();
                    latestReleaseDownloadUrl = await this.GetGitHubDownloadLinkAsync_NETCore(hasRuntime);
                }
                catch
                {
                    // Cannot do much. Failed to access/retrieve data from the website.
                }

                if (!string.IsNullOrEmpty(latestReleaseDownloadUrl))
                {
                    string searchValue = "/Searcher_v";
                    string strSiteVersion = latestReleaseDownloadUrl.Substring(latestReleaseDownloadUrl.IndexOf(searchValue) + searchValue.Length)
                        .Replace(".Portable", string.Empty)
                        .Replace(".Light", string.Empty)
                        .Replace(".x64", string.Empty)
                        .Replace(".x86", string.Empty)
                        .Replace(".exe", string.Empty)
                        .Replace(".zip", string.Empty);
                    Version appVersion = new Version(Common.VersionNumber);
                    Version? siteVersion;

                    if (Version.TryParse(strSiteVersion, out siteVersion))
                    {
                        newReleaseExists = siteVersion > appVersion;
                    }
                }

                return newReleaseExists;
            });

            return retVal;
        }

        /// <summary>
        /// Check if a new version exists on SourceForge.
        /// </summary>
        /// <returns>Boolean indicating whether a new version exists.</returns>
        private async Task<bool> NewReleaseExistsInSourceForgeAsync()
        {
            bool retVal = await Task.Run<bool>(async () =>
            {
                bool newReleaseExists = false;
                string path = @"https://sourceforge.net/projects/searcher/best_release.json";
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = new TimeSpan(0, 0, 0, 30);
                    string latestReleaseDownloadUrl = string.Empty;

                    try
                    {
                        latestReleaseDownloadUrl = await client.GetStringAsync(new Uri(path));
                    }
                    catch
                    {
                        // Cannot do much. Failed to access/retrieve data from the website.
                    }

                    if (!string.IsNullOrEmpty(latestReleaseDownloadUrl))
                    {
                        JObject jsonSiteVersion = JObject.Parse(latestReleaseDownloadUrl);
                        string fileName = jsonSiteVersion["platform_releases"]?["windows"]?["filename"]?.ToString() ?? string.Empty;

                        if (!string.IsNullOrWhiteSpace(fileName))
                        {
                            string strSiteVersion = Regex.Replace(Path.GetFileName(fileName)
                                .Replace("Searcher_v", string.Empty)
                                .Replace(".zip", string.Empty),
                                ".[A-Z].*", "");
                            Version appVersion = new Version(Common.VersionNumber);
                            Version? siteVersion;

                            if (Version.TryParse(strSiteVersion, out siteVersion))
                            {
                                newReleaseExists = siteVersion > appVersion;
                            }
                        }
                    }

                    return newReleaseExists;
                }
            });

            return retVal;
        }

        /// <summary>
        /// Checks if dotnet runtime is installed on the system.
        /// </summary>
        /// <returns>Boolean indicating whether the dotnet runtime is installed on the system.</returns>
        private bool IsNETWindowsDesktopRuntimeInstalled()
        {
            bool retVal = false;
            Process prcDesktopRuntimeInstalled = new Process();
            prcDesktopRuntimeInstalled.StartInfo = new ProcessStartInfo
            {
                FileName = "dotnet.exe",
                Arguments = "--list-runtimes",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };

            try
            {
                prcDesktopRuntimeInstalled.Start();
                string output = prcDesktopRuntimeInstalled.StandardOutput.ReadToEnd();
                string error = prcDesktopRuntimeInstalled.StandardError.ReadToEnd();

                if (string.IsNullOrWhiteSpace(error))
                {
                    string desktopRuntimeName = "Microsoft.WindowsDesktop.App";

                    if (output.Contains(desktopRuntimeName))
                    {
                        short appNetRuntimeVersion = 6;     // .NET 6 minimum version must be installed.
                        var desktopRuntimes = output.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                            .Where(rt => rt.Contains(desktopRuntimeName));
                        var versions = desktopRuntimes.Select(rt =>
                            Version.Parse(rt.Substring(desktopRuntimeName.Length, rt.IndexOf("[") - desktopRuntimeName.Length).Trim()));
                        retVal = versions.Any(v => v.Major >= appNetRuntimeVersion);
                    }
                }
            }
            catch
            {
                retVal = false; // Could not find dotnet runtime.
            }

            return retVal;
        }

        /// <summary>
        /// Gets the download link in GitHub for the NET Core version of the software.
        /// </summary>
        /// <param name="isNETWindowsDesktopRuntimeInstalled">Is the .NET Windows Desktop Runtime installed. Used to determine whether we get the SelfContained or FrameworkDependant version.</param>
        /// <returns>String contianing the download URL for the application.</returns>
        private async Task<string> GetGitHubDownloadLinkAsync_NETCore(bool isNETWindowsDesktopRuntimeInstalled)
        {
            string processorArchitecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString();
            //  Searcher_v2.x.x.Light.x86.exe
            //  Searcher_v2.x.x.Light.x64.exe
            //  Searcher_v2.x.x.Portable.x86.exe 
            //  Searcher_v2.x.x.Portable.x64.exe 
            string portableType = isNETWindowsDesktopRuntimeInstalled ? "Light" : "Portable";
            string urlSubstring = $"{portableType}.{processorArchitecture}";

            string downloadUrl = string.Empty;
            Octokit.GitHubClient client = new Octokit.GitHubClient(new Octokit.ProductHeaderValue("searcher"));
            Octokit.Release latestRelease = await client.Repository.Release.GetLatest("wizden", "searcher");

            if (latestRelease != null && latestRelease.Assets != null && latestRelease.Assets.Count > 0)
            {
                downloadUrl = latestRelease.Assets
                    .FirstOrDefault(a => a.BrowserDownloadUrl.ToLower()
                        .Contains(urlSubstring.ToLower()))?.BrowserDownloadUrl ?? string.Empty;
            }

            return downloadUrl;
        }

        #endregion Private Methods
    }
}
