// <copyright file="ContentPopup.xaml.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

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

    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for ContentPopup window.
    /// </summary>
    public partial class ContentPopup : Window
    {
        #region Private Fields

        /// <summary>
        /// Boolean to determine whether the window can be closed.
        /// </summary>
        private bool canClose = false;

        /// <summary>
        /// Private store for the array containing the file contents.
        /// </summary>
        private string[] fileLines = null;

        /// <summary>
        /// Private store for the line number to be highlighted.
        /// </summary>
        private int lineNumber = 0;

        /// <summary>
        /// Timeout value in seconds after which content windows closes.
        /// </summary>
        private int windowCloseTimeoutSeconds = 4;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="ContentPopup"/> class.
        /// </summary>
        public ContentPopup()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Initialises a new instance of the <see cref="ContentPopup"/> class.
        /// </summary>
        /// <param name="file">The name of the file for content is displayed.</param>
        /// <param name="lineNo">The line number to set cursor position.</param>
        public ContentPopup(string file, int lineNo)
            : this()
        {
            this.fileLines = System.IO.File.ReadAllLines(file);
            this.lineNumber = lineNo;
            this.TxtContent.Text = string.Join(Environment.NewLine, this.fileLines);
            this.TxtLineNumbers.Text = string.Join(Environment.NewLine, Enumerable.Range(1, this.fileLines.Length).Select(num => num));
            this.TxtLineNumbers.Width = this.fileLines.ToString().Length * 4;
            this.GrdColLineNumber.Width = new GridLength(this.TxtLineNumbers.Width, GridUnitType.Pixel);
            this.TxtLineNumbers.Text += Environment.NewLine + string.Empty;
            this.Title = file;

            Task.Run(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    this.TxtContent.Focus();
                    this.TxtContent.CaretIndex = this.TxtContent.GetCharacterIndexFromLineIndex(lineNo - 1);
                    int visibleLinesInPopup = (this.TxtContent.GetLastVisibleLineIndex() - this.TxtContent.GetFirstVisibleLineIndex()) / 2;
                    this.TxtContent.ScrollToLine((lineNo - visibleLinesInPopup) < 0 ? 0 : (lineNo - visibleLinesInPopup));
                    this.TxtContent.Select(this.TxtContent.CaretIndex, fileLines[lineNo - 1].Length);
                });

                this.CloseWindowOnTimeout(this.WindowCloseTimeoutSeconds);
            });
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets the height of the window.
        /// </summary>
        public int WindowHeight
        {
            get
            {
                return this.WindowHeight;
            }

            set
            {
                if (value < this.MinHeight)
                {
                    value = (int)this.MinHeight;
                }
                else if (value > this.MaxHeight)
                {
                    value = (int)this.MaxHeight;
                }

                this.WindowHeight = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the window.
        /// </summary>
        public int WindowWidth
        {
            get
            {
                return this.WindowWidth;
            }

            set
            {
                if (value < this.MinWidth || value > this.MaxWidth)
                {
                    value = (int)this.MinWidth;
                }
                else if (value > this.MaxWidth)
                {
                    value = (int)this.MaxWidth;
                }

                this.WindowWidth = value;
            }
        }

        /// <summary>
        /// Gets or sets the timeout in seconds after which the popup window closes.
        /// </summary>
        public int WindowCloseTimeoutSeconds
        {
            get
            {
                return this.windowCloseTimeoutSeconds;
            }

            set
            {
                if (value < 2 || value > 20)
                {
                    value = 4;
                }

                this.windowCloseTimeoutSeconds = value;
            }
        }

        #endregion Public Properties

        #region Private Methods

        /// <summary>
        /// Close the window after the specified timeout period if no activity occurs.
        /// </summary>
        /// <param name="timeoutSeconds">The timeout in seconds after which the popup will close.</param>
        private void CloseWindowOnTimeout(int timeoutSeconds)
        {
            Task.Run(() =>
            {
                this.canClose = true;
                System.Threading.Thread.Sleep(timeoutSeconds * 1000);

                Dispatcher.Invoke(() =>
                {
                    if (this.canClose)
                    {
                        this.Close();
                    }
                });
            });
        }

        /// <summary>
        /// Close popup after mouse leaves control.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void ContentPopup_MouseLeave(object sender, MouseEventArgs e)
        {
            this.CloseWindowOnTimeout(this.WindowCloseTimeoutSeconds);
        }

        /// <summary>
        /// Prevent popup from closing if mouse is inside control.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void ContentPopupWindow_MouseEnter(object sender, MouseEventArgs e)
        {
            this.canClose = false;
        }

        /// <summary>
        /// Prevent popup from closing if mouse is inside control.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void TxtContent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        /// <summary>
        /// Scroll line numbers when scrolling main content.
        /// </summary>
        /// <param name="sender">The parameter is not used.</param>
        /// <param name="e">The parameter is not used.</param>
        private void TxtContent_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e)
        {
            this.TxtLineNumbers.ScrollToVerticalOffset(this.TxtContent.VerticalOffset);
        }

        #endregion Private Methods
    }
}
