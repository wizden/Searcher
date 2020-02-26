// <copyright file="ResultsPopout.xaml.cs" company="dennjose">
//     www.dennjose.com. All rights reserved.
// </copyright>
// <author>Dennis Joseph</author>

namespace Searcher
{
    using System.Windows;

    /// <summary>
    /// Interaction logic for Results pop out window.
    /// </summary>
    public partial class ResultsPopout : Window
    {
        /// <summary>
        /// Initialises a new instance of the <see cref="ResultsPopout"/> class.
        /// </summary>
        public ResultsPopout()
        {
            this.InitializeComponent();
            this.SetContentBasedOnLanguage();
        }

        /// <summary>
        /// Set readable content based on selected language.
        /// </summary>
        private void SetContentBasedOnLanguage()
        {
            this.Title = Application.Current.Resources["ResultsPopout"].ToString();
        }
    }
}
