using System.Text.RegularExpressions;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Utility functions.
    /// </summary>
    public static class Util
    {
        /// <summary>
        /// Extracts spreadsheet id from url.
        /// </summary>
        /// <param name="spreadsheetUrl">Spreadsheet url.</param>
        /// <returns>Spreadsheet id or original string if it doesn't match url pattern.</returns>
        public static string ParseUrl(string spreadsheetUrl)
        {
            var match = Regex.Match(spreadsheetUrl, @"(?<=docs\.google\.com\/spreadsheets\/d\/)[\w-]+");

            return match.Success ? match.Value : spreadsheetUrl;
        }
    }
}
