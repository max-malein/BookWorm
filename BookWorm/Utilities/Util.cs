using System;
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

        /// <summary>
        /// Extracts sheet Id from url.
        /// </summary>
        /// <param name="spreadsheetUrl">Spreadsheet URL.</param>
        /// <returns>Sheet Id or null if pattern doesn't match.</returns>
        public static int? GetSheetIdFromUrl(string spreadsheetUrl)
        {
            var match = Regex.Match(spreadsheetUrl, @"(?<=edit\#gid\=)[\d]+");

            if (match.Success)
            {
                return Convert.ToInt32(match.Value);
            }

            return null;
        }
    }
}
