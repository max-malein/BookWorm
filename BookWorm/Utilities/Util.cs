using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
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

        internal static List<List<string>> GetCellCoordinates(List<RowData> rowData, string range)
        {
            return new List<List<string>>();
            throw new NotImplementedException();
        }
    }
}
