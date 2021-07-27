using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Google.Apis.Sheets.v4.Data;
using Color = Google.Apis.Sheets.v4.Data.Color;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Helper operations with Google Sheets data types for sheets.
    /// </summary>
    public static class SheetsUtilities
    {
        /// <summary>
        /// Convert System.Drawing.Color-type ARGB-color into Google.Apis.Sheets.v4.Data.Color-type color.
        /// </summary>
        /// <param name="colorARGB">System.Drawing.Color-type color.</param>
        /// <returns>Google.Apis.Sheets.v4.Data.Color-type color.</returns>
        public static Color GetGoogleSheetsColor(System.Drawing.Color colorARGB)
        {
            var googleSheetColor = new Color
            {
                Alpha = colorARGB.A / 255f,
                Red = colorARGB.R / 255f,
                Green = colorARGB.G / 255f,
                Blue = colorARGB.B / 255f,
            };

            return googleSheetColor;
        }

        /// <summary>
        /// Convert Google.Apis.Sheets.v4.Data.Color-type color into System.Drawing.Color-type ARGB-color.
        /// </summary>
        /// <param name="googleSheetsColor">Google.Apis.Sheets.v4.Data.Color-type color.</param>
        /// <returns>System.Drawing.Color-type ARGB-color.</returns>
        public static System.Drawing.Color GetSystemDrawingColor(Color googleSheetsColor)
        {
            var colorARGB = System.Drawing.Color.Empty;

            if (googleSheetsColor != null)
            {
                var alpha = 255;

                if (googleSheetsColor.Alpha != null)
                {
                    alpha = Convert.ToInt32(googleSheetsColor.Alpha * 255);
                }

                var red = Convert.ToInt32(googleSheetsColor.Red * 255);
                var green = Convert.ToInt32(googleSheetsColor.Green * 255);
                var blue = Convert.ToInt32(googleSheetsColor.Blue * 255);

                colorARGB = System.Drawing.Color.FromArgb(alpha, red, green, blue);
            }

            return colorARGB;
        }

        /// <summary>
        /// Formats color for gh panel.
        /// </summary>
        /// <param name="colorARGB">System.Drawing.Color-type color.</param>
        /// <returns>RGB-color as "R,G,B" string.</returns>
        public static string GetFormattedARGB(System.Drawing.Color colorARGB)
        {
            return $"{colorARGB.R},{colorARGB.G},{colorARGB.B}";
        }

        /// <summary>
        /// Gets spreadsheet range in A1 notation wich will retrive from the spreadsheet.
        /// If range include only sheet name, its refer all cells in named sheet.
        /// If range include only cell range its refer cells of the first visible sheet.
        /// </summary>
        /// <param name="sheetName">Sheet name.</param>
        /// <param name="cellRange">Range of cells in A1 notaton.</param>
        /// <returns>Spreadsheet range in A1 notation.</returns>
        public static string GetSpreadsheetRange(string sheetName, string cellRange)
        {
            var spreadsheetRange = string.Empty;

            if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellRange))
            {
                spreadsheetRange = $"\'{sheetName}\'!{cellRange}";
            }
            else if (!string.IsNullOrEmpty(sheetName))
            {
                spreadsheetRange = sheetName;
            }
            else if (!string.IsNullOrEmpty(cellRange))
            {
                spreadsheetRange = cellRange;
            }

            return spreadsheetRange;
        }

        /// <summary>
        /// Add new sheet to the spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <param name="sheetName">Sheet Name.</param>
        /// <returns>Sheet Id of created sheet.</returns>
        public static int? CreateNewSheet(string spreadsheetId, string sheetName)
        {
            var requests = new List<Request>();

            var addSheetRequest = new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties { Title = sheetName },
                },
            };

            requests.Add(addSheetRequest);

            var requestBody = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests,
            };

            var request = Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();
            var sheetId = response.Replies.FirstOrDefault().AddSheet.Properties.SheetId;

            return sheetId;
        }

        /// <summary>
        /// Get sheet Id by it's name.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <param name="sheetName">Sheet Name.</param>
        /// <returns>Sheet Id or null if sheet doesn't exist.</returns>
        public static int? GetSheetId(string spreadsheetId, string sheetName)
        {
            var request = Credentials.Service.Spreadsheets.Get(spreadsheetId);

            // For partial response. Multiple different fields are comma separated and subfields are dot-separated.
            // But if 403 error has happen use slashes.
            // Field names can be specified in camelCase or separated_by_underscores.
            // For convenience, multiple subfields from the same type can be listed within parentheses.
            // And without whitespaces.
            request.Fields = "sheets.properties(title,sheetId)";

            var spreadsheet = request.Execute();

            var sheets = spreadsheet.Sheets;

            foreach (var sheet in sheets)
            {
                if (sheet.Properties.Title == sheetName)
                {
                    return sheet.Properties.SheetId;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets all sheets names on the spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <returns>Sheets names.</returns>
        public static List<string> GetAllSheetNames(string spreadsheetId)
        {
            var request = Credentials.Service.Spreadsheets.Get(spreadsheetId);
            request.Fields = "sheets.properties.title";
            var spreadsheet = request.Execute();
            var sheets = spreadsheet.Sheets;

            return sheets.Select(s => s.Properties.Title).ToList();
        }

        /// <summary>
        /// Gets ranges of merged cells on the sheet.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <param name="spreadsheetRange">The range to retrieve from spreadsheet.</param>
        /// <returns>The ranges that are merged together.</returns>
        public static List<GridRange> GetMergeRanges(string spreadsheetId, string spreadsheetRange)
        {
            var request = Credentials.Service.Spreadsheets.Get(spreadsheetId);
            request.Ranges = spreadsheetRange;
            request.Fields = "sheets.merges";

            var spreadsheet = request.Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault();

            return sheet.Merges.ToList();
        }

        /// <summary>
        /// Gets top left cell coordinates in range of merged cells.
        /// </summary>
        /// <param name="coordinates">Coordinates of the cell in a form of a Point.</param>
        /// <param name="mergeData">GridRange of merged cells.</param>
        /// <returns>Point representation of cell coordinates, where X - column index and Y - row index.</returns>
        internal static Point? FindMergeOrigin(Point? coordinates, List<GridRange> mergeData)
        {
            var column = coordinates.Value.X;
            var containsColumn = mergeData.Where(md => md.StartColumnIndex <= column && md.EndColumnIndex - 1 >= column).ToList();
            if (containsColumn.Count == 0)
                return null;
            var row = coordinates.Value.Y;
            var containsColumnAndRow = containsColumn.Where(md => md.StartRowIndex <= row && md.EndRowIndex - 1 >= row).ToList();
            if (containsColumnAndRow.Count == 0)
                return null;
            return new Point?(new Point(containsColumnAndRow[0].StartColumnIndex.Value, containsColumnAndRow[0].StartRowIndex.Value));
        }
    }
}
