using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4.Data;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Helper operations with Google Sheets data types for cells.
    /// </summary>
    public static class CellsUtilities
    {
        /// <summary>
        /// Splits cells list in rows by grid range.
        /// </summary>
        /// <param name="cells">Cells.</param>
        /// <param name="gridRange">Grid range.</param>
        /// <returns>Rows.</returns>
        public static List<RowData> GetRows(List<CellData> cells, GridRange gridRange)
        {
            if (gridRange == null) return null;

            var rowLength = gridRange.EndColumnIndex - gridRange.StartColumnIndex;

            var rows = new List<RowData>();
            var row = new RowData();
            var rowValues = new List<CellData>();

            for (int i = 0; i < cells.Count; i++)
            {
                rowValues.Add(cells[i]);

                if ((i + 1) % rowLength == 0)
                {
                    row.Values = rowValues;
                    rows.Add(row);

                    row = new RowData();
                    rowValues = new List<CellData>();
                }
            }

            // if number of cellData is less than cells in GridRange the last row is not added in the for-loop
            // it needs to be added manually
            if (rowValues.Any())
            {
                row.Values = rowValues;
                rows.Add(row);
            }

            return rows;
        }

        /// <summary>
        /// Creates GridRange object from A1 notation range.
        /// </summary>
        /// <param name="CellRangeA1">Cells range in A1 notation.</param>
        /// <param name="sheetId">Sheet Id.</param>
        /// <returns>new Grid Range.</returns>
        public static GridRange GridRangeFromA1(string CellRangeA1, int sheetId)
        {
            var gridRange = new GridRange();
            gridRange.SheetId = sheetId;

            //bool letters = Regex.Matches(CellRangeA1, @"[a-zA-Z]").Count > 0;
            //bool numbers = CellRangeA1.Any(c => char.IsDigit(c));

            //var rangeBounds = CellRangeA1.Split(':');
            var rangeBounds = CellRangeA1.Split(':');
            string startLetters = Regex.Match(rangeBounds[0], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
            string startNumbers = Regex.Match(rangeBounds[0], @"\d+").Value;


            //CellRangeA1 = "D:A6";
            if (!CellRangeA1.Contains(':'))
            {
                // Case A5
                gridRange.StartColumnIndex = ColumnNameToNumber(startLetters) - 1;
                gridRange.StartRowIndex = Convert.ToInt32(startNumbers) - 1;
                // Case SheetName
            }
            else
            {
                string endLetters = Regex.Match(rangeBounds[1], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                string endNumbers = Regex.Match(rangeBounds[1], @"\d+").Value;

                // Case A:B
                if (startNumbers.Length == 0 && endNumbers.Length == 0)
                {
                    gridRange.StartColumnIndex = ColumnNameToNumber(startLetters) - 1;
                    gridRange.EndColumnIndex = ColumnNameToNumber(endLetters);
                }

                // Case A5:A
                if (endNumbers.Length == 0)
                {
                    gridRange.StartRowIndex = ColumnNameToNumber(startLetters) - 1;
                    gridRange.EndColumnIndex = ColumnNameToNumber(endLetters);
                    gridRange.StartRowIndex = Convert.ToInt32(startNumbers) - 1;
                }

                // Case 2:4
                else if (startLetters.Length == 0 && endLetters.Length == 0)
                {
                    gridRange.StartRowIndex = Convert.ToInt32(startNumbers) - 1;
                    gridRange.EndRowIndex = Convert.ToInt32(endNumbers);
                }

                // Case 5:B5
                else if (startLetters.Length == 0)
                {
                    gridRange.EndColumnIndex = ColumnNameToNumber(endLetters);
                    gridRange.StartRowIndex = Convert.ToInt32(startNumbers) - 1;
                    gridRange.EndRowIndex = Convert.ToInt32(endNumbers);
                }

                // Case A3:D4
                else
                {
                    gridRange.StartColumnIndex = ColumnNameToNumber(startLetters) - 1;
                    gridRange.EndColumnIndex = ColumnNameToNumber(endLetters);
                    gridRange.StartRowIndex = Convert.ToInt32(startNumbers) - 1;
                    gridRange.EndRowIndex = Convert.ToInt32(endNumbers);
                }

            }

            return gridRange;
        }

        /// <summary>
        /// Gets the cells coordinates which represents their row and column indices.
        /// </summary>
        /// <param name="rowsData">Rows of cells in range.</param>
        /// <param name="startRowInd">Index of first row in cells range.</param>
        /// <param name="startColumnInd">Index of first column in cells range.</param>
        /// <returns>Row and column indicies of the cells.</returns>
        internal static List<List<int[]>> GetCellCoordinates(IList<RowData> rowsData, int startRowInd, int startColumnInd)
        {
            var coordinatesInGrid = new List<List<int[]>>();

            var rowInd = startRowInd;

            foreach (var row in rowsData)
            {
                var coordinatesInRow = new List<int[]>();
                var columnInd = startColumnInd;

                // If row full of null-cells
                if (row.Values == null)
                {
                    coordinatesInRow = null;
                    coordinatesInGrid.Add(coordinatesInRow);
                    rowInd++;
                    continue;
                }

                // Null-cell within right bound of non-null-cell is ok, so it also will get coordinate.
                // So does rows. I.e. Rows and cells as "null-data-null-data-null" actually becomes "null-data-null-data" in response.
                // Merged null-cells are considered as non-null-cells.
                foreach (var cell in row.Values)
                {
                    var indexCoordinates = new int[2];

                    indexCoordinates[0] = rowInd;
                    indexCoordinates[1] = columnInd;

                    coordinatesInRow.Add(indexCoordinates);

                    columnInd++;
                }

                coordinatesInGrid.Add(coordinatesInRow);

                rowInd++;
            }

            return coordinatesInGrid;
        }

        /// <summary>
        /// Convert column letter to index.
        /// </summary>
        /// <param name="columnName">Column name.</param>
        /// <returns>Column number for column name.</returns>
        private static int ColumnNameToNumber(string columnName)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                char letter = columnName[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }

            return result;
        }
    }
}
