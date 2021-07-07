using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4.Data;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Helper operations with Google Sheets data types for cells.
    /// </summary>
    public class CellsUtilities
    {
        /// <summary>
        /// Splits cells list in rows by grid range.
        /// </summary>
        /// <param name="cells">Cells.</param>
        /// <param name="gridRange">Grid range.</param>
        /// <returns>Rows.</returns>
        public List<RowData> GetRows(List<CellData> cells, GridRange gridRange)
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
        /// <param name="a1NotatonRange">Cells range in A1 notation.</param>
        /// <param name="sheetId">Sheet Id.</param>
        /// <returns>new Grid Range.</returns>
        public GridRange GridRangeFromA1(string a1NotatonRange, int sheetId)
        {
            var gridRange = new GridRange();
            gridRange.SheetId = sheetId;

            // problems with lower-case and column name to number
            bool letters = Regex.Matches(a1NotatonRange, @"[a-zA-Z]").Count > 0;
            bool numbers = a1NotatonRange.Any(c => char.IsDigit(c));

            var rangeBounds = a1NotatonRange.Split(':');

            // Only columns case.
            if (letters && !numbers)
            {
                gridRange.StartColumnIndex = ColumnNameToNumber(rangeBounds[0]) - 1;
                gridRange.EndColumnIndex = ColumnNameToNumber(rangeBounds[1]);
            }

            // Only rows case.
            else if (numbers && !letters)
            {
                gridRange.StartRowIndex = Convert.ToInt32(rangeBounds[0]) - 1;
                gridRange.EndRowIndex = Convert.ToInt32(rangeBounds[1]);
            }

            // Coordinate case.
            // failed in A5:A type case that refers to all the cells of the first column, from row 5 onward.
            else
            {
                var firstColName = Regex.Match(rangeBounds[0], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                gridRange.StartColumnIndex = ColumnNameToNumber(firstColName) - 1;
                gridRange.StartRowIndex = Convert.ToInt32(Regex.Match(rangeBounds[0], @"\d+").Value) - 1;

                if (rangeBounds.Length > 1)
                {
                    var secondColName = Regex.Match(rangeBounds[1], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                    gridRange.EndColumnIndex = ColumnNameToNumber(secondColName);
                    gridRange.EndRowIndex = Convert.ToInt32(Regex.Match(rangeBounds[1], @"\d+").Value);
                }

                // Only one cell needed.
                else
                {
                    gridRange.EndColumnIndex = gridRange.StartColumnIndex + 1;
                    gridRange.EndRowIndex = gridRange.StartRowIndex + 1;
                }
            }

            return gridRange;
        }

        /// <summary>
        /// Convert column letter to index.
        /// </summary>
        /// <param name="columnName">Column name.</param>
        /// <returns>Column number for column name.</returns>
        private int ColumnNameToNumber(string columnName)
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
