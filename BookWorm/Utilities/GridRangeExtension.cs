using System;
using Google.Apis.Sheets.v4.Data;

namespace BookWorm.Utilities
{
    public static class GridRangeExtension
    {
        /// <summary>
        /// Evenly distributes cells in rows.
        /// </summary>
        /// <param name="gridRange">Grid range.</param>
        /// <param name="cellsCount">Cells count.</param>
        /// <returns>Fitted range.</returns>
        public static GridRange FitRangeToCells(this GridRange gridRange, int cellsCount)
        {
            // Case SheetName
            if (gridRange.EndColumnIndex == null && gridRange.EndRowIndex == null)
            {
                gridRange.EndColumnIndex = (int)Math.Ceiling(Math.Sqrt(cellsCount));
                gridRange.EndRowIndex = gridRange.EndColumnIndex;
            }

            // Case A:B
            else if (gridRange.StartRowIndex == null && gridRange.EndRowIndex == null)
            {
                var colCount = gridRange.EndColumnIndex - gridRange.StartColumnIndex;
                gridRange.StartRowIndex = 0;
                gridRange.EndRowIndex = (int)Math.Ceiling((double)cellsCount / (int)colCount);
            }

            // Case A5:A
            else if (gridRange.EndRowIndex == null)
            {
                var colCount = gridRange.EndColumnIndex - gridRange.StartColumnIndex;
                gridRange.EndRowIndex = gridRange.StartRowIndex + (int)Math.Ceiling((double)cellsCount / (int)colCount);
            }

            // Case B5:5
            else if (gridRange.EndColumnIndex == null && gridRange.StartColumnIndex != null)
            {
                var rowCount = gridRange.EndRowIndex - gridRange.StartRowIndex;
                gridRange.EndColumnIndex = gridRange.StartColumnIndex + (int)Math.Ceiling((double)cellsCount / (int)rowCount);
            }

            // Case 2:4
            else if (gridRange.StartColumnIndex == null && gridRange.EndColumnIndex == null)
            {
                var rowCount = gridRange.EndRowIndex - gridRange.StartRowIndex;
                gridRange.StartColumnIndex = 0;
                gridRange.EndColumnIndex = (int)Math.Ceiling((double)cellsCount / (int)rowCount);
            }

            return gridRange;
        }
    }
}
