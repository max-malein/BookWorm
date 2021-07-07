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
            // Only columns case.
            // Works fine without that because row length depends on columns indicies (for GetRows),
            // so you don't actually need end row index. But explicit is better than implicit.
            if (gridRange.EndRowIndex == null)
            {
                var colCount = gridRange.EndColumnIndex - gridRange.StartColumnIndex;
                gridRange.StartRowIndex = 0;
                gridRange.EndRowIndex = (int)Math.Ceiling((double)cellsCount / (int)colCount);
            }

            // Only rows case.
            else if (gridRange.EndColumnIndex == null)
            {
                var rowCount = gridRange.EndRowIndex - gridRange.StartRowIndex;
                gridRange.StartColumnIndex = 0;
                gridRange.EndColumnIndex = (int)Math.Ceiling((double)cellsCount / (int)rowCount);
            }

            return gridRange;
        }
    }
}
