using Google.Apis.Sheets.v4.Data;
using System.Drawing;

namespace BookWorm.Spreadsheets
{
    internal class CellCoordinates
    {
        public CellData Value { get; set; }
        public Point? Coordinates { get; set; }
    }
}