using Google.Apis.Sheets.v4.Data;
using System.Drawing;

namespace BookWorm.Utilities
{
    internal class CellWithCoordinates
    {
        public CellData Value { get; set; }
        public Point? Coordinates { get; set; }
    }
}