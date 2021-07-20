using NUnit.Framework;
using BookWorm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace BookWorm.Utilities.Tests
{
    [TestFixture()]
    public class CellsUtilitiesTests
    {
        [Test()]
        public void GridRangeFromA1Test()
        {
            var gridRange1 = CellsUtilities.GridRangeFromA1("A1", 0);
            GridRange correctGridRange = SetGridRange(0, 1, 0, 1);
            Assert.AreEqual(correctGridRange, gridRange1);
        }

        private GridRange SetGridRange(int columnStart, int columnEnd, int rowStart, int rowEnd)
        {
            return new GridRange()
            {
                StartColumnIndex = columnStart,
                EndColumnIndex = columnEnd,
                StartRowIndex = rowStart,
                EndRowIndex = rowEnd,
                SheetId = 0,
            };
        }
    }
}