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
        [TestCase("Sheet!A1", 0, 1, 0, 1, TestName = "Single cell")]
        [TestCase("Sheet1!B2:AA4", 1, 26, 1, 3, TestName = "Range of cells")]
        [TestCase("My Sheet 1", 0, null, 0, null, TestName = "full sheet")]
        [TestCase("My Sheet 1!B:B", 1, 1, null, null, TestName = "single column")]
        public void GridRangeFromA1_CorrectValuesTest(
            string code,
            int? expectedColumnStart,
            int? expectedColumnEnd,
            int? expectedRowStart,
            int? expectedRowEnd)
        {
            var gridRange = CellsUtilities.GridRangeFromA1(code, 0);

            Assert.AreEqual(expectedColumnStart, gridRange.StartColumnIndex, "Column start index");
            Assert.AreEqual(expectedColumnEnd, gridRange.EndColumnIndex, "Column end index");
            Assert.AreEqual(expectedRowStart, gridRange.StartRowIndex, "Row start index");
            Assert.AreEqual(expectedRowEnd, gridRange.EndRowIndex, "Row end index");
        }
    }
}