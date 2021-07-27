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
        [TestCase("", 0, null, 0, null, TestName = "Full Sheet")]
        [TestCase("A1", 0, 1, 0, 1, TestName = "Single cell")]
        [TestCase("B2:AA4", 1, 27, 1, 4, TestName = "Range of cells")]
        [TestCase("B:B", 1, 2, null, null, TestName = "Endless column from 0")]
        [TestCase("3:3", null, null, 2, 3, TestName = "Endless row from 0")]
        [TestCase("B:e", 1, 5, null, null, TestName = "Endless columns from 0")]
        [TestCase("3:5", null, null, 2, 5, TestName = "Endless rows from 0")]
        [TestCase("B1:B", 1, 2, 0, null, TestName = "Endless column from 1")]
        [TestCase("C3:3", 2, null, 2, 3, TestName ="Endless row from C")]
    
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