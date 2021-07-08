using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using BookWorm.Goo;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Data = Google.Apis.Sheets.v4.Data;

namespace BookWorm.Spreadsheets
{
    public class WriteCells : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WriteCells"/> class.
        /// </summary>
        public WriteCells()
          : base(
                "WriteCells",
                "WriteCell",
                "Updates values and format for a given range of cells",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddGenericParameter("Cells", "C", "Cells", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Run", "R", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            var cellsGoo = new List<GH_CellData>();
            var gridRange = new GridRange();
            var rows = new List<RowData>();

            string fieldMask = string.Empty;

            var run = false;

            if (!DA.GetDataList("Cells", cellsGoo)) return;

            DA.GetData("Run", ref run);

            if (!run)
            {
                return;
            }

            var cells = cellsGoo.Select(c => c.Value).ToList();

            var sheetId = SheetsUtilities.GetSheetId(SpreadsheetId, SheetName);

            if (sheetId == null)
            {
                sheetId = SheetsUtilities.CreateNewSheet(SpreadsheetId, SheetName);
            }

            gridRange = CellsUtilities.GridRangeFromA1(this.Range, (int)sheetId);

            // Evenly distribute cells in rows range and set bounds explicitly.
            gridRange.FitRangeToCells(cells.Count);

            rows = CellsUtilities.GetRows(cells, gridRange);

            // A list of updates to apply to the spreadsheet.
            // Requests will be applied in the order they are specified.
            // If any request is not valid, no requests will be applied.
            var requests = new List<Data.Request>();

            // New and empty request instance.
            var updateCellRequest = new Data.Request();

            var updateCells = new UpdateCellsRequest
            {
                Rows = rows,
                Fields = "*",
                Range = gridRange,
            };

            updateCellRequest.UpdateCells = updateCells;

            requests.Add(updateCellRequest);

            // Main request of matrioshka-request.
            var requestBody = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests,
            };

            var request = Credentials.Service.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId);

            var response = request.Execute();
        }

        /// <inheritdoc/>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <inheritdoc/>
        public override Guid ComponentGuid
        {
            get { return new Guid("e0bc3449-e003-4bf2-aefa-fffafa239645"); }
        }
    }
}