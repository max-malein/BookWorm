using System;
using System.Collections.Generic;
using System.Linq;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Spreadsheets
{
    public class UnmergeCells : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the UnmergeCells class.
        /// </summary>
        public UnmergeCells()
          : base(
                "Unmerge Cells",
                "Nickname",
                "Description",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet Name", GH_ParamAccess.item);

            pManager.AddTextParameter(
                "Cell Ranges",
                "CR",
                "Ranges of cells in \'a1\' notation. For example A1:B5 - range of cells, A15 - single cell, A:C - range of columns, etc.",
                GH_ParamAccess.list);

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetUrl = string.Empty;
            var sheetName = string.Empty;
            var ranges = new List<string>();

            var run = false;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;
            SpreadsheetId = Util.ParseUrl(spreadsheetUrl);

            if (!DA.GetData(1, ref sheetName)) return;

            // Check ranges
            if (!DA.GetDataList(2, ranges)) return;
            var cellRangesFormatted = ranges.Select(str => str.ToUpper()).ToList();

            DA.GetData(3, ref run);

            if (!run)
            {
                return;
            }

            var sheetId = SheetsUtilities.GetSheetId(SpreadsheetId, sheetName);

            if (sheetId == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The sheet {sheetName} does not exist");
                return;
            }

            var gridRanges = new List<GridRange>();

            var requests = new List<Request>();

            foreach (var cellRange in cellRangesFormatted)
            {
                var gridRange = CellsUtilities.GridRangeFromA1(cellRange, (int)sheetId);
                gridRanges.Add(gridRange);

                var unmergeCellRequest = new Request
                {
                    UnmergeCells = new UnmergeCellsRequest { Range = gridRange },
                };

                requests.Add(unmergeCellRequest);
            }

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
            get { return new Guid("00437ebd-fde8-4db1-80af-ae511a04b51e"); }
        }
    }
}