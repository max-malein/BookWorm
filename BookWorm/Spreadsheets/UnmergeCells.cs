using System;
using System.Collections.Generic;
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
            base.RegisterInputParams(pManager);

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            var run = false;

            DA.GetData("Run", ref run);

            if (!run)
            {
                return;
            }

            var sheetId = SheetsUtilities.GetSheetId(SpreadsheetId, SheetName);
            var gridRange = CellsUtilities.GridRangeFromA1(Range, (int)sheetId);

            var requests = new List<Request>();

            var unmergeCellRequest = new Request();

            var unmergeCells = new UnmergeCellsRequest
            {
                Range = gridRange,
            };

            unmergeCellRequest.UnmergeCells = unmergeCells;

            requests.Add(unmergeCellRequest);

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