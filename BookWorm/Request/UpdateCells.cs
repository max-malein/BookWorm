using System;
using System.Collections.Generic;
using System.Linq;
using BookWorm.Goo;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Data = Google.Apis.Sheets.v4.Data;

namespace BookWorm.Request
{
    public class UpdateCells : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the UpdateCells class.
        /// </summary>
        public UpdateCells()
          : base(
                "Update Cells",
                "Nickname",
                "Description",
                "BookWorm",
                "Request")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Id", "Id", "Spreadsheet Id or spreadsheet url", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet name. If such sheet doesn't exist it will be created.", GH_ParamAccess.item);

            pManager.AddGenericParameter("Cells", "C", "Cells in a Rows", GH_ParamAccess.list);

            pManager.AddTextParameter("Range", "Rng", "The range to write data to in A1 notation", GH_ParamAccess.item);

            pManager.AddTextParameter(
                "Fields",
                "F",
                "Field Mask that applies for cells update.\n" +
                "The root is the CellData; 'row.values.' should not be specified.\n" +
                "\"*\" can be used as \"every field\" but wildrard syntax can produce unwanted results if the API is updated in the future.",
                GH_ParamAccess.item,
                "*");

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var spreadsheetId = string.Empty;
            var sheetName = string.Empty;

            var cellsGoo = new List<GH_CellData>();

            var a1NotatonRange = string.Empty;
            var gridRange = new GridRange();
            var rows = new List<RowData>();

            string fieldMask = string.Empty;

            var run = false;

            if (!DA.GetData(0, ref spreadsheetId)) return;

            if (!DA.GetData(1, ref sheetName)) return;

            if (!DA.GetDataList(2, cellsGoo)) return;

            if (!DA.GetData(3, ref a1NotatonRange)) return;

            DA.GetData(4, ref fieldMask);

            DA.GetData(5, ref run);

            if (!run)
            {
                return;
            }

            var sheetUtilities = new SheetsUtilities();
            var cellUtilities = new CellsUtilities();

            var cells = cellsGoo.Select(c => c.Value).ToList();

            var sheetId = sheetUtilities.GetSheetId(spreadsheetId, sheetName);

            if (sheetId == null)
            {
                sheetId = sheetUtilities.CreateNewSheet(spreadsheetId, sheetName);
            }

            gridRange = cellUtilities.GridRangeFromA1(a1NotatonRange, (int)sheetId);

            // Evenly distribute cells in rows range and set bounds explicitly.
            gridRange.FitRangeToCells(cells.Count);

            rows = cellUtilities.GetRows(cells, gridRange);

            // A list of updates to apply to the spreadsheet.
            // Requests will be applied in the order they are specified.
            // If any request is not valid, no requests will be applied.
            var requests = new List<Data.Request>();

            // New and empty request instance.
            var updateCellRequest = new Data.Request();

            var updateCells = new UpdateCellsRequest
            {
                Rows = rows,
                Fields = fieldMask,
                Range = gridRange,
            };

            updateCellRequest.UpdateCells = updateCells;

            requests.Add(updateCellRequest);

            // Main request of matrioshka-request.
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

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