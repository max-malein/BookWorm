using System;
using System.Collections.Generic;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Data = Google.Apis.Sheets.v4.Data;

namespace BookWorm.Request
{
    public class UnmergeCells : GH_Component
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
                "Request")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Id", "Id", "Spreadsheet Id or spreadsheet url", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet name", GH_ParamAccess.item);

            pManager.AddTextParameter("Range", "Rng", "The range of cells to unmerge in A1 notation", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var spreadsheetId = string.Empty;
            var sheetName = string.Empty;

            var a1NotatonRange = string.Empty;

            var run = false;

            if (!DA.GetData(0, ref spreadsheetId) || string.IsNullOrEmpty(spreadsheetId)) return;

            if (!DA.GetData(1, ref sheetName) || string.IsNullOrEmpty(sheetName)) return;

            if (!DA.GetData(2, ref a1NotatonRange) || string.IsNullOrEmpty(a1NotatonRange)) return;

            DA.GetData(3, ref run);

            if (!run)
            {
                return;
            }

            var sheetUtilities = new SheetsUtilities();
            var cellUtilities = new CellsUtilities();

            var sheetId = sheetUtilities.GetSheetId(spreadsheetId, sheetName);
            var gridRange = cellUtilities.GridRangeFromA1(a1NotatonRange, (int)sheetId);

            var requests = new List<Data.Request>();

            var unmergeCellRequest = new Data.Request();

            var unmergeCells = new UnmergeCellsRequest
            {
                Range = gridRange,
            };

            unmergeCellRequest.UnmergeCells = unmergeCells;

            requests.Add(unmergeCellRequest);

            // Main request of matrioshka-request.
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("00437ebd-fde8-4db1-80af-ae511a04b51e"); }
        }
    }
}