using System;
using System.Collections.Generic;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Spreadsheets
{
    public class DeleteSheet : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the DeleteSpreadsheet class.
        /// </summary>
        public DeleteSheet()
          : base(
                "Delete Sheet",
                "DelSheet",
                "Description",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // There is a conflict with base calling, so InputParams just rewrited.
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet Name", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // There is a conflict with base calling, so SolveInstance also just rewrited.
            string spreadsheetUrl = string.Empty;
            var sheetName = string.Empty;
            bool run = false;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;

            // Also there is matter point - delete sheet explicitly
            if (!DA.GetData(1, ref sheetName)) return;

            DA.GetData(2, ref run);

            if (!run)
            {
                return;
            }

            var spreadsheetId = Util.ParseUrl(spreadsheetUrl);
            var sheetId = SheetsUtilities.GetSheetId(spreadsheetId, sheetName);

            if (sheetId == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The sheet {sheetName} does not exist");
                return;
            }

            var requests = new List<Request>();

            var deleteSheetRequest = new Request();

            var deleteSheet = new DeleteSheetRequest { SheetId = sheetId };

            deleteSheetRequest.DeleteSheet = deleteSheet;

            requests.Add(deleteSheetRequest);

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
            get { return new Guid("fa2132f5-5623-450e-aae2-71f93e7d7227"); }
        }
    }
}