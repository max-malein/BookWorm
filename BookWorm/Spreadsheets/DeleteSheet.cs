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
            base.RegisterInputParams(pManager);

            var unregParam = this.Params.Input[2];
            this.Params.UnregisterInputParameter(unregParam);

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

            var requests = new List<Request>();

            var deleteSheetRequest = new Request();

            var deleteSheet = new DeleteSheetRequest { SheetId = sheetId };

            deleteSheetRequest.DeleteSheet = deleteSheet;

            requests.Add(deleteSheetRequest);

            // Main request of matrioshka-request.
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

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
            get { return new Guid("fa2132f5-5623-450e-aae2-71f93e7d7227"); }
        }
    }
}