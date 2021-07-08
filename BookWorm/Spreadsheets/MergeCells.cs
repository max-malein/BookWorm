using System;
using System.Collections.Generic;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Data = Google.Apis.Sheets.v4.Data;

namespace BookWorm.Spreadsheets
{
    public class MergeCells : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the MergeCells class.
        /// </summary>
        public MergeCells()
          : base(
                "Merge Cells",
                "Nickname",
                "Description",
                "BookWorm",
                "Request")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            pManager.AddTextParameter(
                "Merge Type",
                "MT",
                "How the cells should be merged.\n\n"
                + "MERGE_ALL - сreate a single merge from the range.\n"
                + "MERGE_COLUMNS - сreate a merge for each column in the range.\n"
                + "MERGE_ROWS - сreate a merge for each row in the range.",
                GH_ParamAccess.item);

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

            var mergeType = string.Empty;

            var run = false;

            if (!DA.GetData(3, ref mergeType) || string.IsNullOrEmpty(mergeType)) return;

            DA.GetData(4, ref run);

            if (!run)
            {
                return;
            }


            var sheetId = SheetsUtilities.GetSheetId(SpreadsheetId, SheetName);
            var gridRange = CellsUtilities.GridRangeFromA1(Range, (int)sheetId);

            var requests = new List<Data.Request>();

            var mergeCellRequest = new Data.Request();

            var mergeCells = new MergeCellsRequest
            {
                Range = gridRange,
                MergeType = mergeType,
            };

            mergeCellRequest.MergeCells = mergeCells;

            requests.Add(mergeCellRequest);

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
            get { return new Guid("4be332a6-8069-4f99-aac9-d632dbf920d1"); }
        }
    }
}