using System;
using System.Collections.Generic;
using System.Linq;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

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

            pManager.AddIntegerParameter(
                "Merge Types",
                "MT",
                "How the cells should be merged.\n\n"
                + "0 - MERGE_ALL - сreate a single merge from the range.\n"
                + "1 - MERGE_COLUMNS - сreate a merge for each column in the range.\n"
                + "2 - MERGE_ROWS - сreate a merge for each row in the range.",
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

            var mergeTypes = new List<int>();

            var run = false;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;
            var spreadsheetId = Util.ParseUrl(spreadsheetUrl);

            if (!DA.GetData(1, ref sheetName)) return;

            // Check ranges
            if (!DA.GetDataList(2, ranges)) return;
            var cellRangesFormatted = ranges.Select(str => str.ToUpper()).ToList();

            var rangesCount = cellRangesFormatted.Count();

            // Check merge types
            if (DA.GetDataList(3, mergeTypes))
            {
                if (rangesCount != mergeTypes.Count)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Count of cell ranges and merge types must be equal");
                    return;
                }

                foreach (var mergeType in mergeTypes)
                {
                    if (mergeType < 0 || mergeType > 2)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The merge type {mergeType} does not exist");
                        return;
                    }
                }
            }

            DA.GetData(4, ref run);

            if (!run)
            {
                return;
            }

            var sheetId = SheetsUtilities.GetSheetId(spreadsheetId, sheetName);
            var gridRanges = new List<GridRange>();

            var requests = new List<Request>();

            for (int i = 0; i < rangesCount; i++)
            {
                var gridRange = CellsUtilities.GridRangeFromA1(cellRangesFormatted[i], (int)sheetId);
                gridRanges.Add(gridRange);

                var mergeCellRequest = new Request
                {
                    MergeCells = new MergeCellsRequest
                    {
                        Range = gridRange,
                        MergeType = Enum.GetName(typeof(MergeTypes), mergeTypes[i]),
                    },
                };

                requests.Add(mergeCellRequest);
            }

            // Main request of matrioshka-request.
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();
        }

        /// <summary>
        /// MERGE_ALL - сreate a single merge from the range.
        /// MERGE_COLUMNS - сreate a merge for each column in the range.
        /// MERGE_ROWS - сreate a merge for each row in the range.
        /// </summary>
        private enum MergeTypes
        {
            MERGE_ALL,
            MERGE_COLUMNS,
            MERGE_ROWS,
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