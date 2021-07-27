using System;
using System.Collections.Generic;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace GoogleDocs.Spreadsheets
{
    public class WriteCellValues : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WriteCellValues"/> class.
        /// </summary>
        public WriteCellValues()
          : base(
                "Write Cell Values",
                "WriteCellValues",
                "A very basic google spreadsheet writer. "
                + "Writes a rows of values, within a given range. "
                + "For a more comprehensive solution, please use WriteCell component",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            pManager.AddTextParameter(
                "Values",
                "V",
                "Null values will be skipped."
                + "\nTo set a cell to an empty value, set the string value to an empty string.",
                GH_ParamAccess.list);
            pManager.AddBooleanParameter("Append", "A", "If true, values will be added to the next possible empty row", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Result", "R", "Result", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            var inputData = new List<string>();
            bool write = false;
            bool append = false;

            if (!DA.GetDataList("Values", inputData)) return;
            DA.GetData("Append", ref append);
            DA.GetData("Run", ref write);

            if (!write) return;

            var gridRange = CellsUtilities.GridRangeFromA1(CellRange, 0);
            gridRange.FitRangeToCells(inputData.Count);

            var rows = CellsUtilities.GetRows(inputData, gridRange);

            var valueRange = new ValueRange
            {
                MajorDimension = "ROWS",
                Range = this.SpreadsheetRange,
                Values = rows.ToArray(),
            };

            if (append)
            {
                var request = Credentials.Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, SpreadsheetRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var result = request.Execute();
                DA.SetData(0, result.Updates.ToString());
            }
            else
            {
                var request = Credentials.Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, SpreadsheetRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var result = request.Execute();
                DA.SetData(0, result.ToString());
            }
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
            get { return new Guid("669A9C19-1370-4613-AC10-1D3264E3D290"); }
        }
    }
}
