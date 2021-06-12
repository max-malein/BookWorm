using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BookWorm.Spreadsheets
{
    /// <summary>
    /// 
    /// </summary>
    public class ReadCell_Component : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ReadCells class.
        /// </summary>
        public ReadCell_Component()
          : base(
                "Read Cell",
                "ReadCell",
                "Reads a range of cells",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Id", "Id", "Spreadsheet Id", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet Name", GH_ParamAccess.item);

            pManager.AddTextParameter("Cell Range", "C", "Range of cells", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Read", "R", "Read data from spreadsheet", GH_ParamAccess.item, false);

            // not exist yet. Just copypasted from old component as idea
            pManager.AddBooleanParameter("Same Length", "S", "All output rows will be the same length", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Cells", "C", "Cells in each row", GH_ParamAccess.tree);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetId = string.Empty;
            string sheetName = string.Empty;
            string range = string.Empty;
            bool read = false;
            bool sameLength = false;

            if (!DA.GetData(0, ref spreadsheetId))
            {
                return;
            }

            if (!DA.GetData(1, ref sheetName))
            {
                return;
            }

            if (!DA.GetData(2, ref range))
            {
                return;
            }

            DA.GetData(3, ref read);
            DA.GetData(4, ref sameLength);

            if (!read)
            {
                return;
            }

            // The range to retrieve from the spreadsheet.
            // Single quotes for cases with space between sheet name parts.
            string requestRange = $"'{sheetName}'!{range}";

            var request = Utilities.Credentials.Service.Spreadsheets.Get(spreadsheetId);
            request.Ranges = requestRange;
            request.IncludeGridData = true;

            var spreadsheet = request.Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault();

            if (sheet == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Can't read this sheet");
                return;
            }

            if (sheet.Properties.SheetType != "GRID")
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Sheet type is not \"GRID\"");
                return;
            }

            // Solver uses request range as item.
            var rowDataPerRequest = sheet.Data.Select(d => d.RowData.ToList()).ToList();
            var rowData = rowDataPerRequest[0];

            var outputGhCells = new GH_Structure<GH_CellData>();
            var runCountIndex = RunCount - 1;

            for (int i = 0; i < rowData.Count; i++)
            {
                var path = new GH_Path(runCountIndex, i);

                // Null-cell as the only cell in a row returns null-row, i.e. null instead of list of cells.
                // So you can get null-exeption if user requests column.
                var ghCells = rowData[i].Values?.Select(cd => new GH_CellData(cd)).ToList();

                // That stuff and "Values?" solve it.
                if (ghCells == null)
                {
                    ghCells = new List<GH_CellData>();
                }

                outputGhCells.AppendRange(ghCells, path);
            }

            DA.SetDataTree(0, outputGhCells);
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
            get { return new Guid("fca8e52a-8615-4d72-930c-cdf721e5e42e"); }
        }
    }
}