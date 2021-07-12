using BookWorm.Goo;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace BookWorm.Spreadsheets
{
    /// <summary>
    /// 
    /// </summary>
    public class ReadCell_Component : ReadWriteBaseComponent
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadCell_Component"/> class.
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
            base.RegisterInputParams(pManager);

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
            base.SolveInstance(DA);

            bool read = false;
            bool sameLength = false;
            DA.GetData("Read", ref read);
            DA.GetData("Same Length", ref sameLength);

            if (!read) return;

            var request = Credentials.Service.Spreadsheets.Get(SpreadsheetId);
            request.Ranges = SpreadsheetRange;

            // Filter for quicker response.
            // If spreadsheet contains a lot of data you actually don't need data that you don't need.
            // If fields are set InclideGridData parameter is ignored.
            request.Fields = "sheets(properties.sheetType,data.rowData,merges)";
            //request.IncludeGridData = true;

            var spreadsheet = request.Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault();

            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Can't read this sheet");
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
            List<List<string>> a1s = CellsUtilities.GetCellCoordinates(rowData, CellRange);

            var outputGhCells = new GH_Structure<GH_CellData>();
            var runCountIndex = RunCount - 1;

            for (int i = 0; i < rowData.Count; i++)
            {
                var path = new GH_Path(runCountIndex, i);

                // Null-cell as the only cell in a row returns null-row, i.e. null instead of list of cells.
                // So you can get null-exeption if user requests column.
                //var ghCells = rowData[i].Values?.Select((cd, index) => new GH_CellData(cd, a1s[i, index] )).ToList();
                var ghCells = new List<GH_CellData>();
                for (int j = 0; j < rowData[i].Values.Count; j++)
                {
                    var value = rowData[i].Values[j];
                    var ghCell = new GH_CellData(value);
                    ghCells.Add(ghCell);
                }

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