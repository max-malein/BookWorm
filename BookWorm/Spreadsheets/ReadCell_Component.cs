using BookWorm.Goo;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using System;
using System.Collections.Generic;
using System.Drawing;
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

            pManager.AddBooleanParameter("Duplicate Merged", "M", "If true, all merged cells will have same value. If false, only the upper left cell will have a value", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Same Length", "S", "All output rows will be the same length", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Read", "R", "Read data from spreadsheet", GH_ParamAccess.item, false);
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

            bool duplicatedMerged;
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
            request.Fields = "sheets(properties.sheetType,data,merges)";
            //request.IncludeGridData = true;

            var spreadsheet = request.Execute();

            // Solver uses request range as item. So you always will get only one sheet and so on.
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

            // Мерж ренжи
            List<GridRange> merges = new List<GridRange>();

            if (sheet.Merges != null)
            {
                merges = sheet.Merges.ToList();
            }

            var gridData = sheet.Data.FirstOrDefault();

            // Rows and cells as "null-data-null-data-null" actually becomes "null-data-null-data" in response.
            // Merged null-cells are considered as non-null-cells.
            var rowsData = gridData.RowData;

            // if one tried read a range of null-cells
            if (rowsData == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The range contains only null-cells");
                return;
            }

            var startRowInd = Convert.ToInt32(gridData.StartRow);
            var startColumnInd = Convert.ToInt32(gridData.StartColumn);



            var coord = CellsUtilities.GetCellCoordinates(rowsData, startRowInd, startColumnInd);



            var outputGhCells = new GH_Structure<GH_CellData>();
            var runCountIndex = RunCount - 1;


            for (int i = 0; i < rowsData.Count; i++)
            {
                var path = new GH_Path(runCountIndex, i);

                var ghCells = new List<GH_CellData>();

                // If row doesn't contain any cell with valid data.
                if (rowsData[i].Values == null)
                {
                    outputGhCells.AppendRange(ghCells, path);
                    continue;
                }

                for (int j = 0; j < rowsData[i].Values.Count; j++)
                {
                    CellData value = null;

                    //if (duplicatedMerged)
                    //{
                    //    Point? mergeOrigin = FindMergeOrigin(coord[i][j].Coordinates, mergeData);

                    //    if (mergeOrigin != null)
                    //    {
                    //        value = coord.SelectMany(r => r.Where(v => v.Coordinates == mergeOrigin)).FirstOrDefault().Value; // needs to be tested
                    //    }
                    //    else
                    //    {
                    //        value = rowsData[i].Values[j];
                    //    }
                    //}
                    //else
                    //{
                    value = rowsData[i].Values[j];
                    //}

                    var ghCell = new GH_CellData(value);
                    ghCells.Add(ghCell);
                }

                outputGhCells.AppendRange(ghCells, path);
            }
            
            DA.SetDataTree(0, outputGhCells);
        }

        private Point? FindMergeOrigin(Point? coordinates, List<GridRange> mergeData)
        {
            var x = coordinates.Value.X;
            var e = mergeData.Where(md => md.StartColumnIndex <= x && md.EndColumnIndex - 1 >= x).ToList();
            if (e.Count == 0)
                return null;
            var y = coordinates.Value.Y;
            var f = e.Where(md => md.StartRowIndex <= y && md.EndRowIndex - 1 >= y).ToList();
            if (f.Count == 0)
                return null;
            return new Point?(new Point(f[0].StartColumnIndex.Value, f[0].StartRowIndex.Value));
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