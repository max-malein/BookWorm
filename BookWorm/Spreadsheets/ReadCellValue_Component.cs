using System;
using System.Collections.Generic;
using System.Linq;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace GoogleDocs.Spreadsheets
{
    public class ReadCellValue_Component : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReadCellValue_Component"/> class.
        /// </summary>
        public ReadCellValue_Component()
          : base(
                "Read Cell Value",
                "ReadCellValue",
                "Basic reader from a range of cells. For a more comprehesive solution please use ReadCells component.",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            pManager.AddBooleanParameter(
                "Duplicate Merged",
                "M",
                "If true, all merged cells will have same value. If false, only the upper left cell will have a value",
                GH_ParamAccess.item,
                false);

            pManager.AddBooleanParameter("Read", "R", "Read data from spreadsheet", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Same Length", "S", "All output rows will be the same length", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Values", "V", "Values", GH_ParamAccess.tree);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            bool duplicateMerged = false;
            bool read = false;
            bool sameLength = false;

            DA.GetData("Duplicate Merged", ref duplicateMerged);
            DA.GetData("Read", ref read);
            DA.GetData("Same Length", ref sameLength);

            if (!read) return;

            var request = Credentials.Service.Spreadsheets.Values.Get(SpreadsheetId, SpreadsheetRange);
            ValueRange response = request.Execute();

            IList<IList<object>> values = response.Values;

            var merges = new List<GridRange>();
            var gridRange = new GridRange();

            if (duplicateMerged)
            {
                merges = SheetsUtilities.GetMergeRanges(SpreadsheetId, SpreadsheetRange);
                gridRange = CellsUtilities.GridRangeFromA1(CellRange, 0);
            }

            if (values != null && values.Count > 0)
            {
                var data = new GH_Structure<GH_String>();

                for (int i = 0; i < values.Count; i++)
                {
                    //for (int j = 0; j < values[i].Count; j++)
                    //{
                    //    if (duplicateMerged)
                    //    {

                    //    }
                    //    else
                    //    {
                            var path = new GH_Path(RunCount - 1, i);
                            var ghStrings = values[i].Select(s => new GH_String(s.ToString()));

                            data.AppendRange(ghStrings, path);
                    //    }
                    //}
                    
                }

                // same lenght for each row.
                if (sameLength)
                {
                    var maxLength = values.Max(r => r.Count);
                    foreach (var path in data.Paths)
                    {
                        int valuesToAdd = maxLength - data[path].Count;
                        if (valuesToAdd > 0)
                        {
                            var emptys = Enumerable.Repeat(new GH_String(string.Empty), valuesToAdd);
                            data[path].AddRange(emptys);
                        }
                    }
                }

                DA.SetDataTree(0, data);
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No data found.");
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

        /// <summary>
        /// Each component must have a unique Guid to identify it.
        /// It is vital this Guid doesn't change otherwise old ghx files
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5042202c-269f-4617-8b7c-fcb847d56ff6"); }
        }
    }
}
