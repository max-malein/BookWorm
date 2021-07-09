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
                "WriteCellValues",
                "WriteCellValues",
                "A very basic google spreadsheet writer. Writes a row of values, starting from a given cell index. For a more comprehensive solution, please use WriteCell component",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            pManager.AddTextParameter("Values", "V", "Values", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Append", "A", "If true, values will be added to the next possible empty row", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Write", "W", "Write data to spreadsheet", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Result", "R", "Result", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            var inputData = new List<string>();
            bool write = false;
            bool append = false;

            if (!DA.GetDataList("Values", inputData)) return;
            DA.GetData("Append", ref append);
            DA.GetData("Write", ref write);

            if (!write) return;

            var valueRange = new ValueRange
            {
                MajorDimension = "ROWS",
                Range = this.Range,
                Values = new List<IList<object>>() { new List<object>() },
            };

            // you need to explicitly set every value to string, otherwise doesn't work
            for (int i = 0; i < inputData.Count; i++)
            {
                valueRange.Values[0].Add(inputData[i]);
            }

            if (append)
            {
                var request = Credentials.Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, Range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var result = request.Execute();
                DA.SetData(0, result.Updates.ToString());
            }
            else
            {
                var request = Credentials.Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, Range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var result = request.Execute();
                DA.SetData(0, result.ToString());
            }
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
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
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("669A9C19-1370-4613-AC10-1D3264E3D290"); }
        }
    }
}
