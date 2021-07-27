using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Utilities
{
    public class GH_GetSheets : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GH_GetSheets()
          : base(
                "Get Sheets",
                "GetSheets",
                "Get all sheets within a spreadsheed",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Sheet Ids", "SId", "Ids of sheets contained in the spreadsheet", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetUrl = string.Empty;
            bool run = false;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;

            DA.GetData(1, ref run);
            if (!run) return;

            var spreadsheetId = Util.ParseUrl(spreadsheetUrl);
            var sheetIds = SheetsUtilities.GetAllSheetNames(spreadsheetId);

            DA.SetDataList(0, sheetIds);

        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("f14b2e40-23a4-453f-9753-607e2a996d0e"); }
        }
    }
}