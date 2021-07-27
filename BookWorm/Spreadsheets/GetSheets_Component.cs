﻿using BookWorm.Utilities;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Spreadsheets
{
    /// <summary>
    /// Component to get sheet names from a speadsheet.
    /// </summary>
    public class GetSheets_Component : ReadWriteBaseComponent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetSheets_Component"/> class.
        /// </summary>
        public GetSheets_Component()
          : base(
                "Get Sheet Names",
                "GetSheets",
                "Get sheet names from a spreadsheet",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Sheet Names", "N", "Names of all spreadsheet sheets", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetUrl = string.Empty;
            bool run = false;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;
            SpreadsheetId = Util.ParseUrl(spreadsheetUrl);

            DA.GetData(1, ref run);
            if (!run) return;

            var sheetIds = SheetsUtilities.GetAllSheetNames(SpreadsheetId);

            DA.SetDataList(0, sheetIds);
        }

        /// <inheritdoc/>
        public override Guid ComponentGuid
        {
            get { return new Guid("f14b2e40-23a4-453f-9753-607e2a996d0e"); }
        }
    }
}