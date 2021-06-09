using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Construct
{
    public class ConstructNumberFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ConstructNumberFormat class.
        /// </summary>
        public ConstructNumberFormat()
          : base(
                "ConstructPadding",
                "ConPadding",
                "Construct padding",
                "BookWorm",
                "Construct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("NumberFormatType", "NumberFormatType", "The number format is not specified and is based on the contents of the cell. Do not explicitly use this. 0-TEXT, 1-NUMBER, 2-PERCENT, 3-CURRENCY, 4-DATE, 5-TIME, 6-DATE_TIME, 7-SCIENTIFIC", GH_ParamAccess.item);
            pManager.AddTextParameter("Pattern", "Pattern", "Pattern", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("NumberFormat", "NumberFormat", "NumberFormat", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var numberFormatType = NumberFormatType.TEXT;
            var pattern = string.Empty;

            if (!DA.GetData(0, ref numberFormatType))
            {
                return;
            }
            if (!DA.GetData(1, ref pattern))
            {
                return;
            }

            var numberFormat = new NumberFormat()
            {
                Type = numberFormatType.ToString(),
                Pattern = pattern,
            };
            DA.SetData(0, numberFormat);
        }

        /// <summary>
        /// Provides an Icon for the component.
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
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ce00f900-34f0-4a49-86c0-1c8c45db9d65"); }
        }

        public enum NumberFormatType
        {
            TEXT,
            NUMBER,
            PERCENT,
            CURRENCY,
            DATE,
            TIME,
            DATE_TIME,
            SCIENTIFIC,
        }
    }
}