using BookWorm.Goo;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Deconstruct
{
    public class DeconstructNumberFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the DeconstructNumberFormat class.
        /// </summary>
        public DeconstructNumberFormat()
          : base("DeconstructNumberFormat", "DecNumberFormat",
              "Deconstruct number format",
              "BookWorm", "Deconstruct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("NumberFormat", "NumberFormat", "NumberFormat", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("NumberFormatType", "NumberFormatType", "The number format is not specified and is based on the contents of the cell. Do not explicitly use this. 0-TEXT, 1-NUMBER, 2-PERCENT, 3-CURRENCY, 4-DATE, 5-TIME, 6-DATE_TIME, 7-SCIENTIFIC", GH_ParamAccess.item);
            pManager.AddTextParameter("Pattern", "Pattern", "Pattern", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var numberFormatGoo = new GH_NumberFormat();
            if (!DA.GetData(0, ref numberFormatGoo) || numberFormatGoo.Value == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in NumberFormat input");
                return;
            }

            var numberFormat = numberFormatGoo.Value;

            DA.SetData(0, numberFormat.Type);
            DA.SetData(1, numberFormat.Pattern);
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
            get { return new Guid("6d026e13-0202-4e84-9f54-2af001827b12"); }
        }
    }
}