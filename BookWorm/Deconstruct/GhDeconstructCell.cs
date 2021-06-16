using System;
using BookWorm.Goo;
using Grasshopper.Kernel;

namespace BookWorm.Deconstruct
{
    /// <summary>
    /// Deconstruct GoogleSheets CellData by some of it's component parts.
    /// </summary>
    public class GhDeconstructCell : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GhDeconstructCell class.
        /// </summary>
        public GhDeconstructCell()
          : base(
                "Deconstruct Cell",
                "DeCell",
                "Deconstruct cell by some of it's component parts",
                "BookWorm",
                "Deconstruct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Cell", "C", "Data about the specific cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Formatted Value", "FV", "Formatted value of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter(
                "User Entered Value",
                "UV",
                "The value the user entered in the cell. e.g, 1234 , 'Hello' , or =NOW()."
                + "\n Note: Dates, Times and DateTimes are represented as doubles in serial number format.",
                GH_ParamAccess.item);

            pManager.AddGenericParameter("Cell Format", "CF", "The format of the cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GH_CellData cellDataGoo = null;

            if (!DA.GetData(0, ref cellDataGoo) || !cellDataGoo.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in Cell input");
                return;
            }

            var cellData = cellDataGoo.Value;

            var formattedValue = cellData.FormattedValue;
            var userEnteredValue = new GH_ExtendedValue(cellData.UserEnteredValue);
            var cellFormat = new GH_CellFormat(cellData.UserEnteredFormat);

            DA.SetData(0, formattedValue);
            DA.SetData(1, userEnteredValue);
            DA.SetData(2, cellFormat);
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
            get { return new Guid("a5f68522-f13d-42c9-a855-4b1d01a63da8"); }
        }
    }
}