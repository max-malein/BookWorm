using System;
using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Construct
{
    /// <summary>
    /// Construct GoogleSheets CellData from some of it's component parts.
    /// </summary>
    public class GhConstructCell : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GhConstructCell"/> class.
        /// </summary>
        public GhConstructCell()
          : base(
                "Construct Cell",
                "ConCell",
                "Construct cell from some of it's component parts",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(
                "User Entered Value",
                "UV",
                "The value the user entered in the cell.",
                GH_ParamAccess.item);

            pManager.AddGenericParameter("Cell Format", "CF", "The format of the cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Cell", "C", "Data about the specific cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GH_ExtendedValue userEnteredValueGoo = null;
            GH_CellFormat userCellFormatGoo = null;

            if (!DA.GetData(0, ref userEnteredValueGoo))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in UV input");
                return;
            }

            if (!DA.GetData(1, ref userCellFormatGoo))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in CF input");
                return;
            }

            var cellData = new CellData
            {
                UserEnteredValue = userEnteredValueGoo.Value,
                UserEnteredFormat = userCellFormatGoo.Value,
            };

            var cellDataGoo = new GH_CellData(cellData);

            DA.SetData(0, cellDataGoo);

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
            get { return new Guid("bd29230b-9399-4248-aae6-5ea70b922770"); }
        }
    }
}