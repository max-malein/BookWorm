using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using System;

namespace BookWorm.Construct
{
    public class ConstructPadding : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ConstructPadding class.
        /// </summary>
        public ConstructPadding()
          : base(
                "ConstructPadding",
                "ConPadding",
                "Construct padding",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Top", "Top", "Top", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Bottom", "Bottom", "Bottom", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Left", "Left", "Left", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Right", "Right", "Right", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Padding", "P", "Padding", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var padding = new Padding();

            var bTop = 0;
            var bBottom = 0;
            var bLeft = 0;
            var bRight = 0;

            if (DA.GetData(0, ref bTop))
            {
                padding.Top = bTop;
            }

            if (DA.GetData(1, ref bBottom))
            {
                padding.Bottom = bBottom;
            }

            if (DA.GetData(0, ref bLeft))
            {
                padding.Left = bLeft;
            }

            if (DA.GetData(1, ref bRight))
            {
                padding.Right = bRight;
            }

            DA.SetData(0, padding);
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
            get { return new Guid("d427cece-1160-4db4-815b-fedc48030a08"); }
        }
    }
}