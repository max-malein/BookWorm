using System;
using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Construct
{
    public class ConstructPadding : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructPadding"/> class.
        /// </summary>
        public ConstructPadding()
          : base(
                "ConstructPadding",
                "ConPadding",
                "The amount of padding around the cell, in pixels",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Top", "Top", "Top padding in pixels", GH_ParamAccess.item, 0);
            pManager.AddIntegerParameter("Bottom", "Bottom", "Bottom padding in pixels", GH_ParamAccess.item, 0);
            pManager.AddIntegerParameter("Left", "Left", "Left padding in pixels", GH_ParamAccess.item, 0);
            pManager.AddIntegerParameter("Right", "Right", "Right padding in pixels", GH_ParamAccess.item, 0);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Padding", "Padding", "Paddings around the cell in pixels", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var padding = new Padding();

            var top = 0;
            var bottom = 0;
            var left = 0;
            var right = 0;

            // Possibilities for negative pixel value sounds funny but nothing bad will happen on sheets side.
            // Just strange results. Just as planned mb.
            if (DA.GetData(0, ref top))
            {
                padding.Top = top;
            }

            if (DA.GetData(1, ref bottom))
            {
                padding.Bottom = bottom;
            }

            if (DA.GetData(2, ref left))
            {
                padding.Left = left;
            }

            if (DA.GetData(3, ref right))
            {
                padding.Right = right;
            }

            var paddingGoo = new GH_Padding(padding);
            DA.SetData(0, paddingGoo);
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
            get { return new Guid("d427cece-1160-4db4-815b-fedc48030a08"); }
        }
    }
}