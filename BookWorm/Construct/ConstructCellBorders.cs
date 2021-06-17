// <copyright file="ConstructCellBorders.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Construct
{
    using System;
    using BookWorm.Goo;
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel;

    /// <summary>
    /// Construct cell borders.
    /// </summary>
    public class ConstructCellBorders : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructCellBorders"/> class.
        /// </summary>
        public ConstructCellBorders()
          : base(
                "ConstructCellBorders",
                "ConCellBorders",
                "Construct cell borders",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("BorderTop", "BTop", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderBottom", "BBottom", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderLeft", "BLeft", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderRight", "BRight", "Border", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Borders", "Borders", "Borders", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var borders = new Borders();

            var bTop = new GH_CellBorder();
            var bBottom = new GH_CellBorder();
            var bLeft = new GH_CellBorder();
            var bRight = new GH_CellBorder();

            if (DA.GetData(0, ref bTop))
            {
                borders.Top = bTop.Value;
            }

            if (DA.GetData(1, ref bBottom))
            {
                borders.Bottom = bBottom.Value;
            }

            if (DA.GetData(2, ref bLeft))
            {
                borders.Left = bLeft.Value;
            }

            if (DA.GetData(3, ref bRight))
            {
                borders.Right = bRight.Value;
            }

            var borderGoo = new GH_CellBorders(borders);
            DA.SetData(0, borderGoo);
        }

        /// <summary>
        /// Gets provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon =>

                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                null;

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("ee97311d-ecea-4261-9ef4-d17c0b125bfb");
    }
}