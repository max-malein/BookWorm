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
            pManager.AddGenericParameter("BorderTop", "Top", "Top border of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderBottom", "Bottom", "Bottom border of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderLeft", "Left", "Left border of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderRight", "Right", "Right border of the cell", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Borders", "Borders", "The borders of the cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
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

        /// <inheritdoc/>
        protected override System.Drawing.Bitmap Icon =>

                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                null;

        /// <inheritdoc/>
        public override Guid ComponentGuid => new Guid("ee97311d-ecea-4261-9ef4-d17c0b125bfb");
    }
}