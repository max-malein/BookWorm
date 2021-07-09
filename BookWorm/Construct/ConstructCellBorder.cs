namespace BookWorm.Construct
{
    using System;
    using BookWorm.Goo;
    using BookWorm.Utilities;
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel;

    /// <summary>
    /// Construct cell border.
    /// </summary>
    public class ConstructCellBorder : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructCellBorder"/> class.
        /// </summary>
        public ConstructCellBorder()
           : base(
                 "ConstructCellBorder",
                 "ConCellBorder",
                 "Construct cell border",
                 "BookWorm",
                 "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter(
                "Border Style",
                "Style",
                "0 - NONE (use for erase exist border)\n"
                + "1 - DOTTED\n"
                + "2 - DASHED\n"
                + "3 - SOLID\n"
                + "4 - SOLID_MEDIUM\n"
                + "5 - SOLID_THICK\n"
                + "6 - DOUBLE",
                GH_ParamAccess.item);

            pManager.AddColourParameter("Color", "Color", "Border color", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Border", "Border", "Border", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var style = 0;
            var color = System.Drawing.Color.Empty;
            var border = new Border();

            if (!DA.GetData(0, ref style) || (style < 0 || style > 6))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The style type {style} does not exist");
                return;
            }

            border.Style = Enum.GetName(typeof(BorderStyle), style);

            if (DA.GetData(1, ref color))
            {
                border.Color = SheetsUtilities.GetGoogleSheetsColor(color);
            }

            var borderGoo = new GH_CellBorder(border);
            DA.SetData(0, borderGoo);
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
        public override Guid ComponentGuid => new Guid("30b79fc6-f04c-4aac-9aaf-80a994e303e0");

        /// <summary>
        /// NONE - No border. Used only when updating a border in order to erase it.
        /// DOTTED - The border is dotted.
        /// DASHED - The border is dashed.
        /// SOLID - The border is a thin solid line.
        /// SOLID_MEDIUM - The border is a medium solid line.
        /// SOLID_THICK - The border is a thick solid line.
        /// DOUBLE - The border is two solid lines.
        /// </summary>
        private enum BorderStyle
        {
            NONE,
            DOTTED,
            DASHED,
            SOLID,
            SOLID_MEDIUM,
            SOLID_THICK,
            DOUBLE,
        }
    }
}