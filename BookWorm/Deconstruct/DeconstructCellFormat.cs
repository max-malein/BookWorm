namespace BookWorm.Deconstruct
{
    using System;
    using BookWorm.Goo;
    using BookWorm.Utilities;
    using Grasshopper.Kernel;

    public class DeconstructCellFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeconstructCellFormat"/> class.
        /// Deconstruct cell format.
        /// </summary>
        public DeconstructCellFormat()
          : base(
                "DeconstructCellFormat",
                "DecCellFormat",
                "Deconstruct cell format",
                "BookWorm",
                "Deconstruct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("CellFormat", "CellFormat", "Cell format", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Number Format", "NumberFormat", "The number format of a cell", GH_ParamAccess.item);

            pManager.AddTextParameter("Background color", "Color", "Background color of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("Borders", "Borders", "The borders of the cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("Padding", "Padding", "Padding around the cell, in pixels", GH_ParamAccess.item);

            pManager.AddIntegerParameter(
                "Horizontal Alignment",
                "HorizontalAlign",
                "0 - LEFT\n"
                + "1 - CENTER\n"
                + "2 - RIGHT",
                GH_ParamAccess.item);

            pManager.AddIntegerParameter(
                "Vertical Alignment",
                "VerticalAlign",
                "0 - TOP\n"
                + "1 - MIDDLE\n"
                + "2 - BOTTOM",
                GH_ParamAccess.item);

            pManager.AddIntegerParameter(
                "Wrap",
                "Wrap",
                "0 - OVERFLOW_CELL\n"
                + "1 - LEGACY_WRAP\n"
                + "2 - CLIP\n"
                + "3 - WRAP",
                GH_ParamAccess.item);

            pManager.AddIntegerParameter(
                "Text Direction",
                "TextDirection",
                "0 - LEFT_TO_RIGHT\n"
                + "1 - RIGHT_TO_LEFT",
                GH_ParamAccess.item);

            pManager.AddGenericParameter("Text Format", "TextFormat", "The format of a run of text in a cell", GH_ParamAccess.item);

            pManager.AddGenericParameter("Text Rotation", "TextRotation", "The rotation applied to text in a cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var cellFormatGoo = new GH_CellFormat();
            if (!DA.GetData(0, ref cellFormatGoo) || cellFormatGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in CellFormat input");
                return;
            }

            var cellFormat = cellFormatGoo.Value;

            var formattedColor = string.Empty;
            if (cellFormat.BackgroundColor != null)
            {
                var color = SheetsUtilities.GetSystemDrawingColor(cellFormat.BackgroundColor);
                formattedColor = SheetsUtilities.GetFormattedARGB(color);
            }

            object horAlignType = null;
            if (cellFormat.HorizontalAlignment != null)
            {
                horAlignType = Enum.Parse(typeof(Construct.ConstructCellFormat.HorizontalAlign), cellFormat.HorizontalAlignment);
            }

            object vertAlignType = null;
            if (cellFormat.VerticalAlignment != null)
            {
                vertAlignType = Enum.Parse(typeof(Construct.ConstructCellFormat.VerticalAlign), cellFormat.VerticalAlignment);
            }

            object wrapType = null;
            if (cellFormat.WrapStrategy != null)
            {
                wrapType = Enum.Parse(typeof(Construct.ConstructCellFormat.WrapStrategy), cellFormat.WrapStrategy);
            }

            object textDirType = null;
            if (cellFormat.TextDirection != null)
            {
                textDirType = Enum.Parse(typeof(Construct.ConstructCellFormat.TextDirection), cellFormat.TextDirection);
            }

            DA.SetData(0, new GH_NumberFormat(cellFormat.NumberFormat));

            DA.SetData(1, formattedColor);

            DA.SetData(2, new GH_CellBorders(cellFormat.Borders));

            DA.SetData(3, new GH_Padding(cellFormat.Padding));

            DA.SetData(4, horAlignType);

            DA.SetData(5, vertAlignType);

            DA.SetData(6, wrapType);

            DA.SetData(7, textDirType);

            DA.SetData(8, new GH_TextFormat(cellFormat.TextFormat));

            DA.SetData(9, new GH_TextRotation(cellFormat.TextRotation));
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
            get { return new Guid("e3cef773-f955-4a81-95ee-ba35a9e0c495"); }
        }
    }
}