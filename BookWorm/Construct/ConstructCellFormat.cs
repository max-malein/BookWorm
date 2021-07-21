using System;
using BookWorm.Goo;
using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Construct
{
    public class ConstructCellFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructCellFormat"/> class.
        /// </summary>
        public ConstructCellFormat()
          : base(
                "ConstructCellFormat",
                "ConCellFormat",
                "Construct cell format",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Number Format", "NumberFormat", "The number format of a cell", GH_ParamAccess.item);

            pManager.AddColourParameter("Background color", "Color", "Background color of the cell", GH_ParamAccess.item);

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

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Cell Format", "CellFormat", "The format of a cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var numberFormat = new GH_NumberFormat();

            var colorARGB = System.Drawing.Color.Empty;

            var borders = new GH_CellBorders();

            var padding = new GH_Padding();

            var horizontalAlign = 0;
            var verticalAlign = 0;

            var wrap = 0;

            var textDirection = 0;
            var textFormat = new GH_TextFormat();
            var textRotation = new GH_TextRotation();

            var cellFormat = new CellFormat();

            if (DA.GetData(0, ref numberFormat))
            {
                cellFormat.NumberFormat = numberFormat.Value;
            }

            if (DA.GetData(1, ref colorARGB))
            {
                cellFormat.BackgroundColor = SheetsUtilities.GetGoogleSheetsColor(colorARGB);
            }

            if (DA.GetData(2, ref borders))
            {
                cellFormat.Borders = borders.Value;
            }

            if (DA.GetData(3, ref padding))
            {
                cellFormat.Padding = padding.Value;
            }

            if (DA.GetData(4, ref horizontalAlign))
            {
                if (horizontalAlign < 0 || horizontalAlign > 2)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The horizontal align type {horizontalAlign} does not exist");
                    return;
                }

                cellFormat.HorizontalAlignment = Enum.GetName(typeof(HorizontalAlign), horizontalAlign);
            }

            if (DA.GetData(5, ref verticalAlign))
            {
                if (verticalAlign < 0 || verticalAlign > 2)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The vertical align type {verticalAlign} does not exist");
                    return;
                }

                cellFormat.VerticalAlignment = Enum.GetName(typeof(VerticalAlign), verticalAlign);
            }

            if (DA.GetData(6, ref wrap))
            {
                if (wrap < 0 || wrap > 3)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The wrap strategy type {wrap} does not exist");
                    return;
                }

                cellFormat.WrapStrategy = Enum.GetName(typeof(WrapStrategy), wrap);
            }

            if (DA.GetData(7, ref textDirection))
            {
                if (textDirection < 0 || textDirection > 1)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The text direction type {textDirection} does not exist");
                    return;
                }

                cellFormat.TextDirection = Enum.GetName(typeof(TextDirection), textDirection);
            }

            if (DA.GetData(8, ref textFormat))
            {
                cellFormat.TextFormat = textFormat.Value;
            }

            if (DA.GetData(9, ref textRotation))
            {
                cellFormat.TextRotation = textRotation.Value;
            }

            var cellFormatGoo = new GH_CellFormat(cellFormat);
            DA.SetData(0, cellFormatGoo);
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
            get { return new Guid("4d32bb2d-ef10-4687-803f-e37bde2ae27a"); }
        }

        /// <summary>
        /// The horizontal alignment of text in a cell.
        /// </summary>
        private enum HorizontalAlign
        {
            LEFT,
            CENTER,
            RIGHT,
        }

        /// <summary>
        /// The vertical alignment of text in a cell.
        /// </summary>
        private enum VerticalAlign
        {
            TOP,
            MIDDLE,
            BOTTOM,
        }

        /// <summary>
        /// How to wrap text in a cell.
        /// </summary>
        private enum WrapStrategy
        {
            OVERFLOW_CELL,
            LEGACY_WRAP,
            CLIP,
            WRAP,
        }

        /// <summary>
        /// The direction of text in a cell.
        /// </summary>
        private enum TextDirection
        {
            LEFT_TO_RIGHT,
            RIGHT_TO_LEFT,
        }
    }
}