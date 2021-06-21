using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using System;

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
            pManager.AddGenericParameter("NumberFormat", "NumberFormat", "Number format", GH_ParamAccess.item);
            pManager.AddTextParameter("Color", "Color", "Color", GH_ParamAccess.item);
            pManager.AddGenericParameter("Borders", "Borders", "Borders", GH_ParamAccess.item);
            pManager.AddGenericParameter("Padding", "Padding", "Padding", GH_ParamAccess.item);
            pManager.AddTextParameter("HorizontalAligment", "HorizontalAligment", "Horizontal aligment", GH_ParamAccess.item);
            pManager.AddTextParameter("VerticalAligment", "VerticalAligment", "Vertical aligment", GH_ParamAccess.item);
            pManager.AddTextParameter("Wrap", "Wrap", "Wrap", GH_ParamAccess.item);
            pManager.AddTextParameter("TextDirection", "TextDirection", "Text direction", GH_ParamAccess.item);
            pManager.AddGenericParameter("TextFormat", "TextFormat", "Text format", GH_ParamAccess.item);
            pManager.AddGenericParameter("TextRotation", "TextRotation", "Text rotation", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("CellFormat", "CellFormat", "Cell format", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var numberFormat = new GH_NumberFormat();
            var color = string.Empty;
            var borders = new GH_CellBorders();
            var padding = new GH_Padding();
            var horizontalAligment = string.Empty;
            var verticalAligment = string.Empty;
            var wrap = string.Empty;
            var textDirection = string.Empty;
            var textFormat = new GH_TextFormat();
            var textRotation = new GH_TextRotation();

            var cellFormat = new CellFormat();

            if (DA.GetData(0, ref numberFormat))
            {
                cellFormat.NumberFormat = numberFormat.Value;
            }

            if (DA.GetData(1, ref color))
            {
                cellFormat.BackgroundColor = new Color
                {
                    Alpha = 255,
                    Red = float.Parse(color.Split(',')[0]),
                    Green = float.Parse(color.Split(',')[1]),
                    Blue = float.Parse(color.Split(',')[2]),
                };
            }

            if (DA.GetData(2, ref borders))
            {
                cellFormat.Borders = borders.Value;
            }

            if (DA.GetData(3, ref padding))
            {
                cellFormat.Padding = padding.Value;
            }

            if (DA.GetData(4, ref horizontalAligment))
            {
                cellFormat.HorizontalAlignment = horizontalAligment;
            }

            if (DA.GetData(5, ref verticalAligment))
            {
                cellFormat.VerticalAlignment = verticalAligment;
            }

            if (DA.GetData(6, ref wrap))
            {
                cellFormat.WrapStrategy = wrap;
            }

            if (DA.GetData(7, ref textDirection))
            {
                cellFormat.TextDirection = textDirection;
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
            get { return new Guid("4d32bb2d-ef10-4687-803f-e37bde2ae27a"); }
        }
    }
}