using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Construct
{
    public class ConstructCellFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ConstructCellFormat class.
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

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("NumberFormat", "NumberFormat", "Number format", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "Color", "Color", GH_ParamAccess.item);
            pManager.AddGenericParameter("Borders", "Borders", "Borders", GH_ParamAccess.item);
            pManager.AddGenericParameter("Padding", "Padding", "Padding", GH_ParamAccess.item);
            pManager.AddTextParameter("HorizontalAligment", "HorizontalAligment", "Horizontal aligment", GH_ParamAccess.item);
            pManager.AddTextParameter("VerticalAligment", "VerticalAligment", "Vertical aligment", GH_ParamAccess.item);
            pManager.AddTextParameter("Wrap", "Wrap", "Wrap", GH_ParamAccess.item);
            pManager.AddTextParameter("TextDirection", "TextDirection", "Text direction", GH_ParamAccess.item);
            pManager.AddGenericParameter("TextFormat", "TextFormat", "Text format", GH_ParamAccess.item);
            pManager.AddGenericParameter("TextRotation", "TextRotation", "Text rotation", GH_ParamAccess.item);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
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
            var numberFormat =new NumberFormat();
            var color = new Color();
            var borders = new Borders();
            var padding = new Padding();
            var horizontalAligment = string.Empty;
            var verticalAligment = string.Empty;
            var wrap = string.Empty;
            var textDirection = string.Empty;
            var textFormat = new TextFormat();
            var textRotation = new TextRotation();

            if (!DA.GetData(0, ref numberFormat))
            {
                return;
            }
            if (!DA.GetData(1, ref color))
            {
                return;
            }
            if (!DA.GetData(2, ref borders))
            {
                return;
            }
            if (!DA.GetData(3, ref padding))
            {
                return;
            }
            if (!DA.GetData(4, ref horizontalAligment))
            {
                return;
            }
            if (!DA.GetData(5, ref verticalAligment))
            {
                return;
            }
            if (!DA.GetData(6, ref wrap))
            {
                return;
            }
            if (!DA.GetData(7, ref textDirection))
            {
                return;
            }
            if (!DA.GetData(8, ref textFormat))
            {
                return;
            }
            if (!DA.GetData(9, ref textRotation))
            {
                return;
            }

            var cellFormat = new CellFormat()
            {
                NumberFormat=numberFormat,
                BackgroundColor=color,
                Borders=borders,
                Padding=padding,
                HorizontalAlignment=horizontalAligment,
                VerticalAlignment=verticalAligment,
                WrapStrategy=wrap,
                TextDirection=textDirection,
                TextFormat=textFormat,
                TextRotation=textRotation,
            };

            DA.SetData(0, cellFormat);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
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