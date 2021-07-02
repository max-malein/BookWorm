namespace BookWorm.Deconstruct
{
    using System;
    using BookWorm.Goo;
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
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var cellFormatGoo = new GH_CellFormat();
            if (!DA.GetData(0, ref cellFormatGoo) || cellFormatGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in CellFormat input");
                return;
            }

            var cellFormat = cellFormatGoo.Value;
            DA.SetData(0, new GH_NumberFormat(cellFormat.NumberFormat));

            var color = string.Empty;
            if (!(cellFormat.BackgroundColor == null))
            {
                color = cellFormat.BackgroundColor.Alpha.ToString() + "," + cellFormat.BackgroundColor.Red.ToString() + "," + cellFormat.BackgroundColor.Green.ToString() + "," + cellFormat.BackgroundColor.Blue.ToString();
                DA.SetData(1, cellFormat.BackgroundColor);
            }

            DA.SetData(1, color);

            DA.SetData(2, new GH_CellBorders( cellFormat.Borders));

            DA.SetData(3, new GH_Padding(cellFormat.Padding));

            DA.SetData(4, cellFormat.HorizontalAlignment);

            DA.SetData(5, cellFormat.VerticalAlignment);

            DA.SetData(6, cellFormat.WrapStrategy);

            DA.SetData(7, cellFormat.TextDirection);

            DA.SetData(8, new GH_TextFormat(cellFormat.TextFormat));

            DA.SetData(9, new GH_TextRotation(cellFormat.TextRotation));
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
            get { return new Guid("e3cef773-f955-4a81-95ee-ba35a9e0c495"); }
        }
    }
}