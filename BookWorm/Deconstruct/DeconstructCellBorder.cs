namespace BookWorm.Deconstruct
{
    using System;
    using BookWorm.Goo;
    using Grasshopper.Kernel;

    public class DeconstructCellBorder : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeconstructCellBorder"/> class.
        /// Deconstruct cell border.
        /// </summary>
        public DeconstructCellBorder()
          : base(
                "DeconstructCellBorder",
                "DecCellBorder",
                "Deconstruct cell border",
                "BookWorm",
                "Deconstruct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Border", "Border", "Border", GH_ParamAccess.item);
 
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Style", "Style", "0 - STYLE_UNSPECIFIED, 1  -DOTTED, 2 - DASHED, 3 - SOLID, 4 - SOLID_MEDIUM, 5 - SOLID_THICK, 6 - NONE, 7 - DOUBLE", GH_ParamAccess.item);
            pManager.AddTextParameter("Color", "Color", "Border color", GH_ParamAccess.item);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var borderGoo = new GH_CellBorder();
            if (!DA.GetData(0, ref borderGoo) || borderGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in Border input");
                return;
            }

            var border = borderGoo.Value;
            var color = string.Empty;
            if (!(border.Color == null))
            {
              color = border.Color.Alpha.ToString() + "," + border.Color.Red.ToString() + "," + border.Color.Green.ToString() + "," + border.Color.Blue.ToString();
            }

            DA.SetData(0, border.Style);
            DA.SetData(1, color);

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
            get { return new Guid("e172f719-d9cb-418d-b2ea-0a60fb5f4fc8"); }
        }
    }
}