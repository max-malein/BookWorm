namespace BookWorm.Deconstruct
{
    using System;
    using BookWorm.Goo;
    using Grasshopper.Kernel;

    public class DeconstructTextRotation : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeconstructTextRotation"/> class.
        /// Deconstruct text rotation.
        /// </summary>
        public DeconstructTextRotation()
          : base(
                "DeconstructTextRotation",
                "DecTextRotation",
                "Deconstruct text rotation",
                "BookWorm", 
                "Deconstruct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("TextRotation", "TextRotation", "Text rotation", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("angle", "angle", "angle", GH_ParamAccess.item);
            pManager.AddBooleanParameter("vertical", "vertical", "vertical", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var textRotationGoo = new GH_TextRotation();
            if (!DA.GetData(0, ref textRotationGoo) || textRotationGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in TextRotation input");
                return;
            }

            var textRotation = textRotationGoo.Value;

            DA.SetData(0, textRotation.Angle);
            DA.SetData(1, textRotation.Vertical);
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
            get { return new Guid("490bee06-24e4-482e-82d6-1cee90aecb8a"); }
        }
    }
}