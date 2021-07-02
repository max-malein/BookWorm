namespace BookWorm.Deconstruct
{
    using System;
    using BookWorm.Goo;
    using Grasshopper.Kernel;

    public class DeconstructExtendedValue : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeconstructExtendedValue"/> class.
        /// Deconstruct extended value.
        /// </summary>
        public DeconstructExtendedValue()
          : base(
                "DeconstructExtendedValue", 
                "DecExtendedValue",
                "Deconstruct extended value",
                "BookWorm", 
                "Deconstruct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("ExtendedValue", "ExtendedValue", "Extended value", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Value", "Value", "Value", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var extGoo = new GH_ExtendedValue();
            if (!DA.GetData(0, ref extGoo) || extGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in ExtendedValue input");
                return;
            }

            var extVal = extGoo.Value;

            if (!(extVal.BoolValue == null))
            {
                DA.SetData(0, extVal.BoolValue);
            }

            if (!(extVal.FormulaValue == null))
            {
                DA.SetData(0, extVal.FormulaValue);
            }

            if (!(extVal.NumberValue == null))
            {
                DA.SetData(0, extVal.NumberValue);
            }

            if (!(extVal.StringValue == null))
            {
                DA.SetData(0, extVal.StringValue);
            }


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
            get { return new Guid("62101ce2-f0d8-4bd5-9a6a-9060bf972c2c"); }
        }
    }
}