using BookWorm.Goo;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Deconstruct
{
    public class DeconstructPadding : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the DeconstructPadding class.
        /// </summary>
        public DeconstructPadding()
          : base("DeconstructPadding", "DecPadding",
              "Deconstruct padding",
              "BookWorm", "Deconstruct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Padding", "Padding", "Padding", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Top", "Top", "Top (int)", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Bottom", "Bottom", "Bottom (int)", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Left", "Left", "Left (int)", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Right", "Right", "Right (int)", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var paddingGoo = new GH_Padding();
            if (!DA.GetData(0, ref paddingGoo) || paddingGoo.Value == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in Padding input");
                return;
            }

            var padding = paddingGoo.Value;

            DA.SetData(0, padding.Top);
            DA.SetData(1, padding.Bottom);
            DA.SetData(2, padding.Left);
            DA.SetData(3, padding.Right);
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
            get { return new Guid("688ab7b3-10c9-4134-be83-6f5c45e01cff"); }
        }
    }
}