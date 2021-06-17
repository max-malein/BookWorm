using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Construct
{
    public class ConstructTextRotation : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ConstructTextRotation class.
        /// </summary>
        public ConstructTextRotation()
          : base("ConstructTextRotation",
                "ConTextRotation",
                "Construct text rotation",
                "BookWorm",
                "Construct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("angle", "angle", "angle", GH_ParamAccess.item);
            pManager.AddBooleanParameter("vertical", "vertical", "vertical", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("TextRotation", "TextRotation", "Text rotation", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var angle = 90;
            var vertical = false;

            var textRotation = new TextRotation();

            if (DA.GetData(0, ref angle))
            {
                textRotation.Angle = angle;
            }

            if (DA.GetData(1, ref vertical))
            {
                textRotation.Vertical = vertical;
            }

            var textRotationGoo = new GH_TextRotation(textRotation);
            DA.SetData(0, textRotationGoo);
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
            get { return new Guid("d361a5a9-5614-4da6-9bc2-aa32ec2dc24f"); }
        }
    }
}