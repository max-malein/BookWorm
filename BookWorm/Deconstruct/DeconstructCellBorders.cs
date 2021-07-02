using BookWorm.Goo;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Deconstruct
{
    public class DeconstructCellBorders : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the DeconstructCellBorders class.
        /// </summary>
        public DeconstructCellBorders()
          : base("DeconstructCellBorders", "DecCellBorders",
              "Deconstruct cell borders",
              "BookWorm", "Deconstruct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Borders", "Borders", "Borders", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("BorderTop", "BTop", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderBottom", "BBottom", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderLeft", "BLeft", "Border", GH_ParamAccess.item);

            pManager.AddGenericParameter("BorderRight", "BRight", "Border", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var bordersGoo = new GH_CellBorders();
            if (!DA.GetData(0, ref bordersGoo) || bordersGoo.Value == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in Borders input");
                return;
            }

            var textRotation = bordersGoo.Value;

            DA.SetData(0, new GH_CellBorder(textRotation.Top));
            DA.SetData(1, new GH_CellBorder(textRotation.Bottom));
            DA.SetData(2, new GH_CellBorder(textRotation.Left));
            DA.SetData(3, new GH_CellBorder(textRotation.Right));

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
            get { return new Guid("c7584d39-1b05-4a3f-8820-7e2929a3f9fa"); }
        }
    }
}