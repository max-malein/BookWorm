using System;
using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Construct
{
    public class ConstructTextRotation : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructTextRotation"/> class.
        /// </summary>
        public ConstructTextRotation()
          : base(
                "ConstructTextRotation",
                "ConTextRotation",
                "Construct text rotation.\nNote: only one parameter can be set",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Angle", "Angle", "Angle in degrees between -90 and 90", GH_ParamAccess.item);
            pManager.AddBooleanParameter(
                "Vertical",
                "Vertical",
                "If true, text reads top to bottom, but the orientation of individual characters is unchanged",
                GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("TextRotation", "TextRotation", "The rotation applied to text in a cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var angle = 90;
            var vertical = false;

            var textRotation = new TextRotation();

            // Only one field can be set.
            if (DA.GetData(0, ref angle) && !DA.GetData(1, ref vertical))
            {
                if (angle < -90 || angle > 90)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Angle value must be between -90 and 90");
                    return;
                }

                textRotation.Angle = angle;
            }
            else if (DA.GetData(1, ref vertical) && !DA.GetData(0, ref angle))
            {
                textRotation.Vertical = vertical;
            }
            else if (DA.GetData(0, ref angle) && DA.GetData(1, ref vertical))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Only one parameter can be set");
                return;
            }

            var textRotationGoo = new GH_TextRotation(textRotation);
            DA.SetData(0, textRotationGoo);
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
            get { return new Guid("d361a5a9-5614-4da6-9bc2-aa32ec2dc24f"); }
        }
    }
}