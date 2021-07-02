using BookWorm.Goo;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Deconstruct
{
    public class DeconstructTextFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the DeconstructTextFormat class.
        /// </summary>
        public DeconstructTextFormat()
          : base("DeconstructTextFormat", "DecTextFormat",
              "Deconstruct text format",
              "BookWorm", "Deconstruct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("TextFormat", "TextFormat", "TextFormat", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("foregroundColor", "foregroundColor", "foreground color", GH_ParamAccess.item);
            pManager.AddTextParameter("fontFamily", "fontFamily", "Font family", GH_ParamAccess.item);
            pManager.AddIntegerParameter("fontSize", "fontSize", "Font size", GH_ParamAccess.item);
            pManager.AddBooleanParameter("bold", "bold", "bold", GH_ParamAccess.item);
            pManager.AddBooleanParameter("italic", "italic", "italic", GH_ParamAccess.item);
            pManager.AddBooleanParameter("strikethrought", "strikethrought", "strikethrought", GH_ParamAccess.item);
            pManager.AddBooleanParameter("underline", "underline", "underline", GH_ParamAccess.item);
            pManager.AddTextParameter("link", "link", "link", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var textFormatGoo = new GH_TextFormat();
            if (!DA.GetData(0, ref textFormatGoo) || textFormatGoo.Value == null)
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Null in TextFormat input");
                return;
            }

            var textFormat = textFormatGoo.Value;

            var color = string.Empty;
            if (!(textFormat.ForegroundColor == null))
            {
                color = textFormat.ForegroundColor.Alpha.ToString() + "," + textFormat.ForegroundColor.Red.ToString() + "," + textFormat.ForegroundColor.Green.ToString() + "," + textFormat.ForegroundColor.Blue.ToString();
                DA.SetData(0, color);
            }
            else
            {
                DA.SetData(0, color);
            }

            DA.SetData(1, textFormat.FontFamily);
            DA.SetData(2, textFormat.FontSize);
            DA.SetData(3, textFormat.Bold);
            DA.SetData(4, textFormat.Italic);
            DA.SetData(5, textFormat.Strikethrough);
            DA.SetData(6, textFormat.Underline);
            DA.SetData(7, textFormat.ETag);
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
            get { return new Guid("2517d59e-30df-4674-ab46-b66d348c5578"); }
        }
    }
}