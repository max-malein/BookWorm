// <copyright file="ConstructTextFormat.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Construct
{
    using System;
    using BookWorm.Goo;
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel;

    /// <summary>
    /// Construct text format.
    /// </summary>
    public class ConstructTextFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructTextFormat"/> class.
        /// </summary>
        public ConstructTextFormat()
          : base(
                "ConstructTextFormat",
                "ConTextFormat",
                "Construct text format",
                "BookWorm",
                "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("foregroundColor", "foregroundColor", "foreground color", GH_ParamAccess.item);
            pManager.AddTextParameter("fontFamily", "fontFamily", "Font family", GH_ParamAccess.item);
            pManager.AddIntegerParameter("fontSize", "fontSize", "Font size", GH_ParamAccess.item);
            pManager.AddBooleanParameter("bold", "bold", "bold", GH_ParamAccess.item);
            pManager.AddBooleanParameter("italic", "italic", "italic", GH_ParamAccess.item);
            pManager.AddBooleanParameter("strikethrought", "strikethrought", "strikethrought", GH_ParamAccess.item);
            pManager.AddBooleanParameter("underline", "underline", "underline", GH_ParamAccess.item);
            pManager.AddTextParameter("link", "link", "link", GH_ParamAccess.item, "http://parametrica.team");

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("TextFormat", "TextFormat", "TextFormat", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            var color = string.Empty;
            var fontFamily = "Arial";
            int fontSize = 14;
            bool bold = false;
            bool italic = false;
            bool strikethrought = false;
            bool underLine = false;
            var link= "http://parametrica.team";

            var textFormat = new TextFormat();

            if (DA.GetData(0, ref color))
            {
                textFormat.ForegroundColor = new Color
                {
                    Alpha = 255F,
                    Red = float.Parse(color.Split(',')[0]),
                    Green = float.Parse(color.Split(',')[1]),
                    Blue = float.Parse(color.Split(',')[2]),
                };
            }

            if (DA.GetData(1, ref fontFamily))
            {
                textFormat.FontFamily = fontFamily;
            }

            if (DA.GetData(2, ref fontSize))
            {
                textFormat.FontSize = fontSize;
            }

            if (DA.GetData(3, ref bold))
            {
                textFormat.Bold = bold;
            }

            if (DA.GetData(4, ref italic))
            {
                textFormat.Italic = italic;
            }

            if (DA.GetData(5, ref strikethrought))
            {
                textFormat.Italic = strikethrought;
            }

            if (DA.GetData(6, ref underLine))
            {
                textFormat.Underline = underLine;
            }

            if (DA.GetData(7, ref link))
            {
                textFormat.ETag = link;
            }

            

            var textFormatGoo = new GH_TextFormat(textFormat);
            DA.SetData(0, textFormatGoo);
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
            get { return new Guid("8322e28a-63ba-493a-9b56-f97f26711d59"); }
        }
    }
}