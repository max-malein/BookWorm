// <copyright file="ConstructTextFormat.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Construct
{
    using System;
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
            pManager.AddColourParameter("foregroundColor", "foregroundColor", "foreground color", GH_ParamAccess.item);
            pManager.AddTextParameter("fontFamily", "fontFamily", "Font family", GH_ParamAccess.item);
            pManager.AddIntegerParameter("fontSize", "fontSize", "Font size", GH_ParamAccess.item);
            pManager.AddBooleanParameter("bold", "bold", "bold", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("italic", "italic", "italic", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("strikethrought", "strikethrought", "strikethrought", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("underline", "underline", "underline", GH_ParamAccess.item, false);
            pManager.AddTextParameter("link", "link", "link", GH_ParamAccess.item, "http://parametrica.team");
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


            var foregroundColor = new Color();
            var fontFamily = "Arial";
            var fontSize = 14;
            var bold = false;
            var italic = false;
            var strikethrought = false;
            var underLine = false;
            var link= "http://parametrica.team";

            DA.GetData(0, ref foregroundColor);
            DA.GetData(1, ref fontFamily);
            DA.GetData(2, ref fontSize);
            DA.GetData(3, ref bold);
            DA.GetData(4, ref italic);
            DA.GetData(5, ref underLine);
            DA.GetData(6, ref link);

            var textFormat = new TextFormat()
            {
                ForegroundColor = foregroundColor,
                FontFamily = fontFamily,
                FontSize = fontSize,
                Bold = bold,
                Italic = italic,
                Strikethrough = strikethrought,
                Underline = underLine,
                ETag = link,

            };

            DA.SetData(0, textFormat);

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