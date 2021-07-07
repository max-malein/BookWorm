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
            pManager.AddColourParameter("Foreground Colour", "Colour", "Foreground colour of text", GH_ParamAccess.item);

            pManager.AddTextParameter("Font Family", "FontFamily", "Font family", GH_ParamAccess.item);

            pManager.AddIntegerParameter("Font Size", "FontSize", "Font size", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Bold", "Bold", "Is bold", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Italic", "Italic", "Is italic", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Strikethrought", "Strikethrought", "Is strikethrought", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Underline", "Underline", "Is underline", GH_ParamAccess.item);

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
            var colorARGB = System.Drawing.Color.Empty;
            var fontFamily = string.Empty;
            int fontSize = 0;
            bool isBold = false;
            bool isItalic = false;
            bool isStrikethrought = false;
            bool isUnderline = false;

            var textFormat = new TextFormat();

            if (DA.GetData(0, ref colorARGB))
            {
                var googleSheetsColor = new Utilities.SheetsUtilities();

                textFormat.ForegroundColor = googleSheetsColor.GetGoogleSheetsColor(colorARGB);
            }

            if (DA.GetData(1, ref fontFamily) && !string.IsNullOrEmpty(fontFamily))
            {
                textFormat.FontFamily = fontFamily;
            }

            if (DA.GetData(2, ref fontSize))
            {
                textFormat.FontSize = fontSize;
            }

            if (DA.GetData(3, ref isBold))
            {
                textFormat.Bold = isBold;
            }

            if (DA.GetData(4, ref isItalic))
            {
                textFormat.Italic = isItalic;
            }

            if (DA.GetData(5, ref isStrikethrought))
            {
                textFormat.Strikethrough = isStrikethrought;
            }

            if (DA.GetData(6, ref isUnderline))
            {
                textFormat.Underline = isUnderline;
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