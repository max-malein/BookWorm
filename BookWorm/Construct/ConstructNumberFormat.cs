using System;
using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;

namespace BookWorm.Construct
{
    public class ConstructNumberFormat : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructNumberFormat"/> class.
        /// </summary>
        public ConstructNumberFormat()
          : base(
                "ConstructNumberFormat",
                "ConNumberFormat",
                "Construct number format",
                "BookWorm",
                "Construct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter(
                "NumberFormatType",
                "NumberFormatType",
                "0-TEXT\n "
                + "1-NUMBER\n "
                + "2-PERCENT\n "
                + "3-CURRENCY\n "
                + "4-DATE\n "
                + "5-TIME\n "
                + "6-DATE_TIME\n "
                + "7-SCIENTIFIC",
                GH_ParamAccess.item);

            pManager.AddTextParameter("Pattern", "Pattern", "Pattern", GH_ParamAccess.item);

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
            pManager.AddGenericParameter("NumberFormat", "NumberFormat", "NumberFormat", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var numberFormatType = 0;
            var pattern = string.Empty;
            var numberFormat = new NumberFormat();

            if (DA.GetData(0, ref numberFormatType))
            {
                numberFormat.Type = numberFormatType.ToString();
            }

            if (DA.GetData(1, ref pattern))
            {
                numberFormat.Type = pattern;
            }

            var numberFormatGoo = new GH_NumberFormat(numberFormat);
            DA.SetData(0, numberFormatGoo);
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
            get { return new Guid("ce00f900-34f0-4a49-86c0-1c8c45db9d65"); }
        }

        /// <summary>
        /// The number format of the cell. In this documentation the locale is assumed to be en_US,
        /// but the actual format depends on the locale of the spreadsheet.
        ///
        /// TEXT - Text formatting, e.g 1000.12
        /// NUMBER - Number formatting, e.g, 1,000.12
        /// PERCENT - Percent formatting, e.g 10.12%
        /// CURRENCY - Currency formatting, e.g $1,000.12
        /// DATE - Date formatting, e.g 9/26/2008
        /// TIME - Time formatting, e.g 3:59:00 PM
        /// DATE_TIME - Date+Time formatting, e.g 9/26/08 15:59:00
        /// SCIENTIFIC - Scientific number formatting, e.g 1.01E+03
        /// </summary>
        public enum NumberFormatType
        {
            TEXT,
            NUMBER,
            PERCENT,
            CURRENCY,
            DATE,
            TIME,
            DATE_TIME,
            SCIENTIFIC,
        }
    }
}