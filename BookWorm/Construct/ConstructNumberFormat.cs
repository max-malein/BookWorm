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

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter(
                "Number Format Type",
                "NumberFormatType",
                "The actual format depends on the locale of the spreadsheet\n\n"
                + "0 - TEXT\n"
                + "1 - NUMBER\n"
                + "2 - PERCENT\n"
                + "3 - CURRENCY\n"
                + "4 - DATE\n"
                + "5 - TIME\n"
                + "6 - DATE_TIME\n"
                + "7 - SCIENTIFIC",
                GH_ParamAccess.item);

            pManager.AddTextParameter("Pattern", "Pattern", "Pattern string used for formatting", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Number Format", "NumberFormat", "The number format of a cell", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var formatType = 0;
            var pattern = string.Empty;
            var numberFormat = new NumberFormat();

            if (DA.GetData(0, ref formatType))
            {
                if (formatType < 0 || formatType > 7)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"The number format type {formatType} does not exist");
                    return;
                }

                numberFormat.Type = Enum.GetName(typeof(NumberFormatType), formatType);
            }

            if (DA.GetData(1, ref pattern) && !string.IsNullOrEmpty(pattern))
            {
                numberFormat.Type = pattern;
            }

            var numberFormatGoo = new GH_NumberFormat(numberFormat);
            DA.SetData(0, numberFormatGoo);
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
            get { return new Guid("ce00f900-34f0-4a49-86c0-1c8c45db9d65"); }
        }

        /// <summary>
        /// The number format of the cell. In this documentation the locale is assumed to be en_US,
        /// but the actual format depends on the locale of the spreadsheet.
        ///
        /// TEXT - Text formatting, e.g 1000.12.
        /// NUMBER - Number formatting, e.g, 1,000.12.
        /// PERCENT - Percent formatting, e.g 10.12%.
        /// CURRENCY - Currency formatting, e.g $1,000.12.
        /// DATE - Date formatting, e.g 9/26/2008.
        /// TIME - Time formatting, e.g 3:59:00 PM.
        /// DATE_TIME - Date+Time formatting, e.g 9/26/08 15:59:00.
        /// SCIENTIFIC - Scientific number formatting, e.g 1.01E+03.
        /// </summary>
        private enum NumberFormatType
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