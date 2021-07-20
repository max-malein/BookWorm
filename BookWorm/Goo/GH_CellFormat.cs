using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellFormat Goo.
    /// </summary>
    public class GH_CellFormat : GH_Goo<CellFormat>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellFormat"/> class.
        /// Default.
        /// </summary>
        public GH_CellFormat()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellFormat"/> class.
        /// CellFormat Goo.
        /// </summary>
        /// <param name="cellFormat">CellFormat.</param>
        public GH_CellFormat(CellFormat cellFormat)
        {
            Value = cellFormat;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellFormat"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellFormatGoo">CellFormat Goo.</param>
        public GH_CellFormat(GH_CellFormat cellFormatGoo)
        {
            if (cellFormatGoo != null)
            {
                var cellFormatJson = JsonConvert.SerializeObject(cellFormatGoo.Value, Formatting.Indented);
                var cellFormat = JsonConvert.DeserializeObject<CellFormat>(cellFormatJson);
                Value = cellFormat;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "CellFormat";

        /// <inheritdoc/>
        public override string TypeDescription => "The format of a cell";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_CellFormat(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var formattedColor = string.Empty;

            if (Value.BackgroundColor != null)
            {
                var color = SheetsUtilities.GetSystemDrawingColor(Value.BackgroundColor);

                formattedColor = SheetsUtilities.GetFormattedARGB(color);
            }

            // Feels like displaying more info than color is fine idea but idk.
            //if (Value.NumberFormat != null)
            //{

            //}

            //if (Value.TextFormat != null)
            //{

            //}

            //if (Value.Borders != null)
            //{

            //}

            return $@"Background color: {formattedColor}";
        }
    }
}
