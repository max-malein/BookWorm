using System;
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
        /// Default.
        /// </summary>
        public GH_CellFormat()
        {
        }

        /// <summary>
        /// CellFormat Goo.
        /// </summary>
        /// <param name="cellFormat">CellFormat.</param>
        public GH_CellFormat(CellFormat cellFormat)
        {
            Value = cellFormat;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellFormatGoo">CellFormat Goo.</param>
        public GH_CellFormat(GH_CellFormat cellFormatGoo)
        {
            if (cellFormatGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
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

            var colorString = string.Empty;

            if (Value.BackgroundColor != null)
            {
                var color = Value.BackgroundColor;

                var alpha = color.Alpha;
                var red = color.Red;
                var green = color.Green;
                var blue = color.Blue;

                colorString = $@"Background color: A:{alpha}, R:{red}, G:{green}, B:{blue}";
            }

            // Feels like displaying more info than color is fine idea but idk.
            //if (Value.NumberFormat != null)
            //{

            //}

            //if (Value.TextFormat != null)
            //{

            //}

            return colorString;
        }
    }
}
