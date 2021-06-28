using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellData Goo.
    /// </summary>
    public class GH_NumberFormat : GH_Goo<NumberFormat>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_NumberFormat"/> class.
        /// Default.
        /// </summary>
        public GH_NumberFormat()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_NumberFormat"/> class.
        /// CellData Goo.
        /// </summary>
        /// <param name="cellData">CellData.</param>
        public GH_NumberFormat(NumberFormat numberFormat)
        {
            Value = numberFormat;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_NumberFormat"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellDataGoo">CellData Goo.</param>
        public GH_NumberFormat(GH_NumberFormat NumberFormatGoo)
        {
            if (NumberFormatGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var numberFormatJson = JsonConvert.SerializeObject(NumberFormatGoo.Value, Formatting.Indented);
                var numberFormatData = JsonConvert.DeserializeObject<NumberFormat>(numberFormatJson);
                Value = numberFormatData;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "NumberFormat";

        /// <inheritdoc/>
        public override string TypeDescription => "The number format of a cell.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_NumberFormat(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var typeStringValue = string.Empty;

            // Only requested from spreadsheet cells got formatted value.
            if (Value.Type != null)
            {
                typeStringValue = $@" Number format type: {Value.Type}";
            }

            return $"{typeStringValue}";
        }
    }
}
