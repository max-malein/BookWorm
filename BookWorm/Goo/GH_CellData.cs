using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellData Goo.
    /// </summary>
    public class GH_CellData : GH_Goo<CellData>
    {
        /// <summary>
        /// Default.
        /// </summary>
        public GH_CellData()
        {
        }

        /// <summary>
        /// CellData Goo.
        /// </summary>
        /// <param name="cellData">CellData.</param>
        public GH_CellData(CellData cellData)
        {
            Value = cellData;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellDataGoo">CellData Goo.</param>
        public GH_CellData(GH_CellData cellDataGoo)
        {
            if (cellDataGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var cellDataJson = JsonConvert.SerializeObject(cellDataGoo.Value, Formatting.Indented);
                var cellData = JsonConvert.DeserializeObject<CellData>(cellDataJson);
                Value = cellData;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "CellData";

        /// <inheritdoc/>
        public override string TypeDescription => "Data about a specific Google Sheets Cell";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_CellData(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var cellStringValue = string.Empty;

            // Only requested from spreadsheet cells got formatted value.
            if (Value.FormattedValue != null)
            {
                cellStringValue = $@"Formatted value: {Value.FormattedValue}";
            }

            // Manually created cell have only user entered value
            else if (Value.UserEnteredValue != null)
            {
                var userValueGoo = new GH_ExtendedValue(Value.UserEnteredValue);

                cellStringValue = $@"{userValueGoo}";
            }

            var cellFormatString = string.Empty;

            if (Value.UserEnteredFormat != null)
            {
                var cellFormatGoo = new GH_CellFormat(Value.UserEnteredFormat);
                cellFormatString = $@"{cellFormatGoo}";
            }

            return $"{cellStringValue} \n{cellFormatString}";
        }
    }
}
