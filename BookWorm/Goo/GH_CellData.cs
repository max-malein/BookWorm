using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// Goo для 
    /// </summary>
    public class GH_CellData : GH_Goo<CellData>
    {
        // constructors
        public GH_CellData()
        {
        }

        public GH_CellData(CellData cellData)
        {
            Value = cellData;
        }

        /// <summary>
        /// Копия goo ячейки.
        /// </summary>
        /// <param name="cellDataGoo">Goo ячейки.</param>
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
        public override string TypeDescription => "Google Sheets Cell Data";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_CellData(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value != null)
            {
                var str = Value.FormattedValue;
                return $@"Cell: {str}";
            }

            return string.Empty;
        }
    }
}
