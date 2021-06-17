using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellData Goo.
    /// </summary>
    public class GH_CellBorders : GH_Goo<Borders>
    {
        /// <summary>
        /// Default.
        /// </summary>
        public GH_CellBorders()
        {
        }

        /// <summary>
        /// CellData Goo.
        /// </summary>
        /// <param name="cellData">CellData.</param>
        public GH_CellBorders(Borders borders)
        {
            Value = borders;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellDataGoo">CellData Goo.</param>
        public GH_CellBorders(GH_CellBorders bordersGoo)
        {
            if (bordersGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var bordersJson = JsonConvert.SerializeObject(bordersGoo.Value, Formatting.Indented);
                var borders = JsonConvert.DeserializeObject<Borders>(bordersJson);
                Value = borders;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "Borders";

        /// <inheritdoc/>
        public override string TypeDescription => "The borders of the cell.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_CellBorders(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var top = string.Empty;

            if (Value.Top != null)
            {
                top = $@"Formatted value: {Value.Top}";
            }

            var bottom = string.Empty;

            if (Value.Bottom != null)
            {
                bottom = $@"Formatted value: {Value.Bottom}";
            }

            var right = string.Empty;

            if (Value.Right != null)
            {
                bottom = $@"Formatted value: {Value.Right}";
            }

            var left = string.Empty;

            if (Value.Right != null)
            {
                bottom = $@"Formatted value: {Value.Left}";
            }

            return $"{top} \n{bottom} \n {right} \n {left}";
        }
    }
}
