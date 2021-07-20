using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// Cell Borders Goo.
    /// </summary>
    public class GH_CellBorders : GH_Goo<Borders>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorders"/> class.
        /// Default.
        /// </summary>
        public GH_CellBorders()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorders"/> class.
        /// Cell Borders Goo.
        /// </summary>
        /// <param name="cellBorders">Cell Borders.</param>
        public GH_CellBorders(Borders cellBorders)
        {
            Value = cellBorders;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorders"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellBordersGoo">Cell Borders Goo.</param>
        public GH_CellBorders(GH_CellBorders cellBordersGoo)
        {
            if (cellBordersGoo != null)
            {
                var bordersJson = JsonConvert.SerializeObject(cellBordersGoo.Value, Formatting.Indented);
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
                top = $"Top\n{new GH_CellBorder(Value.Top)}\n\n";
            }

            var bottom = string.Empty;

            if (Value.Bottom != null)
            {
                bottom = $"Bottom\n{new GH_CellBorder(Value.Bottom)}\n\n";
            }

            var left = string.Empty;

            if (Value.Left != null)
            {
                left = $"Left\n{new GH_CellBorder(Value.Left)}\n\n";
            }

            var right = string.Empty;

            if (Value.Right != null)
            {
                right = $"Right\n{new GH_CellBorder(Value.Right)}\n\n";
            }

            return $"{top}{bottom}{left}{right}";
        }
    }
}
