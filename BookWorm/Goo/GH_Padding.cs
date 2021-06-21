using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellData Goo.
    /// </summary>
    public class GH_Padding : GH_Goo<Padding>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_Padding"/> class.
        /// Default.
        /// </summary>
        public GH_Padding()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_Padding"/> class.
        /// CellData Goo.
        /// </summary>
        /// <param name="padding">CellData.</param>
        public GH_Padding(Padding padding)
        {
            Value = padding;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_Padding"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="paddingGoo">CellData Goo.</param>
        public GH_Padding(GH_Padding paddingGoo)
        {
            if (paddingGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var paddingJson = JsonConvert.SerializeObject(paddingGoo.Value, Formatting.Indented);
                var padding = JsonConvert.DeserializeObject<Padding>(paddingJson);
                Value = padding;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "Padding";

        /// <inheritdoc/>
        public override string TypeDescription => "The amount of padding around the cell, in pixels. When updating padding, every field must be specified.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_Padding(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var bottom = string.Empty;

            if (Value.Bottom != null)
            {
                bottom = $@"Formatted value: {Value.Bottom.Value}";
            }

            var top = string.Empty;

            if (Value.Bottom != null)
            {
                top = $@"Formatted value: {Value.Top.Value}";
            }

            var right = string.Empty;

            if (Value.Right != null)
            {
                right = $@"Formatted value: {Value.Right.Value}";
            }

            var left = string.Empty;

            if (Value.Left != null)
            {
                left = $@"Formatted value: {Value.Left.Value}";
            }

            return $"{bottom} \n{top} \n{right} \n{left}";
        }
    }
}
