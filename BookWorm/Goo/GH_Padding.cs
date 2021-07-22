using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// Cell Padding Goo.
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
        /// Cell Padding Goo.
        /// </summary>
        /// <param name="padding">Cell Padding.</param>
        public GH_Padding(Padding padding)
        {
            Value = padding;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_Padding"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="paddingGoo">Cell Padding Goo.</param>
        public GH_Padding(GH_Padding paddingGoo)
        {
            if (paddingGoo != null)
            {
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

            var top = string.Empty;

            if (Value.Top != null)
            {
                top = $"Top padding: {Value.Top.Value}\n";
            }

            var bottom = string.Empty;

            if (Value.Bottom != null)
            {
                bottom = $"Bottom padding: {Value.Bottom.Value}\n";
            }

            var left = string.Empty;

            if (Value.Left != null)
            {
                left = $"Left padding: {Value.Left.Value}\n";
            }

            var right = string.Empty;

            if (Value.Right != null)
            {
                right = $"Right padding: {Value.Right.Value}\n";
            }

            return $"{top}{bottom}{left}{right}";
        }
    }
}
