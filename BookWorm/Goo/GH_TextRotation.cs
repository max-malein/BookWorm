using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// TextRotation Goo.
    /// </summary>
    public class GH_TextRotation : GH_Goo<TextRotation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextRotation"/> class.
        /// Default.
        /// </summary>
        public GH_TextRotation()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextRotation"/> class.
        /// TextRotaion Goo.
        /// </summary>
        /// <param name="textRotation">TextRotation.</param>
        public GH_TextRotation(TextRotation textRotation)
        {
            Value = textRotation;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextRotation"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="textRotationGoo">TextRotation Goo.</param>
        public GH_TextRotation(GH_TextRotation textRotationGoo)
        {
            if (textRotationGoo != null)
            {
                var textRotationJson = JsonConvert.SerializeObject(textRotationGoo.Value, Formatting.Indented);
                var textRotation = JsonConvert.DeserializeObject<TextRotation>(textRotationJson);
                Value = textRotation;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "TextRotation";

        /// <inheritdoc/>
        public override string TypeDescription => "The rotation applied to text in a cell.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_TextRotation(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var outputString = string.Empty;

            if (Value.Angle != null)
            {
                var angleString = $"Rotation angle: {Value.Angle}";
                outputString = angleString;
            }
            else if (Value.Vertical != null)
            {
                var verticalString = $"Vertical string: {Value.Vertical}";
                outputString = verticalString;
            }

            return outputString;
        }
    }
}
