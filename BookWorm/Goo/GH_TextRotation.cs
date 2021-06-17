using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellData Goo.
    /// </summary>
    public class GH_TextRotation : GH_Goo<TextRotation>
    {
        /// <summary>
        /// Default.
        /// </summary>
        public GH_TextRotation()
        {
        }

        /// <summary>
        /// CellData Goo.
        /// </summary>
        /// <param name="textRotation">CellData.</param>
        public GH_TextRotation(TextRotation textRotation)
        {
            Value = textRotation;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="textRotationGoo">CellData Goo.</param>
        public GH_TextRotation(GH_TextRotation textRotationGoo)
        {
            if (textRotationGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
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

            var angleString = string.Empty;

            if (Value.Angle != null)
            {
                angleString = $@"Formatted value: {Value.Angle}";
            }

            var verticalString = string.Empty;

            if (Value.Angle != null)
            {
                verticalString = $@"Formatted value: {Value.Vertical}";
            }





            return $"{angleString} \n{verticalString}";
        }
    }
}
