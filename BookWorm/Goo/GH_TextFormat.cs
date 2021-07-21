using BookWorm.Utilities;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// TextFormat Goo.
    /// </summary>
    public class GH_TextFormat : GH_Goo<TextFormat>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextFormat"/> class.
        /// Default.
        /// </summary>
        public GH_TextFormat()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextFormat"/> class.
        /// TextFormat Goo.
        /// </summary>
        /// <param name="textFormat">TextFormat.</param>
        public GH_TextFormat(TextFormat textFormat)
        {
            Value = textFormat;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_TextFormat"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="textFormatGoo">TextFormat Goo.</param>
        public GH_TextFormat(GH_TextFormat textFormatGoo)
        {
            if (textFormatGoo != null)
            {
                var textFormatJson = JsonConvert.SerializeObject(textFormatGoo.Value, Formatting.Indented);
                var textFormat = JsonConvert.DeserializeObject<TextFormat>(textFormatJson);
                Value = textFormat;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "TextFormat";

        /// <inheritdoc/>
        public override string TypeDescription => "The format of a run of text in a cell. Absent values indicate that the field isn't specified.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_TextFormat(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            // Text color
            var textColor = string.Empty;

            if (Value.ForegroundColor != null)
            {
                var color = SheetsUtilities.GetSystemDrawingColor(Value.ForegroundColor);

                var formattedColor = SheetsUtilities.GetFormattedARGB(color);

                textColor = $"Background color: {formattedColor}\n";
            }

            // Font family and size
            var fontString = string.Empty;

            if (Value.FontFamily != null)
            {
                fontString = $"Font: {Value.FontFamily}\n";
            }

            // Font size
            var fontSize = string.Empty;

            if (Value.FontSize != null)
            {
                fontSize = $"Font size: {Value.FontSize}\n";
            }

            // Bold
            var boldString = string.Empty;

            if (Value.Bold != null)
            {
                boldString = $"Bold: {Value.Bold}\n";
            }

            // Italic
            var italicString = string.Empty;

            if (Value.Italic != null)
            {
                italicString = $"Italic: {Value.Italic}\n";
            }

            // Strikethrought
            var strikethroughString = string.Empty;

            if (Value.Strikethrough != null)
            {
                strikethroughString = $"Strikethrought: {Value.Strikethrough}\n";
            }

            // Underline
            var underline = string.Empty;

            if (Value.Underline != null)
            {
                underline = $"Underline: {Value.Strikethrough}\n";
            }

            return $"{textColor}{fontString}{fontSize}{boldString}{italicString}{strikethroughString}{underline}";
        }
    }
}
