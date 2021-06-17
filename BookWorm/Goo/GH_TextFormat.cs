using System;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// CellFormat Goo.
    /// </summary>
    public class GH_TextFormat : GH_Goo<TextFormat>
    {
        /// <summary>
        /// Default.
        /// </summary>
        public GH_TextFormat()
        {
        }

        /// <summary>
        /// CellFormat Goo.
        /// </summary>
        /// <param name="cellFormat">CellFormat.</param>
        public GH_TextFormat(TextFormat textFormat)
        {
            Value = textFormat;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellFormatGoo">CellFormat Goo.</param>
        public GH_TextFormat(GH_TextFormat textFormatGoo)
        {
            if (textFormatGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
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

            var textColorString = string.Empty;

            if (Value.ForegroundColor != null)
            {
                var color = Value.ForegroundColor;

                var alpha = color.Alpha;
                var red = color.Red;
                var green = color.Green;
                var blue = color.Blue;

                textColorString = $@"Background color: A:{alpha}, R:{red}, G:{green}, B:{blue}";
            }

            var boldString = string.Empty;

            if (Value.Bold != null)
            {
                 boldString = $@"Bold: {Value.Bold}"; 
            }

            var fontString = string.Empty;

            if (Value.FontFamily != null)
            {
                fontString = $@"FontFamily: {Value.FontFamily}";
            }

            var fontSizeString = string.Empty;

            if (Value.FontSize != null)
            {
                fontSizeString = $@"FontSize: {Value.FontSize}";
            }

            var italicString = string.Empty;

            if (Value.Italic != null)
            {
                italicString = $@"FontSize: {Value.Italic}";
            }

            var strikethroughString = string.Empty;

            if (Value.Strikethrough != null)
            {
                strikethroughString = $@"FontSize: {Value.Italic}";
            }

            return $"{textColorString} \n{boldString} \n{fontString} \n{fontSizeString} \n{italicString} \n{strikethroughString}";
        }
    }
}
