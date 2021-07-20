namespace BookWorm.Goo
{
    using BookWorm.Utilities;
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel.Types;
    using Newtonsoft.Json;

    /// <summary>
    /// Cell Border Goo.
    /// </summary>
    public class GH_CellBorder : GH_Goo<Border>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorder"/> class.
        /// Default.
        /// </summary>
        public GH_CellBorder()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorder"/> class.
        /// CellBorder Goo.
        /// </summary>
        /// <param name="cellBorder">Cell Border.</param>
        public GH_CellBorder(Border cellBorder)
        {
            this.Value = cellBorder;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorder"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellBorderGoo">Cell Border Goo.</param>
        public GH_CellBorder(GH_CellBorder cellBorderGoo)
        {
            if (cellBorderGoo != null)
            {
                var cellBorderJson = JsonConvert.SerializeObject(cellBorderGoo.Value, Formatting.Indented);
                var cellBorder = JsonConvert.DeserializeObject<Border>(cellBorderJson);
                this.Value = cellBorder;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => this.Value != null;

        /// <inheritdoc/>
        public override string TypeName => "Border";

        /// <inheritdoc/>
        public override string TypeDescription => "A border along a cell.";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_CellBorder(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (this.Value == null)
            {
                return string.Empty;
            }

            var style = Value.Style;
            var color = SheetsUtilities.GetSystemDrawingColor(Value.Color);

            var formattedColor = SheetsUtilities.GetFormattedARGB(color);

            return $"Style: {style}\nColor: {formattedColor}";
        }
    }
}