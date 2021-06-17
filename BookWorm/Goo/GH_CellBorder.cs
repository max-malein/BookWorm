// <copyright file="GH_CellBorder.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Goo
{
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel.Types;
    using Newtonsoft.Json;

    /// <summary>
    /// CellData Goo.
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
        /// CellData Goo.
        /// </summary>
        /// <param name="cellBorder">CellData.</param>
        public GH_CellBorder(Border cellBorder)
        {
            this.Value = cellBorder;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GH_CellBorder"/> class.
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="cellBorderGoo">CellData Goo.</param>
        public GH_CellBorder(GH_CellBorder cellBorderGoo)
        {
            if (cellBorderGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var cellDataJson = JsonConvert.SerializeObject(cellBorderGoo.Value, Formatting.Indented);
                var cellBorder = JsonConvert.DeserializeObject<Border>(cellDataJson);
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

            var style = string.Empty;

            // Only requested from spreadsheet cells got formatted value.
            if (this.Value.Style != null)
            {
                style = $@"Style: {this.Value.Style}";
            }

            return $"{style}    ";
        }
    }
}