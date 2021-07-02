using Google.Apis.Sheets.v4.Data;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Helper operations with color for Google Sheets.
    /// </summary>
    public class GoogleSheetsColor
    {
        /// <summary>
        /// Convert ARGB-System.Drawing.Color-type color into Google.Apis.Sheets.v4.Data.Color-type color.
        /// </summary>
        /// <param name="colorARGB">System.Drawing.Color-type color.</param>
        /// <returns>Google.Apis.Sheets.v4.Data.Color-type color.</returns>
        public Color GetGoogleSheetsColor(System.Drawing.Color colorARGB)
        {
            var googleSheetColor = new Color
            {
                Alpha = colorARGB.A / 255f,
                Red = colorARGB.R / 255f,
                Green = colorARGB.G / 255f,
                Blue = colorARGB.B / 255f,
            };

            return googleSheetColor;
        }
    }
}
