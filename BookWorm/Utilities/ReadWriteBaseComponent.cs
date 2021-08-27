using System;
using System.Windows.Forms;
using Grasshopper.Kernel;

namespace BookWorm.Utilities
{
    /// <summary>
    /// Base component for reading and writing spreadsheets.
    /// </summary>
    public class ReadWriteBaseComponent : GH_Component
    {
        /// <summary>
        /// Gets spreadsheet Id.
        /// </summary>
        public string SpreadsheetId { get; set; }

        /// <summary>
        /// Gets sheet name.
        /// </summary>
        public string SheetName { get; private set; }

        /// <summary>
        /// Gets a range of cells in A1 notaton.
        /// </summary>
        public string CellRange { get; private set; }

        /// <summary>
        /// Gets a range of cells with given sheet name in A1 notation on a spreadsheet.
        /// If range include only sheet name, its refer all cells in named sheet.
        /// If range include only cell range its refer cells of the first visible sheet.
        /// </summary>
        public string SpreadsheetRange => SheetsUtilities.GetSpreadsheetRange(SheetName, CellRange);

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadWriteBaseComponent"/> class.
        /// </summary>
        /// <param name="name">Name.</param>
        /// <param name="nick">Nickname.</param>
        /// <param name="desc">Description.</param>
        /// <param name="tab">Tab.</param>
        /// <param name="subTab">Group.</param>
        public ReadWriteBaseComponent(string name, string nick, string desc, string tab, string subTab)
            : base(name, nick, desc, tab, subTab)
        {
        }

        /// <inheritdoc/>
        public override Guid ComponentGuid => throw new NotImplementedException();

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);

            pManager.AddTextParameter("Sheet Name", "N", "Sheet Name", GH_ParamAccess.item, string.Empty);

            pManager.AddTextParameter(
                "Cell Range",
                "CR",
                "Range of cells in \'a1\' notation. For example A1:B5 - range of cells, A15 - single cell, A:C - range of columns, etc.",
                GH_ParamAccess.item,
                string.Empty);

            pManager[1].Optional = true;
            pManager[2].Optional = true;
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetUrl = string.Empty;
            var sheetName = string.Empty;
            string range = string.Empty;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;
            SpreadsheetId = Util.ParseUrl(spreadsheetUrl);

            if (DA.GetData(1, ref sheetName) && sheetName != string.Empty)
            {
                SheetName = sheetName;
            }
            else
            {
                var sheetId = Util.GetSheetIdFromUrl(spreadsheetUrl);

                if (sheetId != null)
                {
                    SheetName = SheetsUtilities.GetSheetName(SpreadsheetId, (int)sheetId);
                }
            }

            if (DA.GetData(2, ref range))
                CellRange = range.ToUpper();
        }

        /// <inheritdoc/>
        public override void AppendAdditionalMenuItems(ToolStripDropDown menu)
        {
            Menu_AppendSeparator(menu);
            Menu_AppendItem(menu, "Open spreadsheet in browser...", ElementClicked, true, false);
        }

        private void ElementClicked(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(SpreadsheetId))
            {
                System.Diagnostics.Process.Start(@"https://docs.google.com/spreadsheets/d/" + SpreadsheetId);
            }
        }
    }
}
