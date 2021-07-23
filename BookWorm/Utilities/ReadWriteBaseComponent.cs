using System;
using System.Windows.Forms;
using Grasshopper.Kernel;

namespace BookWorm.Utilities
{
    public class ReadWriteBaseComponent : GH_Component
    {
        /// <summary>
        /// Spreadsheet Id.
        /// </summary>
        public string SpreadsheetId { get; private set; }

        /// <summary>
        /// Sheet name.
        /// </summary>
        public string SheetName { get; private set; }

        /// <summary>
        /// Range of cells in A1 notaton.
        /// </summary>
        public string CellRange { get; private set; }

        /// <summary>
        /// Range of cells with given sheet name in A1 notation on a spreadsheet.
        /// If range include only sheet name, its refer all cells in named sheet.
        /// If range include only cell range its refer cells of the first visible sheet.
        /// </summary>
        public string SpreadsheetRange => SheetsUtilities.GetSpreadsheetRange(SheetName, CellRange);

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

            if (DA.GetData(1, ref sheetName))
                SheetName = sheetName;

            if (DA.GetData(2, ref range))
                CellRange = range.ToUpper();

            if (string.IsNullOrEmpty(SpreadsheetRange))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "For correct request at least need sheet name or cells range");
                return;
            }
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
