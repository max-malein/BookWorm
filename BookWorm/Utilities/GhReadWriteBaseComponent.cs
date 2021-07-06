using System;
using System.Windows.Forms;
using Grasshopper.Kernel;

namespace BookWorm.Utilities
{
    public class GhReadWriteBaseComponent : GH_Component
    {
        private string spreadsheetId;

        public GhReadWriteBaseComponent(string name, string nick, string desc, string tab, string subTab)
            : base(name, nick, desc, tab, subTab)
        {
        }

        public override Guid ComponentGuid => throw new NotImplementedException();

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet URL", "U", "Google spreadsheet URL or spreadsheet ID", GH_ParamAccess.item);
            pManager.AddTextParameter("Sheet Name", "N", "Sheet Name", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Range", "C", "Range of cells", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetUrl = string.Empty;
            string sheetName = string.Empty;
            string range = string.Empty;
            spreadsheetId = string.Empty;

            if (!DA.GetData(0, ref spreadsheetUrl)) return;
            spreadsheetId = Utilities.Util.ParseUrl(spreadsheetUrl);

            if (!DA.GetData(1, ref sheetName)) return;
            if (!DA.GetData(2, ref range)) return;
        }

        /// <inheritdoc/>
        public override void AppendAdditionalMenuItems(ToolStripDropDown menu)
        {
            Menu_AppendSeparator(menu);
            Menu_AppendItem(menu, "Open spreadsheet in browser...", ElementClicked, true, false);
        }

        private void ElementClicked(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(spreadsheetId))
            {
                System.Diagnostics.Process.Start(@"https://docs.google.com/spreadsheets/d/" + spreadsheetId);
            }
        }
    }
}
