using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BookWorm.Utilities
{
    public class TestComponent : ReadWriteBaseComponent
    {
        public TestComponent()
            : base(
                "Test",
                "TestReadCell",
                "Reads a range of cells",
                "BookWorm",
                "Spreadsheet")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            pManager.AddBooleanParameter("Read", "R", "Read data from spreadsheet", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Same Length", "S", "All output rows will be the same length", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);

            pManager.AddTextParameter("Debug", "D", "Debug", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            base.SolveInstance(DA);

            bool read = false;
            bool same = false;

            DA.GetData("Read", ref read);
            DA.GetData("Same Length", ref same);

            DA.SetData(0, "Success");
        }

        public override Guid ComponentGuid => new Guid("2971B7BF-8FAB-450F-B552-BCFD55A65F44");
    }
}
