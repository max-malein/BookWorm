using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using System.Linq;

namespace GoogleDocs.Spreadsheets
{
    public class ReadCellRange : GH_Component
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Bookworm";

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public ReadCellRange()
          : base(
                "ReadCellRange",
                "ReadCell",
                "Reads a range of cells",
                "BookWorm",
                "Spreadsheet")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("SpreadsheetId", "Id", "Spreadsheet Id", GH_ParamAccess.item);

            pManager.AddTextParameter("SheetName", "N", "Sheet Name", GH_ParamAccess.item);

            pManager.AddTextParameter("CellRange", "C", "Range of cells", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Read", "R", "Read data from spreadsheet", GH_ParamAccess.item, false);

            pManager.AddBooleanParameter("SameLength", "S", "All output rows will be the same length", GH_ParamAccess.item, false);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Values", "V", "Values", GH_ParamAccess.tree);
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetId = string.Empty;
            string sheetName = string.Empty;
            string range = string.Empty;
            bool read = false;
            bool sameLength = false;

            if (!DA.GetData(0, ref spreadsheetId)) return;
            if (!DA.GetData(1, ref sheetName)) return;
            if (!DA.GetData(2, ref range)) return;
            DA.GetData(3, ref read);
            DA.GetData(4, ref sameLength);

            if (!read) return;

            GH_AssemblyInfo info = Grasshopper.Instances.ComponentServer.FindAssembly(new Guid("56dfe1a3-4e7b-425f-b169-965c0d1f7977"));
            string assemblyLocation = Path.GetDirectoryName(info.Location);

            UserCredential credential;

            using (var stream =
                new FileStream(Path.Combine(assemblyLocation, "credentials.json"), FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = Path.Combine(assemblyLocation, "token.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            // Single quotes for cases with space between sheet name parts.
            string requestRange = $"'{sheetName}'!{range}";

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, requestRange);

            ValueRange response = request.Execute();

            IList<IList<object>> values = response.Values;

            if (values != null && values.Count > 0)
            {
                var data = new GH_Structure<GH_String>();
                for (int i = 0; i < values.Count; i++)
                {
                    var path = new GH_Path(i);
                    var ghStrings = values[i].Select(s => new GH_String(s.ToString()));

                    data.AppendRange(ghStrings, path);
                }

                //одинаковые длины для списков
                if (sameLength)
                {
                    var maxLength = values.Max(r => r.Count);
                    foreach (var path in data.Paths)
                    {
                        int valuesToAdd = maxLength - data[path].Count;
                        if (valuesToAdd > 0)
                        {
                            var emptys = Enumerable.Repeat(new GH_String(string.Empty), valuesToAdd);
                            data[path].AddRange(emptys);
                        }
                    }
                }

                DA.SetDataTree(0, data);
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No data found.");
            }
        }

        /// <inheritdoc/>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it.
        /// It is vital this Guid doesn't change otherwise old ghx files
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5042202c-269f-4617-8b7c-fcb847d56ff6"); }
        }
    }
}
