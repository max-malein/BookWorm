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

namespace GoogleDocs.Spreadsheets
{
    public class WriteCellRange : GH_Component
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets reader";

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public WriteCellRange()
          : base("WriteCellRange", "WriteCell",
              "Writes a range of cells",
              "BookWorm", "Spreadsheet")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("SpreadsheetId", "Id", "Spreadsheet Id", GH_ParamAccess.item);
            pManager.AddTextParameter("SheetName", "N", "Sheet Name", GH_ParamAccess.item);
            pManager.AddTextParameter("CellRange", "C", "Starting cell", GH_ParamAccess.item);
            pManager.AddTextParameter("Values", "V", "Values", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Append", "A", "If true, values will be added to the next possible empty row", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Write", "W", "Write data from spreadsheet", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Result", "R", "Result", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheetId = string.Empty;
            string sheet = string.Empty;
            string range = string.Empty;
            var inputData = new List<string>();
            bool write = false;
            bool append = false;

            if (!DA.GetData(0, ref spreadsheetId)) return;
            if (!DA.GetData(1, ref sheet)) return;
            if (!DA.GetData(2, ref range)) return;
            if (!DA.GetDataList(3, inputData)) return;
            DA.GetData(4, ref append);
            DA.GetData(5, ref write);

            if (!write) return;

            UserCredential credential;

            GH_AssemblyInfo info = Grasshopper.Instances.ComponentServer.FindAssembly(new Guid("56dfe1a3-4e7b-425f-b169-965c0d1f7977"));
            string assemblyLocation = Path.GetDirectoryName(info.Location);

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
            String requestRange = $"{sheet}!{range}";

            var valueRange = new ValueRange
            {
                MajorDimension = "ROWS",
                Range = requestRange,
                Values = new List<IList<object>>() { new List<object>() }
            };

            //WTF???
            //you need to explicitly set every value to string, otherwise doesn't work
            for (int i = 0; i < inputData.Count; i++)
            {
                valueRange.Values[0].Add(inputData[i]);
            }

            if (append)
            {
                var request = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, requestRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var result = request.Execute();
                DA.SetData(0, result.Updates.ToString());
            }
            else
            {
                var request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, requestRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var result = request.Execute();
                DA.SetData(0, result.ToString());
            }
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                //return Resources.IconForThisComponent;
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
            get { return new Guid("669A9C19-1370-4613-AC10-1D3264E3D290"); }
        }
    }
}
