using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

using Data = Google.Apis.Sheets.v4.Data;

namespace BookWorm.Request
{
    public class UpdateCells : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the UpdateCells class.
        /// </summary>
        public UpdateCells()
          : base("UpdateCells",
                "Nickname",
                "Description",
                "BookWorm",
                "Request")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Cells", "C", "Cells in a Rows", GH_ParamAccess.tree);
            pManager.AddTextParameter("Range", "Rng", "Grid Range in a1 notation", GH_ParamAccess.item);
            pManager.AddTextParameter("Fields", "F", "Field Mask", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item);

        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("debug rows", "R", "debug rows", GH_ParamAccess.list);

            // что-то возвращает вообще?
        }

        /// <inheritdoc/>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var cells = new GH_Structure<IGH_Goo>();

            var rows = new List<RowData>();

            var a1NotatonRange = string.Empty;

            var gridRange = new GridRange();

            // Нужен метод, переводящий а1-нотацию ренджа в грид рендж. И метод, что получает щит айди.
            // это пример для листа "Лист лист" (его айди 420837689) тыблицы 1jbaOPPZVP5nyDE-QCvQtBNV5eBMV6PDvZfyrDdtQ9xg
            gridRange.SheetId = 420837689;
            gridRange.StartRowIndex = 2;
            gridRange.EndRowIndex = 5;
            gridRange.StartColumnIndex = 2;
            gridRange.EndColumnIndex = 5;

            string fieldMask = string.Empty;

            var run = false;

            if (!DA.GetDataTree(0, out cells))
            {
                return;
            }

            if (!DA.GetData(1, ref a1NotatonRange))
            {
                return;
            }

            if (!DA.GetData(2, ref fieldMask))
            {
                return;
            }

            DA.GetData(3, ref run);

            if (!run)
            {
                return;
            }

            foreach (var branch in cells.Branches)
            {
                RowData row = new RowData();
                var cellsData = branch.Select(ghCell => (ghCell as GH_CellData).Value).ToList();
                row.Values = cellsData;
                rows.Add(row);
            }

            // айди листа для запроса
            var spreadsheetId = string.Empty;
            spreadsheetId = "1jbaOPPZVP5nyDE-QCvQtBNV5eBMV6PDvZfyrDdtQ9xg";

            // A list of updates to apply to the spreadsheet.
            // Requests will be applied in the order they are specified.
            // If any request is not valid, no requests will be applied.
            var requests = new List<Data.Request>();

            // Задаётся сам запрос, а потом запрос и запрос вставляется в запрос, запросом погоняет, по фазам лун юпитера с учётом силы кориолиса
            // UUUUUUSOOOOQQUUUAAAA
            var updateCellRequest = new Data.Request();

            // UUUUUUSOOOOQQUUUAAAA
            var updCellReq = new UpdateCellsRequest
            {
                Rows = rows,
                Fields = fieldMask,
                Range = gridRange,
            };

            updateCellRequest.UpdateCells = updCellReq;
            

            requests.Add(updateCellRequest);

            // Главный запрос, в который встраивается список нужных запросов
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Utilities.Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();

            DA.SetDataList(0, rows);
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

        /// <inheritdoc/>
        public override Guid ComponentGuid
        {
            get { return new Guid("e0bc3449-e003-4bf2-aefa-fffafa239645"); }
        }
    }
}