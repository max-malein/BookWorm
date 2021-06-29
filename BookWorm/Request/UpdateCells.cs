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
          : base(
                "UpdateCells",
                "Nickname",
                "Description",
                "BookWorm",
                "Request")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("SpreadsheetId", "Id", "Spreadsheet Id or spreadsheet url.", GH_ParamAccess.item);

            pManager.AddTextParameter("SheetName", "N", "Sheet name. If such sheet doesn't exist it will be created.", GH_ParamAccess.item);

            pManager.AddGenericParameter("Cells", "C", "Cells in a Rows", GH_ParamAccess.list);

            pManager.AddTextParameter("Range", "Rng", "Grid Range in a1 notation", GH_ParamAccess.item);

            pManager.AddTextParameter("Fields", "F", "Field Mask", GH_ParamAccess.item, "*");

            pManager.AddBooleanParameter("Run", "Run", "Run", GH_ParamAccess.item, false);

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
            var cellsGoo = new List<GH_CellData>();

            var rows = new List<RowData>();

            var a1NotatonRange = string.Empty;

            var gridRange = new GridRange();

            string spreadsheetId = string.Empty;
            string sheetName = string.Empty;
            string fieldMask = string.Empty;

            var run = false;

            if (!DA.GetData(0, ref spreadsheetId)) return;

            if (!DA.GetData(1, ref sheetName)) return;

            if (!DA.GetDataList(2, cellsGoo)) return;

            if (!DA.GetData(3, ref a1NotatonRange)) return;

            DA.GetData(4, ref fieldMask);

            DA.GetData(5, ref run);

            if (!run)
            {
                return;
            }

            var cells = cellsGoo.Select(c => c.Value).ToList();

            var sheetId = GetSheetId(spreadsheetId, sheetName);

            if (sheetId == null)
            {
                sheetId = CreateNewSheet(spreadsheetId, sheetName);
            }

            gridRange = GridRangeFromA1(a1NotatonRange, sheetId, cells.Count);

            rows = GetRows(cells, gridRange);

            // для теста - потом убрать.
            spreadsheetId = "1jbaOPPZVP5nyDE-QCvQtBNV5eBMV6PDvZfyrDdtQ9xg";

            // A list of updates to apply to the spreadsheet.
            // Requests will be applied in the order they are specified.
            // If any request is not valid, no requests will be applied.
            var requests = new List<Data.Request>();

            // Задаётся сам запрос, а потом запрос и запрос вставляется в запрос, запросом погоняет, по фазам лун юпитера с учётом силы кориолиса
            // UUUUUUSOOOOQQUUUAAAA
            var updateCellRequest = new Data.Request();

            // UUUUUUSOOOOQQUUUAAAA
            var updateCell = new UpdateCellsRequest
            {
                Rows = rows,
                Fields = fieldMask,
                Range = gridRange,
            };

            updateCellRequest.UpdateCells = updateCell;

            requests.Add(updateCellRequest);

            // Главный запрос, в который встраивается список нужных запросов
            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Utilities.Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();

            DA.SetDataList(0, rows);
        }

        /// <summary>
        /// Add new sheet to the spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <param name="sheetName">Sheet Name.</param>
        /// <returns>Sheet Id of created sheet.</returns>
        private int? CreateNewSheet(string spreadsheetId, string sheetName)
        {
            var requests = new List<Data.Request>();

            var addSheetRequest = new Data.Request();

            var sheetProperties = new SheetProperties
            {
                Title = sheetName,
            };

            var addSheet = new AddSheetRequest
            {
                Properties = sheetProperties,
            };

            addSheetRequest.AddSheet = addSheet;

            requests.Add(addSheetRequest);

            var requestBody = new BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            var request = Utilities.Credentials.Service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);

            var response = request.Execute();
            var sheetId = response.Replies.FirstOrDefault().AddSheet.Properties.SheetId;

            return sheetId;
        }

        /// <summary>
        /// Get sheet Id by it's name.
        /// </summary>
        /// <param name="spreadsheetId">Spreadsheet Id.</param>
        /// <param name="sheetName">Sheet Name.</param>
        /// <returns>Sheet Id or null if sheet doesn't exist.</returns>
        private int? GetSheetId(string spreadsheetId, string sheetName)
        {
            var request = Utilities.Credentials.Service.Spreadsheets.Get(spreadsheetId);

            // For partial response. Multiple different fields are comma separated and subfields are dot-separated.
            // But if 403 error has happen use slashes.
            // Field names can be specified in camelCase or separated_by_underscores.
            // For convenience, multiple subfields from the same type can be listed within parentheses.
            // And without whitespaces.
            request.Fields = "sheets.properties(title,sheetId)";

            var spreadsheet = request.Execute();

            var sheets = spreadsheet.Sheets;

            foreach (var sheet in sheets)
            {
                if (sheet.Properties.Title == sheetName)
                {
                    return sheet.Properties.SheetId;
                }
            }

            return null;
        }

        private GridRange GridRangeFromA1(string a1NotatonRange, int? sheetId, int numberOfCells)
        {

            var gridRange = new GridRange();

            // это пример для листа "Лист лист" (его айди 420837689) тыблицы 1jbaOPPZVP5nyDE-QCvQtBNV5eBMV6PDvZfyrDdtQ9xg
            gridRange.SheetId = sheetId;
            gridRange.StartRowIndex = 2;
            gridRange.EndRowIndex = 5;
            gridRange.StartColumnIndex = 2;
            gridRange.EndColumnIndex = 5;

            return gridRange;
        }

        private List<RowData> GetRows(List<CellData> cells, GridRange gridRange)
        {
            if (gridRange == null)
                return null;

            // Y:AB - 4 значения в ряду
            var rowLength = gridRange.EndRowIndex - gridRange.StartRowIndex + 1;
            var rows = new List<RowData>();
            var row = new RowData();
            for (int i = 0; i < cells.Count; i++)
            {
                row.Values.Add(cells[i]);
                if ((i + 1) % rowLength == 0)
                {
                    rows.Add(row);
                    row = new RowData();
                }
            }

            return rows;
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