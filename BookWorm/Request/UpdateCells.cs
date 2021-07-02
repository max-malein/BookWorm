using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
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

            pManager.AddTextParameter("Fields", "F", "Field Mask that applies for cells update", GH_ParamAccess.item, "*");

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

            gridRange = GridRangeFromA1(a1NotatonRange, (int)sheetId, cells.Count);

            rows = GetRows(cells, gridRange);

            // A list of updates to apply to the spreadsheet.
            // Requests will be applied in the order they are specified.
            // If any request is not valid, no requests will be applied.
            var requests = new List<Data.Request>();

            // New and empty request instance.
            var updateCellRequest = new Data.Request();

            var updateCell = new UpdateCellsRequest
            {
                Rows = rows,
                Fields = fieldMask,
                Range = gridRange,
            };

            updateCellRequest.UpdateCells = updateCell;

            requests.Add(updateCellRequest);

            // Main request of matrioshka-request.
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns>Column number for column name.</returns>
        private int ColumnNameToNumber(string columnName)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                char letter = columnName[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }

            return result;
        }

        /// <summary>
        /// Creates GridRange object from A1 notation range.
        /// </summary>
        /// <param name="a1NotatonRange">Cells range in A1 notation.</param>
        /// <param name="sheetId">Sheet Id.</param>
        /// <param name="cellsCount">Cells count.</param>
        /// <returns>new Grid Range.</returns>
        private GridRange GridRangeFromA1(string a1NotatonRange, int sheetId, int cellsCount)
        {
            var gridRange = new GridRange();
            gridRange.SheetId = sheetId;

            bool letters = Regex.Matches(a1NotatonRange, @"[a-zA-Z]").Count > 0;
            bool numbers = a1NotatonRange.Any(c => char.IsDigit(c));

            var rangeBounds = a1NotatonRange.Split(':');

            // Only columns case.
            // сломается при одной колонке.
            if (letters && !numbers)
            {
                List<int> colNumbers = new List<int>();

                foreach (string colName in rangeBounds)
                {
                    colNumbers.Add(ColumnNameToNumber(colName));
                }

                gridRange.StartColumnIndex = colNumbers[0] - 1;
                gridRange.EndColumnIndex = colNumbers[1];

                var colCount = colNumbers[1] - colNumbers[0] + 1;
                gridRange.StartRowIndex = 0;
                // ошибка при колонках? или должны заполняться поровну колонки?
                gridRange.EndRowIndex = cellsCount / colCount;
            }

            // Only rows case.
            // сломается при одной строке.
            else if (numbers && !letters)
            {
                var startRow = Convert.ToInt32(rangeBounds[0]);
                var endRow = Convert.ToInt32(rangeBounds[1]);
                gridRange.StartRowIndex = startRow - 1;
                gridRange.EndRowIndex = endRow;

                var rowCount = endRow - startRow + 1;
                gridRange.StartColumnIndex = 0;
                gridRange.EndColumnIndex = cellsCount / rowCount;
            }

            // Coordinate case.
            // сломается при одной ячейке.
            else
            {
                var firstColName = Regex.Match(rangeBounds[0], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                gridRange.StartColumnIndex = ColumnNameToNumber(firstColName) - 1;

                var secondColName = Regex.Match(rangeBounds[1], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                gridRange.EndColumnIndex = ColumnNameToNumber(secondColName);

                gridRange.StartRowIndex = Convert.ToInt32(Regex.Match(rangeBounds[0], @"\d+", RegexOptions.IgnoreCase).Value) - 1;
                gridRange.EndRowIndex = Convert.ToInt32(Regex.Match(rangeBounds[1], @"\d+", RegexOptions.IgnoreCase).Value);
            }

            return gridRange;
        }

        /// <summary>
        /// Splits cells list in rows by grid range.
        /// </summary>
        /// <param name="cells">Cells.</param>
        /// <param name="gridRange">Grid range.</param>
        /// <returns>Rows.</returns>
        private List<RowData> GetRows(List<CellData> cells, GridRange gridRange)
        {
            if (gridRange == null)
                return null;

            var rowLength = gridRange.EndColumnIndex - gridRange.StartColumnIndex;

            var rows = new List<RowData>();
            var row = new RowData();
            var rowValues = new List<CellData>();

            for (int i = 0; i < cells.Count; i++)
            {
                rowValues.Add(cells[i]);

                if ((i + 1) % rowLength == 0)
                {
                    row.Values = rowValues;
                    rows.Add(row);

                    row = new RowData();
                    rowValues = new List<CellData>();
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