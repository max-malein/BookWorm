﻿using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
            pManager.AddTextParameter("SpreadsheetId", "Id", "Spreadsheet Id or spreadsheet url.", GH_ParamAccess.item);
            pManager.AddTextParameter("SheetName", "N", "Sheet name. If such sheet doesn't exist it will be created.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Cells", "C", "Cells in a Rows", GH_ParamAccess.list);
            pManager.AddTextParameter("Range", "Rng", "Grid Range in a1 notation", GH_ParamAccess.item);
            pManager.AddTextParameter("Fields", "F", "Field Mask", GH_ParamAccess.item, "*");
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
            var cellsGoo = new List<GH_CellData>();

            var rows = new List<RowData>();

            var a1NotatonRange = string.Empty;

            var gridRange = new GridRange();

            string spreadsheetId = string.Empty;
            string sheetName = string.Empty;
            string fieldMask = string.Empty;

            var run = false;

            if (DA.GetData(0, ref spreadsheetId)) return;

            if (DA.GetData(1, ref sheetName)) return;

            if (!DA.GetDataList(2, cellsGoo)) return;

            if (!DA.GetData(3, ref a1NotatonRange)) return;

            DA.GetData(4, ref fieldMask);

            DA.GetData(5, ref run);

            if (!run)
            {
                return;
            }

            var cells = cellsGoo.Select(c => c.Value).ToList();

            int sheetId = GetSheetId(spreadsheetId, sheetName);
            if (sheetId == -1)
            {
                sheetId = CreateNewSheet(sheetName);
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

        private int CreateNewSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Находит номер листа по его имени.
        /// </summary>
        /// <param name="spreadsheetId">Id документа.</param>
        /// <param name="sheetName">Название листа.</param>
        /// <returns>Номер листа или -1, если такого листа нет.</returns>
        private int GetSheetId(string spreadsheetId, string sheetName)
        {

        }

        private int ColumnNameToNumber(string colName)
        {
            // Return the column number for this column name.
            int result = 0;
            // Process each letter.
            for (int i = 0; i < colName.Length; i++)
            {
                result *= 26;
                char letter = colName[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result;
        }
        private GridRange GridRangeFromA1(string a1NotatonRange, int sheetId, int numberOfCells)
        {
            // это пример для листа "Лист лист" (его айди 420837689) тыблицы 1jbaOPPZVP5nyDE-QCvQtBNV5eBMV6PDvZfyrDdtQ9xg
            var gridRange = new GridRange();
            gridRange.SheetId = sheetId;

            bool letters = Regex.Matches(a1NotatonRange, @"[a-zA-Z]").Count>0;
            bool numbers = a1NotatonRange.Any(c=>Char.IsDigit(c));

            var spl = a1NotatonRange.Split(':');

            if (letters && !numbers)
            {
                List<int> colNumbers = new List<int>();

                foreach (string colName in spl)
                {
                    colNumbers.Add(ColumnNameToNumber(colName));
                }

                gridRange.StartColumnIndex = colNumbers[0]-1;
                gridRange.EndColumnIndex = colNumbers[1]-1;

                var colCount = colNumbers[1] - colNumbers[0]+1;
                gridRange.StartRowIndex = 0;
                gridRange.EndRowIndex = numberOfCells / colCount - 1;
            }
            else if(numbers && !letters)
            {
                var startRow = Convert.ToInt32(spl[0]);
                var endRow = Convert.ToInt32(spl[1]);
                gridRange.StartRowIndex = startRow - 1;
                gridRange.EndRowIndex = endRow - 1;

                var rowCount = endRow - startRow + 1;
                gridRange.StartColumnIndex = 0;
                gridRange.EndColumnIndex = numberOfCells/rowCount - 1;
            }

            else
            {
                var firstColName = Regex.Match(spl[0], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                gridRange.StartColumnIndex = ColumnNameToNumber(firstColName) - 1;

                var secondColName = Regex.Match(spl[1], @"[A-Z]+", RegexOptions.IgnoreCase).Value;
                gridRange.EndColumnIndex = ColumnNameToNumber(secondColName) - 1;

                gridRange.StartRowIndex = Convert.ToInt32(Regex.Match(spl[0], @"\d+", RegexOptions.IgnoreCase).Value);
                gridRange.EndRowIndex = Convert.ToInt32(Regex.Match(spl[1], @"\d+", RegexOptions.IgnoreCase).Value);

            }


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