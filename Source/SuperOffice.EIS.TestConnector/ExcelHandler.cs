using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace SuperOffice.ErpSync.TestConnector
{
    class ExcelHandler
    {
        ExcelPackage ExcelPackage { get; set; }
        ExcelWorkbook Workbook => ExcelPackage?.Workbook;
        string _excelFilePath = "";

        public Dictionary<string, object>[] GetAllRows(string sheetName)
        {
            var sheet = GetSheet(sheetName);

            if (sheet == null)
            {
                var errorMessage = $"There is no sheet named '{sheetName}' in the Excel file '{_excelFilePath}'.";
                throw new Exception(errorMessage);
            }

            var results = new List<Dictionary<string, object>>();
            var firstRow = sheet.Dimension.Start.Row;
            var lastRow = LastRowIndex(sheet);

            for (var rw = firstRow; rw <= lastRow; rw++)
            {
                var id = GetIdByRowIndex(sheet, rw);

                if (!string.IsNullOrEmpty(id))
                    results.Add(GetRowByIndex(sheet, rw));
            }

            return results.ToArray();
        }

        private object ReadCell(ExcelWorksheet sheet, int row, int col)
        {
            try
            {
                var value = sheet.GetValue(row, col);
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Problem reading string value from row {0}, col {1} in sheet {2}", row, col, sheet.Name), ex);
            }
        }

        private Dictionary<string, int> GetColumns(ExcelWorksheet sheet)
        {
            var columns = new Dictionary<string, int>();
            var dimension = sheet.Dimension;
            if (dimension == null)
                throw new Exception($"Sheet '{sheet.Name}' in Excel file '{_excelFilePath}' did not have a dimension with data.");

            var firstColumn = dimension.Start.Column;
            var lastColumn = dimension.End.Column;
            var firstRow = dimension.Start.Row;

            for (var i = firstColumn; i <= lastColumn; i++)
            {
                var val = ReadCell(sheet, firstRow, i);
                if (val is string strVal)
                { 
                    if (!string.IsNullOrWhiteSpace(strVal) && !columns.ContainsKey(strVal))
                        columns.Add(strVal, i);
                }
            }
            return columns;
        }

        private string GetIdByRowIndex(ExcelWorksheet sheet, int index)
        {
            var columns = GetColumns(sheet);
            if (!columns.ContainsKey("ID"))
                return string.Empty;
            var val = ReadCell(sheet, index, columns["ID"]);
            return val?.ToString() ?? string.Empty;
        }

        private Dictionary<string, object> GetRowByIndex(ExcelWorksheet sheet, int rwIndex)
        {
            var columns = GetColumns(sheet);
            var result = new Dictionary<string, object>();

            foreach (var col in columns)
            {
                var val = ReadCell(sheet, rwIndex, col.Value);
                var nm = col.Key;
                result.Add(nm, val);
            }
            return result;
        }

        private int GetRowIndexById(ExcelWorksheet sheet, int id)
        {
            var columns = GetColumns(sheet);
            var lastRow = LastRowIndex(sheet);

            if (!columns.ContainsKey("ID"))
                return -1;

            for (var i = 1; i <= lastRow; i++)
            {
                var val = ReadCell(sheet, i, columns["ID"]);

                if (val == null)
                    continue;

                if (int.TryParse(val.ToString(), out var tmpId))
                    if (tmpId == id)
                        return i;
            }

            return -1;
        }

        public bool UpdateRowByID(string sheetName, string id, Dictionary<string, object> cellValues)
        {
            var sheet = GetSheet(sheetName);
            var columns = GetColumns(sheet);

            if (!int.TryParse(id, out var intId))
                return false;

            var rwIndex = GetRowIndexById(sheet, intId);

            if (rwIndex <= 0)
                return false;

            var rw = GetRowByIndex(sheet, rwIndex);

            if (rw == null)
                return false;

            foreach (var cell in cellValues)
            {
                if (rw.ContainsKey(cell.Key))
                    if (rw[cell.Key] != cell.Value)
                        rw[cell.Key] = cell.Value;
            }

            // Save row
            foreach (var col in columns)
            {
                if (rw.ContainsKey(col.Key))
                {
                    if (col.Key.ToLower() == "lastmodified")
                        sheet.Cells[rwIndex, col.Value].Value = DateTime.Now;
                    else
                    {
                        var val = rw[col.Key];

                        if (val == null)
                            val = "";

                        sheet.Cells[rwIndex, col.Value].Value = val;
                        
                    }

                }
            }
            ExcelPackage.SaveAs(_excelFilePath);
            return true;
        }

        public int NextId(string sheetName)
        {
            var sheet = GetSheet(sheetName);
            var newId = GetMaxID(sheet) + 1;

            return newId;
        }

        public int NewRow(string sheetName, Dictionary<string, object> cellValues)
        {
            var sheet = GetSheet(sheetName);
            var columns = GetColumns(sheet);
            var rwIndex = LastRowIndex(sheet) + 1;
            var newId = GetMaxID(sheet) + 1;

            foreach (var col in columns)
            {
                if (col.Key == "ID")
                {
                    sheet.Cells[rwIndex, col.Value].Value = newId;
                }
            }

            ExcelPackage.SaveAs(_excelFilePath);

            UpdateRowByID(sheetName, newId.ToString(), cellValues);

            return newId;
        }

        private int GetMaxID(ExcelWorksheet sheet)
        {
            var idColIndex = -1;
            var maxId = -1;
            var rowCount = LastRowIndex(sheet) + 1;

            // Get column index of ID column
            var columns = GetColumns(sheet);
            idColIndex = (
                from c in columns
                where c.Key == "ID"
                select c.Value).FirstOrDefault();

            if (idColIndex < 0)
                return -100;

            for (var i = 1; i <= rowCount; i++)
            {
                var val = ReadCell(sheet, i, columns["ID"]);

                if (val == null)
                    val = "";

                var tmpId = -1;
                if (int.TryParse(val.ToString(), out tmpId))
                {
                    if (tmpId > maxId)
                        maxId = tmpId;
                }
            }

            return maxId;
        }

        private int LastRowIndex(ExcelWorksheet sheet)
        {
            var idColIndex = -1;
            var rowCount = 10000;

            // Get column index of ID column
            var columns = GetColumns(sheet);
            idColIndex = (
                from c in columns
                where c.Key == "ID"
                select c.Value).FirstOrDefault();

            if (idColIndex < 0)
                return -100;

            for (var i = 1; i <= rowCount; i++)
            {
                var val = ReadCell(sheet, i, columns["ID"]);

                if (val == null)
                    val = "";

                if (string.IsNullOrEmpty(val.ToString()))
                    return i - 1;
            }

            return -100;
        }

        public Dictionary<string, object> GetRowByID(string sheetName, string id)
        {
            var sheet = GetSheet(sheetName);


            if (!int.TryParse(id, out var intId))
                return null;

            var rw = GetRowIndexById(sheet, intId);

            if (rw > 0)
                return GetRowByIndex(sheet, rw);

            return null;
        }

        public Dictionary<string, object>[] GetRowBySearchString(string sheetName, string searchString, string[] searchColumns)
        {
            var sheet = GetSheet(sheetName);
            var results = new List<Dictionary<string, object>>();

            if (searchColumns == null || searchColumns.Count() == 0)
            {
                var columns = GetColumns(sheet);
                searchColumns = columns.Select(r => r.Key).ToArray();
            }

            // Perform search
            // TODO: This is quick and dirty; we get all rows and then search them
            var allRows = GetAllRows(sheetName);

            foreach (var rw in allRows)
            {
                // If searchColumns is null or count = 0, search all columns
                foreach (var col in searchColumns)
                {
                    if (rw.ContainsKey(col))
                    {
                        var val = "";

                        if (rw[col] != null)
                            val = rw[col].ToString();

                        if (val.ToLower().Contains(searchString.ToLower()) || searchString.ToLower().Contains(val.ToLower()) && val.Length > 0)
                        {
                            results.Add(rw);
                            break;
                        }
                    }
                }
            }

            return results.ToArray();
        }

        public Dictionary<string, object>[] GetRowBySearchString(string sheetName, string searchString)
        {
            return GetRowBySearchString(sheetName, searchString, null);
        }

        public bool IsExcelOpen()
        {
            return Workbook != null;
        }

        public bool OpenExcelDoc(string fileName)
        {
            if (!fileName.EndsWith(".xlsx") && !fileName.EndsWith("xlsm") && !fileName.EndsWith(".xls"))
                throw new ArgumentException("Unrecognised or missing file extension in filename '" + fileName + "'", "Filename");

            _excelFilePath = fileName;
            using (var fileStream = File.OpenRead(fileName))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage = new ExcelPackage();
                ExcelPackage.Load(fileStream);
            }
            return true;
        }

        private ExcelWorksheet GetSheet(string sheetName)
        {
            return Workbook.Worksheets.SingleOrDefault(x => string.Equals(x.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
