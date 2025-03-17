using ConnectorService.Models;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using System.Drawing;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Globalization;
using System;
using ConnectorService.Models.Excel;
using AngleSharp.Text;
using System.Collections.Generic;
using Aspose.Words.Tables;
using Aspose.Words.XAttr;
using Aspose.Words.Lists;
using System.Reflection;
using AngleSharp.Common;
using static SuperOffice.CRM.ArchiveLists.SaintRestrictionExtenderBase;
using System.Reflection.Metadata.Ecma335;

namespace ConnectorService.Utils
{
    public class ExcelHandler : IExcelHandler
    {
        private readonly ApplicationOptions _applicationOptions;

        private readonly JsonSerializerOptions _jsonSerializerOptions;
        public ExcelHandler(IOptions<ApplicationOptions> applicationOptions)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _applicationOptions = applicationOptions.Value;
            _jsonSerializerOptions = new JsonSerializerOptions { WriteIndented = true };

        }

        private List<T> SheetToCollectionList<T>(string fileName)
        {
            string filePath = Path.Combine(_applicationOptions.ResourcesPath, fileName);
            using var p = new ExcelPackage(filePath);
            ExcelWorksheet sheet = p.Workbook.Worksheets[typeof(T).Name]; // Get the sheet with the same name as the class

            EnsureHeadersAreValid(sheet); // Set "MISSING HEADER" for empty headers before passing it to ToCollection<T>

            int totalRows = sheet.Dimension.End.Row;
            int totalCols = sheet.Dimension.End.Column;

            var sheetData = sheet.Cells[1, 1, totalRows, totalCols]; //[FromRow, FromCol, ToRow, ToCol]

            var collection = sheetData.ToCollection<T>();
            return collection;
        }

        private static void EnsureHeadersAreValid(ExcelWorksheet worksheet)
        {
            int colCount = worksheet.Dimension.Columns;

            for (int col = 1; col <= colCount; col++)
            {
                var headerCell = worksheet.Cells[1, col];
                if (string.IsNullOrWhiteSpace(headerCell.Text))
                {
                    headerCell.Value = "MISSING HEADER";
                }
            }
        }

        public string GetAllCapabilties(string fileName)
        {
            var capabilityItems = SheetToCollectionList<Capabilities>(fileName);
            return JsonSerializer.Serialize(capabilityItems, _jsonSerializerOptions);
        }

        public string GetCapabilityRowByIndex(string fileName, int rowId)
        {
            var capabilityItems = SheetToCollectionList<Capabilities>(fileName);
            return JsonSerializer.Serialize(capabilityItems[rowId - 2], _jsonSerializerOptions);
        }

        public string InsertCapabilities(string fileName)
        {
            string filePath = Path.Combine(_applicationOptions.ResourcesPath, fileName);
            using var p = new ExcelPackage(filePath);
            ExcelWorksheet sheet = p.Workbook.Worksheets[typeof(Capabilities).Name]; // Get the sheet with the same name as the class

            int totalRows = sheet.Dimension.End.Row;

            List<Capabilities> capabilities = new List<Capabilities>();

            capabilities.Add(new Capabilities
            {
                Key = "New Capability",
                Value = true,
                Column_d = "New Status",
                Column_e = "New Owner",
                Column_f = "blabla",
                Column_g = "blabla",
                Column_h = "blabla"
            });

            sheet.Cells[totalRows, 1].LoadFromCollection<Capabilities>(capabilities);
            p.Save();

            return "yes";
        }


        public void WriteToExcelSheet(string fileName)
        {
            string filePath = Path.Combine(_applicationOptions.ResourcesPath, fileName);
            using var p = new ExcelPackage(filePath);

            var sheet = p.Workbook.Worksheets.Add("Sheet 2");
            // FillNumber will add 1, 2, 3, etc in each cell of the range
            sheet.Cells["A1:A5"].FillNumber(x => x.StartValue = 1);
            // Add two more columns with shared formula that refers to eachother.
            sheet.Cells["B1:B5"].Formula = "A$1:A$5 + 1";
            sheet.Cells["C1:C5"].Formula = "B$1:B$5 + 1";
            sheet.Cells["A1:C5"].Style.Fill.SetBackground(Color.LightYellow);
            p.Save();
        }

        public string ReadSheetCell(string fileName, string sheetName, int row, int column)
        {
            string filePath = Path.Combine(_applicationOptions.ResourcesPath, fileName);
            using var p = new ExcelPackage(filePath);

            ExcelWorksheet sheet = p.Workbook.Worksheets[sheetName];
            var cellValue = ReadCell(sheet, row, column);
            return cellValue?.ToString() ?? string.Empty;
        }

        private static object ReadCell(ExcelWorksheet sheet, int row, int col)
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

        //ExcelPackage GetExcelPackage(string fileName)
        //{
        //    string filePath = Path.Combine(_applicationOptions.BaseFilePath, fileName);
        //    return new ExcelPackage(filePath);
        //}
    }
}
