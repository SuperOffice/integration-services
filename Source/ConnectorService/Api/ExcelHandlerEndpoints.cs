using ConnectorService.Models;
using ConnectorService.Models.Excel;
using ConnectorService.Utils;
using Ical.Net.CalendarComponents;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using SuperOffice.Data;
using SuperOffice.Util;
using System;

namespace ConnectorService.Api
{
    public static class ExcelHandlerEndpoints
    {
        private static string _resourcesPath = "Resources";
        public static void AddExcelHandlerEndpoints(this WebApplication app)
        {
            app.MapGet("/", () => "Starter page for the Connector Service!");
            app.MapGet("/read-capabilities/{filename}", GetAllCapabilities);
            app.MapPost("/write-tofile/{filename}", WriteToFile);
            app.MapGet("/read-cell-from-sheet/{filename}", ReadSheetCell);
            app.MapGet("/{fileName}/capability/{index}", GetCapabilityByIndex);
            app.MapPost("/{fileName}/capability", PostCapability);
            app.MapPut("/{fileName}/capability", PutCapability);
            app.MapPost("/upload", UploadFile).DisableAntiforgery();
            app.MapGet("/download/{fileName}", DownloadFileStream);
        }

        public static IResult GetAllCapabilities([FromServices] IExcelHandler excelHandler, string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            return Results.Text(excelHandler.GetAllCapabilties(fileName), contentType: "application/json");
        }

        public static IResult WriteToFile([FromServices] IExcelHandler excelHandler, string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            excelHandler.WriteToExcelSheet(fileName);
            return Results.Json(new { success = "OK" });

        }

        public static IResult ReadSheetCell([FromServices] IExcelHandler excelHandler, string fileName = "ExcelConnectorWithCapabilities.xlsx", string sheetName = "Sheet 2", int row = 1, int column = 1)
        {
            return Results.Ok(excelHandler.ReadSheetCell(fileName, sheetName, row, column));
        }

        private static IResult GetCapabilityByIndex([FromServices] IExcelHandler excelHandler, int index, string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            return Results.Text(excelHandler.GetCapabilityRowByIndex(fileName, index), contentType: "application/json");
        }

        private static IResult PostCapability([FromServices] IExcelHandler excelHandler, Capabilities capability, string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            excelHandler.InsertCapabilities(fileName);
            return Results.Created($"/{fileName}/capability", capability);
        }

        private static IResult PutCapability([FromServices] IExcelHandler excelHandler, Capabilities capability, string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            return Results.Json(capability);
        }

        private static async Task<IResult> UploadFile(IFormFile file, string fileName)
        {
            if (fileName == null)
            {
                fileName = file.FileName;
            }

            if (file.Length > 0)
            {
                var filePath = Path.Combine(_resourcesPath, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return Results.Ok(new { FilePath = filePath });
            }
            return Results.BadRequest("Invalid file.");
        }

        private static IResult DownloadFileStream(string fileName = "ExcelConnectorWithCapabilities.xlsx")
        {
            var filePath = Path.Combine(_resourcesPath, fileName);

            if (!File.Exists(filePath))
            {
                return Results.NotFound("File not found on path " + filePath);
            }

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return Results.File(stream, "application/octet-stream", fileName);
        }
    }
}
