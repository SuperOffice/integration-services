using ConnectorService.Models;
using ConnectorService.Models.Excel;
using ConnectorService.Utils;
using Ical.Net.CalendarComponents;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using NSwag.Generation;
using SuperOffice.Data;
using SuperOffice.Util;
using System;
using static SuperOffice.Configuration.ConfigFile;

namespace ConnectorService.Api
{
    public static class ExcelHandlerEndpoints
    {
        private static string _resourcesPath = "Resources";
        private const string _templateName = "TheOneAndOnlyErpClient.xlsm";

        public static void AddExcelHandlerEndpoints(this WebApplication app)
        {
            app.MapGet("/", () => "Starter page for the Connector Service!").ExcludeFromDescription();
            app.MapGet("/custom.js", () =>
            {
                var js = File.ReadAllText("Swagger/custom.js"); // Adjust path as needed
                return Results.Content(js, "application/javascript");
            }).ExcludeFromDescription();

            //app.MapGet("/read-capabilities/{filename}", GetAllCapabilities);
            //app.MapPost("/write-tofile/{filename}", WriteToFile);
            //app.MapGet("/read-cell-from-sheet/{filename}", ReadSheetCell);
            //app.MapGet("/{fileName}/capability/{index}", GetCapabilityByIndex);
            //app.MapPost("/{fileName}/capability", PostCapability);
            //app.MapPut("/{fileName}/capability", PutCapability);

            app.MapPost("/file/{fileName}", UploadFile).DisableAntiforgery();
            app.MapGet("/file/{fileName}", DownloadFile);
            app.MapPut("/file/{fileName}", UpdateOrCreateFile).DisableAntiforgery();
            app.MapDelete("/file/{fileName}", DeleteFile);
            app.MapGet("/files", GetAllFiles);
        }

        public static IResult GetAllCapabilities([FromServices] IExcelHandler excelHandler, string fileName = _templateName)
        {
            return Results.Text(excelHandler.GetAllCapabilties(fileName), contentType: "application/json");
        }

        public static IResult WriteToFile([FromServices] IExcelHandler excelHandler, string fileName = _templateName)
        {
            excelHandler.WriteToExcelSheet(fileName);
            return Results.Json(new { success = "OK" });

        }

        public static IResult ReadSheetCell([FromServices] IExcelHandler excelHandler, string fileName = _templateName, string sheetName = "Sheet 2", int row = 1, int column = 1)
        {
            return Results.Ok(excelHandler.ReadSheetCell(fileName, sheetName, row, column));
        }

        private static IResult GetCapabilityByIndex([FromServices] IExcelHandler excelHandler, int index, string fileName = _templateName)
        {
            return Results.Text(excelHandler.GetCapabilityRowByIndex(fileName, index), contentType: "application/json");
        }

        private static IResult PostCapability([FromServices] IExcelHandler excelHandler, Capabilities capability, string fileName = _templateName)
        {
            excelHandler.InsertCapabilities(fileName);
            return Results.Created($"/{fileName}/capability", capability);
        }

        private static IResult PutCapability([FromServices] IExcelHandler excelHandler, Capabilities capability, string fileName = _templateName)
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

                var counter = 1;
                while (File.Exists(filePath))
                {
                    filePath = Path.Combine(_resourcesPath, $"{Path.GetFileNameWithoutExtension(fileName)}_{counter}{Path.GetExtension(fileName)}");
                    counter++;
                }

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return Results.Created();
            }
            return Results.BadRequest("Invalid file.");
        }

        private static IResult DownloadFile(string fileName = _templateName)
        {
            var validationResult = ValidateFilePath(fileName, out var filePath);
            if (validationResult != null) return validationResult;

            var provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(fileName, out string contentType))
            {
                contentType = "application/octet-stream"; // Default fallback
            }

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return Results.File(stream, contentType, fileName);
        }

        private static async Task<IResult> UpdateOrCreateFile(IFormFile file, string fileName)
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

                var message = File.Exists(filePath) ? "File updated successfully." : "File created successfully.";
                return Results.Ok(new { FilePath = filePath, Message = message });
            }

            return Results.BadRequest("Invalid file.");
        }

        private static IResult DeleteFile(string fileName)
        {
            var validationResult = ValidateFilePath(fileName, out var filePath);
            if (validationResult != null) return validationResult;

            File.Delete(filePath);
            return Results.NoContent();
        }

        private static IResult GetAllFiles()
        {
            var files = Directory.GetFiles(_resourcesPath)
                .Select(file => Path.GetFileName(file));

            return Results.Json(files);
        }

        private static IResult ValidateFilePath(string fileName, out string filePath)
        {
            if (string.IsNullOrEmpty(_resourcesPath))
            {
                filePath = null;
                return Results.Problem("Resources path is not configured.", statusCode: 500);
            }

            filePath = Path.Combine(_resourcesPath, fileName);
            if (!File.Exists(filePath))
            {
                return Results.NotFound($"File not found: {filePath}");
            }

            return null; // No errors found
        }

    }
}
