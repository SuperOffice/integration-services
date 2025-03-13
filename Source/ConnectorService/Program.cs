using System.Runtime;
using ConnectorService.Extensions;
using ConnectorService.Models;
using ConnectorService.Services;
using ConnectorService.Utils;
using CoreWCF;
using CoreWCF.Configuration;
using CoreWCF.Description;
using Ical.Net.CalendarComponents;
using System.Text.Json;
using Microsoft.Extensions.Options;
using ConnectorService.Api;
using Microsoft.AspNetCore.Mvc.ApplicationParts;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

builder.Services
    .AddConfig(builder.Configuration)
.AddDependencyGroup();

// This value needs to be injected into the ConfigurationManager, as it's used by our packages to validate the cerificate.
System.Configuration.ConfigurationManager.AppSettings["SuperIdCertificate"] = "16b7fb8c3f9ab06885a800c64e64c97c4ab5e98c";

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddOpenApiDocument(config =>
{
    config.DocumentName = "ConnectorServiceAPI";
    config.Title = "ConnectorServiceAPI v1";
    config.Version = "v1";
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseOpenApi();
    app.UseSwaggerUi(config =>
    {
        config.DocumentTitle = "ConnectorServiceAPI";
        config.Path = "/swagger";
        config.DocumentPath = "/swagger/{documentName}/swagger.json";
        config.DocExpansion = "list";
    });
}
app
    .AddWcfEndpoints()
    .EnableWsdlGet();

///Fix to load the SuperOffice.EIS.TestConnector.dll, as it needs to be a loaded assembly before the ERPConnectorWS.cs tries to use it.
Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SuperOffice.EIS.TestConnector.dll"));

app.AddExcelHandlerEndpoints();

app.Run();
