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
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.HttpOverrides;

var builder = WebApplication.CreateBuilder(args);

builder.Services
    .AddConfig(builder.Configuration)
.AddDependencyGroup();

// This value needs to be injected into the ConfigurationManager, as it's used by our packages to validate the cerificate.
System.Configuration.ConfigurationManager.AppSettings["SuperIdCertificate"] = "16b7fb8c3f9ab06885a800c64e64c97c4ab5e98c";

builder.Services.AddOpenApi();

var app = builder.Build();

app.AddOpenApiUi();

app
    .AddWcfEndpoints()
    .EnableWsdlGet();

///Fix to load the SuperOffice.EIS.TestConnector.dll, as it needs to be a loaded assembly before the ERPConnectorWS.cs tries to use it.
Assembly.LoadFrom(Path.Combine(AppContext.BaseDirectory, "SuperOffice.EIS.TestConnector.dll"));

app.AddExcelHandlerEndpoints();

app.Run();
