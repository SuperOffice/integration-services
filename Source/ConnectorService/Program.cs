using System.Runtime;
using ConnectorService.Extensions;
using ConnectorService.Models;
using ConnectorService.Services;
using CoreWCF;
using CoreWCF.Configuration;
using CoreWCF.Description;

var builder = WebApplication.CreateBuilder(args);

builder.Services
    .AddConfig(builder.Configuration)
.AddDependencyGroup();

// This value needs to be injected into the ConfigurationManager, as it's used by our packages to validate the cerificate.
System.Configuration.ConfigurationManager.AppSettings["SuperIdCertificate"] = "16b7fb8c3f9ab06885a800c64e64c97c4ab5e98c";

var app = builder.Build();

app
    .AddWcfEndpoints()
    .EnableWsdlGet();

app.MapGet("/", () => "Hello World!");

app.Run();
