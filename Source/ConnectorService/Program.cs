using ConnectorService.Extensions;
using ConnectorService.Models;
using ConnectorService.Api;
using System.Reflection;
using Azure.Identity;
using NSwag.Generation.Processors.Security;
using NSwag;

var builder = WebApplication.CreateBuilder(args);

// If using KeyVault to store ClientId, PrivateKey and ApiKey
var keyVaultUri = builder.Configuration["VaultUri"];
builder.Configuration.AddAzureKeyVault(new Uri(keyVaultUri), new DefaultAzureCredential());

builder.Services
    .AddConfig(builder.Configuration)
.AddDependencyGroup();

// This value needs to be injected into the ConfigurationManager, as it's used by our packages to validate the cerificate.
System.Configuration.ConfigurationManager.AppSettings["SuperIdCertificate"] = "16b7fb8c3f9ab06885a800c64e64c97c4ab5e98c";

builder.Services.AddOpenApiDocument(config =>
{
    // Define API Key Security Scheme
    config.AddSecurity("ApiKey", Enumerable.Empty<string>(), new OpenApiSecurityScheme
    {
        Type = OpenApiSecuritySchemeType.ApiKey,
        Name = "X-Api-Key", // Header name
        In = OpenApiSecurityApiKeyLocation.Header,
        Description = "Enter your API key to authenticate."
    });

    config.OperationProcessors.Add(new AspNetCoreOperationSecurityScopeProcessor("ApiKey"));

    config.OperationProcessors.Add(new DynamicFileListProcessor("Resources"));
});


var app = builder.Build();

app.AddOpenApiUi();

app
    .AddWcfEndpoints()
    .EnableWsdlGet();

///Fix to load the SuperOffice.EIS.TestConnector.dll, as it needs to be a loaded assembly before the ERPConnectorWS.cs tries to use it.
Assembly.LoadFrom(Path.Combine(AppContext.BaseDirectory, "ErpConnector.dll"));

app.AddExcelHandlerEndpoints();

app.Use(async (context, next) =>
{
    // Allow unrestricted access to the root path
    if ((context.Request.Path == "/") || (context.Request.Path.Value.Contains("custom.js")))
    {
        await next();
        return;
    }

    var providedApiKey = context.Request.Headers["X-Api-Key"].FirstOrDefault();
    var expectedApiKey = builder.Configuration[$"{ConnectorServiceOptions.ConnectorService}:ApiKey"];

    if (string.IsNullOrEmpty(expectedApiKey))
    {
        context.Response.StatusCode = StatusCodes.Status500InternalServerError;
        await context.Response.WriteAsync("ApiKey is missing in configuration.");
        return;
    }

    if (string.IsNullOrEmpty(providedApiKey) || providedApiKey != expectedApiKey)
    {
        context.Response.StatusCode = StatusCodes.Status401Unauthorized;
        await context.Response.WriteAsync("Invalid API Key.");
        return;
    }

    await next();
});

app.Run();
