using ConnectorService.Extensions;
using ConnectorService.Models;
using ConnectorService.Api;
using System.Reflection;
using Azure.Identity;

var builder = WebApplication.CreateBuilder(args);

// If using KeyVault to store ClientId, PrivateKey and ApiKey
var keyVaultUri = builder.Configuration["VaultUri"];
if (!string.IsNullOrWhiteSpace(keyVaultUri))
{
    builder.Configuration.AddAzureKeyVault(new Uri(keyVaultUri), new DefaultAzureCredential());
}

builder.Services
    .AddConfig(builder.Configuration)
    .AddDependencyGroup()
    .AddOpenApi();

// The quote connector base class builds a fallback validator from this setting and throws if it is missing; tokens are validated by SuperOfficeJwksTokenValidator.
System.Configuration.ConfigurationManager.AppSettings["SuperIdCertificate"] = "16b7fb8c3f9ab06885a800c64e64c97c4ab5e98c";

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
    // WCF endpoints are authenticated with SuperOffice-signed tokens, not the API key
    if ((context.Request.Path == "/")
        || (context.Request.Path.Value.Contains("custom.js"))
        || context.Request.Path.StartsWithSegments("/Services"))
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
