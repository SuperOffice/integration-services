using NSwag;
using NSwag.Generation.Processors.Security;

namespace ConnectorService.Extensions
{
    public static class OpenApiExtensions
    {
        public static IServiceCollection AddOpenApi(
             this IServiceCollection services)
        {
            services.AddEndpointsApiExplorer();
            services.AddOpenApiDocument(config =>
            {
                config.Title = "ConnectorServiceAPI v1";
                config.Version = "v1";

                config.AddSecurity("ApiKey", Enumerable.Empty<string>(), new OpenApiSecurityScheme
                {
                    Type = OpenApiSecuritySchemeType.ApiKey,
                    Name = "X-Api-Key",
                    In = OpenApiSecurityApiKeyLocation.Header,
                    Description = "Enter your API key to authenticate."
                });

                config.OperationProcessors.Add(new AspNetCoreOperationSecurityScopeProcessor("ApiKey"));
                config.OperationProcessors.Add(new DynamicFileListProcessor("Resources"));
            });

            return services;
        }

        internal static WebApplication AddOpenApiUi(this WebApplication app)
        {
            app.UseOpenApi();
            app.UseSwaggerUi(config =>
            {
                config.DocumentTitle = "ConnectorServiceAPI";
                config.Path = "/swagger";
                config.DocumentPath = "/swagger/{documentName}/swagger.json";
                config.DocExpansion = "list";
                config.CustomJavaScriptPath = "/custom.js"; //Fetches the custom .js-file from the builtin endpoint, since we dont have access to the swagger-files directly..
            });
            return app;
        }
    }
}
