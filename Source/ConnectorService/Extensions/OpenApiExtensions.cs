using ConnectorService.Models;
using CoreWCF.Configuration;
using Microsoft.Extensions.Options;
using static SuperOffice.Configuration.ConfigFile;

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
                config.DocumentName = "ConnectorServiceAPI";
                config.Title = "ConnectorServiceAPI v1";
                config.Version = "v1";
            });

            return services;
        }

        internal static WebApplication AddOpenApiUi(this WebApplication app)
        {
            //if (app.Environment.IsDevelopment())
            //{
                app.UseOpenApi();
                app.UseSwaggerUi(config =>
                {
                    config.DocumentTitle = "ConnectorServiceAPI";
                    config.Path = "/swagger";
                    config.DocumentPath = "/swagger/{documentName}/swagger.json";
                    config.DocExpansion = "list";
                });
            //}
            return app;
        }
    }
}
