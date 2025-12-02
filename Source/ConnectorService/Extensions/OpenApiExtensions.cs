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
