using ConnectorService.Models;
using ConnectorService.Services;
using ConnectorService.Utils;
using CoreWCF.Configuration;
using CoreWCF.Description;

namespace Microsoft.Extensions.DependencyInjection
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddConfig(
             this IServiceCollection services, IConfiguration config)
        {
            services
                .AddOptions<ApplicationOptions>(config, ApplicationOptions.Application)
                .AddOptions<SuperIdOptions>(config, SuperIdOptions.SuperId)
                .AddOptions<ConnectorServiceOptions>(config, ConnectorServiceOptions.ConnectorService);

            // Override sensitive values with Key Vault secrets
            services.PostConfigure<ConnectorServiceOptions>(options =>
            {
                options.ClientId = config[$"{ConnectorServiceOptions.ConnectorService}:ClientId"]
                                   ?? options.ClientId;  // Preserve existing value if secret is missing

                options.PrivateKeyFile = config[$"{ConnectorServiceOptions.ConnectorService}:PrivateKeyFile"]
                                         ?? options.PrivateKeyFile;
            });

            return services;
        }

        public static IServiceCollection AddDependencyGroup(
             this IServiceCollection services)
        {
            services
                .AddTransient<QuoteConnectorWS>()
                .AddTransient<ErpConnectorWS>()
                .AddCoreWcfDepedency()
                .AddSingleton<IExcelHandler, ExcelHandler>();

            return services;
        }

        private static IServiceCollection AddOptions<TOption>(this IServiceCollection services, IConfiguration config, string section) where TOption : class
        {
            services.Configure<TOption>(
                config.GetSection(section));

            return services;
        }

        private static IServiceCollection AddCoreWcfDepedency(this IServiceCollection services)
        {
            services.AddServiceModelServices()
                    .AddServiceModelMetadata()
                    .AddSingleton<IServiceBehavior, UseRequestHeadersForMetadataAddressBehavior>();
            return services;
        }
    }
}

