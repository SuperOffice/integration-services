using ConnectorService.Models;
using ConnectorService.Services;
using ConnectorService.Utils;
using CoreWCF.Configuration;
using CoreWCF.Description;
using static SuperOffice.Configuration.ConfigFile;

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
                .AddOptions<WcfOptions>(config, WcfOptions.Wcf)
                //.AddOptions<FileConfigDataStoreOptions>(config, FileConfigDataStoreOptions.FileConfigDataStore)
                .AddOptions<ErpConnectorOptions>(config, ErpConnectorOptions.ErpConnector);

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

