using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Swagger;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Reflection;

namespace SmartOffice.Hub.Services
{
    public static class OutlookFeatureRegistration
    {
        public const string SwaggerDocumentName = "outlook-v1";
        public const string NotificationHubRoute = "/hub/notifications";
        public const string OutlookAddinHubRoute = "/hub/outlook-addin";

        public static IServiceCollection AddOutlookFeature(this IServiceCollection services)
        {
            services.AddSingleton<MailStore>();
            services.AddSingleton<ChatStore>();
            services.AddSingleton<CommandResultStore>();
            services.AddSingleton<AddinStatusStore>();
            services.AddSingleton<AttachmentExportService>();
            services.AddSingleton<OutlookSignalRCommandDispatcher>();
            services.AddSingleton<OutlookCommandQueue>();
            services.AddSingleton<OutlookFolderCacheService>();
            services.AddSingleton<OutlookFetchResultService>();
            services.AddSingleton<MockOutlookService>();
            services.AddHostedService<OutlookAddressBookBackgroundSyncService>();
            return services;
        }

        public static void AddOutlookSwaggerDocument(this SwaggerGenOptions options)
        {
            options.SwaggerDoc(SwaggerDocumentName, new OpenApiInfo
            {
                Title = "SmartOffice.Hub Outlook API",
                Version = "v1",
                Description = "Hub/Web UI/AI integration API for Outlook request routing and fetch-result data retrieval.",
            });
            options.DocInclusionPredicate((documentName, apiDescription) =>
                string.Equals(apiDescription.GroupName, documentName, StringComparison.OrdinalIgnoreCase));
            var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
            var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
            if (File.Exists(xmlPath))
                options.IncludeXmlComments(xmlPath);
            options.CustomSchemaIds(type => type.Name switch
            {
                "AddinStatusDto" => "OutlookWorkerStatusDto",
                "AddinLogEntry" => "OutlookWorkerLogEntry",
                _ => type.Name,
            });
            options.OperationFilter<OutlookSwaggerOperationFilter>();
        }

        public static IEndpointRouteBuilder MapOutlookFeatureHubs(this IEndpointRouteBuilder endpoints)
        {
            endpoints.MapHub<NotificationHub>(NotificationHubRoute);
            endpoints.MapHub<OutlookAddinHub>(OutlookAddinHubRoute);
            return endpoints;
        }

        public static void SeedOutlookMock(this IServiceProvider services)
        {
            services.GetRequiredService<MockOutlookService>().Seed();
        }
    }
}
