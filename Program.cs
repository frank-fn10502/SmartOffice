using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Services;
using SmartOffice.Hub.Swagger;
using Microsoft.OpenApi.Models;
using System.Reflection;
using System.Text.Json.Serialization;

namespace SmartOffice.Hub
{
    public class Program
    {
        private const string OutlookSwaggerDocumentName = "outlook-v1";

        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddControllers()
                .AddJsonOptions(options =>
                {
                    options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
                });
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen(options =>
            {
                options.SwaggerDoc(OutlookSwaggerDocumentName, new OpenApiInfo
                {
                    Title = "SmartOffice.Hub Outlook API",
                    Version = "v1",
                    Description = "Hub/Web UI/AI integration API for Outlook command routing and cached snapshots.",
                });
                options.DocInclusionPredicate((documentName, apiDescription) =>
                    string.Equals(apiDescription.GroupName, documentName, StringComparison.OrdinalIgnoreCase));
                var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
                var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
                if (File.Exists(xmlPath))
                    options.IncludeXmlComments(xmlPath);
                options.OperationFilter<OutlookSwaggerOperationFilter>();
            });
            // SignalR 預設 incoming message 上限較小；folder tree 必須使用
            // BeginFolderSync / PushFolderBatch 小批次回推。
            builder.Services.AddSignalR(options =>
            {
                options.MaximumReceiveMessageSize = 256 * 1024;
            })
            .AddJsonProtocol(options =>
            {
                options.PayloadSerializerOptions.Converters.Add(new JsonStringEnumConverter());
            });

            // Hub 目前是 process-local：AddIn 將 Office data 透過 SignalR push 到這裡，
            // Web UI 與未來 MCP client 讀取最新 cached snapshot。
            builder.Services.AddSingleton<MailStore>();
            builder.Services.AddSingleton<ChatStore>();
            builder.Services.AddSingleton<CommandResultStore>();
            builder.Services.AddSingleton<AddinStatusStore>();
            builder.Services.AddSingleton<AttachmentExportService>();
            builder.Services.AddSingleton<OutlookSignalRCommandDispatcher>();
            builder.Services.AddSingleton<OutlookCommandQueue>();
            builder.Services.AddSingleton<OutlookFolderCacheService>();
            builder.Services.AddSingleton<MockOutlookService>();

            builder.Services.AddCors(options =>
            {
                // Prototype / local-network 預設值。若 Hub 要離開可信任 workstation
                // 或 intranet segment，必須先收緊這段 CORS policy。
                options.AddDefaultPolicy(policy =>
                    policy.SetIsOriginAllowed(_ => true).AllowAnyMethod().AllowAnyHeader().AllowCredentials());
            });

            var app = builder.Build();

            app.UseSwagger();
            app.UseSwaggerUI(options =>
            {
                options.DocumentTitle = "SmartOffice.Hub API";
                options.SwaggerEndpoint($"/swagger/{OutlookSwaggerDocumentName}/swagger.json", "Outlook API v1");
            });

            app.UseCors();
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseAuthorization();
            app.MapControllers();
            app.MapHub<NotificationHub>("/hub/notifications");
            app.MapHub<OutlookAddinHub>("/hub/outlook-addin");

            app.Services.GetRequiredService<MockOutlookService>().Seed();

            app.Run();
        }
    }
}
