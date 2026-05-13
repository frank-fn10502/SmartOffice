using SmartOffice.Hub.Services;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json.Serialization;

namespace SmartOffice.Hub
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddControllers()
                .AddJsonOptions(options =>
                {
                    options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
                    options.JsonSerializerOptions.UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow;
                });
            builder.Services.Configure<ApiBehaviorOptions>(options =>
            {
                options.InvalidModelStateResponseFactory = context =>
                {
                    var requestName = context.HttpContext.Request.Path.Value?
                        .Split('/', StringSplitOptions.RemoveEmptyEntries)
                        .LastOrDefault() ?? string.Empty;
                    var allErrors = context.ModelState
                        .Where(item => item.Value?.Errors.Count > 0)
                        .ToDictionary(
                            item => item.Key,
                            item => item.Value!.Errors.Select(error =>
                                string.IsNullOrWhiteSpace(error.ErrorMessage)
                                    ? error.Exception?.Message ?? "Invalid value."
                                    : error.ErrorMessage).ToArray());
                    var errors = allErrors.Count <= 1
                        ? allErrors
                        : allErrors
                            .Where(item => item.Key is not "req" and not "msg" and not "entry")
                            .ToDictionary(item => item.Key, item => item.Value);

                    return new BadRequestObjectResult(new
                    {
                        request = requestName,
                        status = "invalid_request_body",
                        state = "failed",
                        message = "Request JSON does not match this endpoint schema. Remove unknown fields and use the exact property names documented in Swagger.",
                        errors,
                        data = new { },
                    });
                };
            });
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen(options =>
            {
                options.AddOutlookSwaggerDocument();
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

            // Hub 目前是 process-local：執行端將 Office data 透過 SignalR push 到這裡，
            // Web UI 與 API caller 透過 request-* / fetch-result-* 讀取結果。
            builder.Services.AddOutlookFeature();

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
                options.SwaggerEndpoint($"/swagger/{OutlookFeatureRegistration.SwaggerDocumentName}/swagger.json", "Outlook API v1");
            });

            app.UseCors();
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseAuthorization();
            app.MapControllers();
            app.MapOutlookFeatureHubs();

            app.Services.SeedOutlookMock();

            app.Run();
        }
    }
}
