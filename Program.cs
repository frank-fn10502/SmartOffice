using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddControllers();
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            // SignalR 預設 incoming message 上限較小；folder tree 必須使用
            // BeginFolderSync / PushFolderBatch 小批次回推。
            builder.Services.AddSignalR(options =>
            {
                options.MaximumReceiveMessageSize = 256 * 1024;
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
            app.UseSwaggerUI();

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
