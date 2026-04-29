using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Services;
using SmartOffice.Hub.Services.MockAddins;

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
            builder.Services.AddSignalR();

            // Hub 目前是 process-local：Add-in 將 Office data push 到這裡，
            // Web UI 與未來 MCP client 讀取最新 cached snapshot。
            builder.Services.AddSingleton<MailStore>();
            builder.Services.AddSingleton<ChatStore>();
            builder.Services.AddSingleton<CommandQueue>();
            builder.Services.AddSingleton<AddinStatusStore>();
            builder.Services.Configure<AddinMockOptions>(
                builder.Configuration.GetSection("AddinMocks"));
            if (builder.Configuration.GetValue<bool>("AddinMocks:Enabled")
                && builder.Configuration.GetValue<bool>("AddinMocks:Outlook:Enabled"))
                builder.Services.AddHostedService<OutlookMockAddinWorker>();

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

            app.Run();
        }
    }
}
