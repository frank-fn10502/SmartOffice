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
            builder.Services.AddSignalR();

            // The Hub is currently process-local: add-ins push Office data here and
            // Web UI / future MCP clients read the latest cached snapshot.
            builder.Services.AddSingleton<MailStore>();
            builder.Services.AddSingleton<ChatStore>();
            builder.Services.AddSingleton<CommandQueue>();
            builder.Services.AddSingleton<AddinStatusStore>();

            builder.Services.AddCors(options =>
            {
                // Prototype/local-network default. Tighten this before exposing the
                // Hub beyond a trusted workstation or intranet segment.
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
