using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class AttachmentExportService
    {
        private static readonly HashSet<string> BlockedExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".bat", ".cmd", ".com", ".exe", ".js", ".jse", ".msi", ".ps1", ".scr", ".vbs", ".wsf",
        };

        private readonly object _lock = new();
        private string _rootPath;

        public string RootPath
        {
            get { lock (_lock) return _rootPath; }
        }

        public string DefaultRootPath { get; }

        public AttachmentExportService(IWebHostEnvironment environment)
        {
            DefaultRootPath = ResolveDefaultRootPath();
            _rootPath = DefaultRootPath;
            Directory.CreateDirectory(_rootPath);
        }

        public AttachmentExportSettingsDto GetSettings()
        {
            return new AttachmentExportSettingsDto
            {
                RootPath = RootPath,
                DefaultRootPath = DefaultRootPath,
            };
        }

        public AttachmentExportSettingsDto UpdateSettings(string rootPath)
        {
            var nextRoot = string.IsNullOrWhiteSpace(rootPath) ? DefaultRootPath : Path.GetFullPath(rootPath);
            Directory.CreateDirectory(nextRoot);
            lock (_lock)
            {
                _rootPath = nextRoot;
            }
            return GetSettings();
        }

        public string CreateExportPath(string mailId, string subject, DateTime receivedTime, string fileName)
        {
            var mailFolder = $"{receivedTime:yyyyMMdd-HHmmss}_{SanitizeFileName(subject, "mail")}_{SanitizeFileName(mailId, "id")}";
            var directory = Path.Combine(RootPath, "Mails", mailFolder);
            Directory.CreateDirectory(directory);

            var safeName = SanitizeFileName(fileName, "attachment");
            var candidate = Path.Combine(directory, safeName);
            var name = Path.GetFileNameWithoutExtension(safeName);
            var extension = Path.GetExtension(safeName);
            var counter = 1;

            while (File.Exists(candidate))
            {
                candidate = Path.Combine(directory, $"{name}-{counter}{extension}");
                counter++;
            }

            return candidate;
        }

        public string CreateExportPath(string mailId, string fileName)
        {
            return CreateExportPath(mailId, string.Empty, DateTime.Now, fileName);
        }

        public bool IsAllowedPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            var fullPath = Path.GetFullPath(path);
            return fullPath.StartsWith(RootPath, StringComparison.OrdinalIgnoreCase);
        }

        public void OpenExportedFile(string path)
        {
            if (!IsAllowedPath(path))
                throw new InvalidOperationException("Attachment path is outside the configured export root.");

            if (!File.Exists(path))
                throw new FileNotFoundException("Exported attachment file was not found.", path);

            var extension = Path.GetExtension(path);
            if (BlockedExtensions.Contains(extension))
                throw new InvalidOperationException("This attachment type cannot be opened by the Hub host.");

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true,
            });
        }

        private static string SanitizeFileName(string value, string fallback)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var builder = new StringBuilder();
            foreach (var ch in value.Trim())
            {
                builder.Append(invalidChars.Contains(ch) ? '_' : ch);
            }

            var sanitized = builder.ToString().Trim(' ', '.');
            if (string.IsNullOrWhiteSpace(sanitized)) return fallback;
            return sanitized.Length <= 80 ? sanitized : sanitized[..80];
        }

        private static string ResolveDefaultRootPath()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var preferredDrive = new[] { @"E:\", @"D:\", @"C:\" }
                    .FirstOrDefault(Directory.Exists) ?? @"C:\";
                return Path.GetFullPath(Path.Combine(preferredDrive, "SmartOffice", "Attachments"));
            }

            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrWhiteSpace(home))
                home = Environment.GetEnvironmentVariable("HOME") ?? Path.GetTempPath();

            return Path.GetFullPath(Path.Combine(home, "SmartOffice", "Attachments"));
        }
    }
}
