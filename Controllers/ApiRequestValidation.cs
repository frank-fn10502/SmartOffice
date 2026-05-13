using Microsoft.AspNetCore.Mvc;

namespace SmartOffice.Hub.Controllers
{
    internal static class ApiRequestValidation
    {
        public static IActionResult? RequireFields(params (string Name, string? Value)[] fields)
        {
            var missing = fields
                .Where(field => string.IsNullOrWhiteSpace(field.Value))
                .Select(field => field.Name)
                .ToArray();
            return missing.Length == 0 ? null : MissingRequiredFields(missing);
        }

        public static IActionResult MissingRequiredFields(params string[] fields)
        {
            return new BadRequestObjectResult(new
            {
                request = "",
                status = "missing_required_fields",
                state = "failed",
                message = $"Missing required request field(s): {string.Join(", ", fields)}.",
                requiredFields = fields,
                data = new { },
            });
        }
    }
}
