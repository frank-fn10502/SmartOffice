namespace SmartOffice.Hub.Services
{
    internal static class UtcDateTime
    {
        public static DateTime Now => DateTime.UtcNow;

        public static DateTime Normalize(DateTime value)
        {
            return value.Kind switch
            {
                DateTimeKind.Utc => value,
                DateTimeKind.Local => value.ToUniversalTime(),
                _ => DateTime.SpecifyKind(value, DateTimeKind.Local).ToUniversalTime(),
            };
        }

        public static DateTime? Normalize(DateTime? value)
        {
            return value.HasValue ? Normalize(value.Value) : null;
        }
    }
}
