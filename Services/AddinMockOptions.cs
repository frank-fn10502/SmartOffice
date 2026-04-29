namespace SmartOffice.Hub.Services
{
    public class AddinMockOptions
    {
        public bool Enabled { get; set; }
        public int ResponseDelayMilliseconds { get; set; } = 400;
        public OutlookAddinMockOptions Outlook { get; set; } = new();
    }

    public class OutlookAddinMockOptions
    {
        public bool Enabled { get; set; } = true;
    }
}
