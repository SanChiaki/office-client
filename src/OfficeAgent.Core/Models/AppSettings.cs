namespace OfficeAgent.Core.Models
{
    public sealed class AppSettings
    {
        public string ApiKey { get; set; } = string.Empty;

        public string BaseUrl { get; set; } = "https://api.example.com";

        public string Model { get; set; } = "gpt-5-mini";
    }
}
