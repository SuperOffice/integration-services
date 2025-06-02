namespace ConnectorService.Models
{
    public class ApplicationOptions
    {
        public const string Application = "Application";

        public bool EnableHttpsEndpoints { get; set; }

        public string Host { get; set; }
        public string ResourcesPath { get; set; }
    }
}
