namespace ConnectorService.Models
{
    public class ApplicationOptions
    {
        public const string Application = "Application";
        public string ClientId { get; set; }
        public string PrivateKeyFile { get; set; }

        public string BaseFilePath { get; set; }

        public string ExcelPath { get; set; }
    }
}
