namespace ConnectorService.Models
{
    public class ApplicationOptions
    {
        public const string Application = "Application";

        //public string Quote_ClientId { get; set; }
        //public string Quote_PrivateKeyFile { get; set; }

        //public string ERP_ClientId { get; set; }
        //public string ERP_PrivateKeyFile { get; set; }
        //public string BaseFilePath { get; set; }

        //public string ExcelPath { get; set; }

        public bool EnableHttpsEndpoints { get; set; }

        public string Host { get; set; }
        public string ResourcesPath { get; set; }
    }
}
