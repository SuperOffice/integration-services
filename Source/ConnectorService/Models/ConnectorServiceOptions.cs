namespace ConnectorService.Models
{
    public class ConnectorServiceOptions
    {
        public const string ConnectorService = "ConnectorService";
        public string ClientId { get; set; }
        public string PrivateKeyFile { get; set; }
        public string[] ConnectorAssemblies { get; set; }
    }
}
