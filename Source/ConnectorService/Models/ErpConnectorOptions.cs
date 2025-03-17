namespace ConnectorService.Models
{
    public class ErpConnectorOptions
    {
        public const string ErpConnector = "ErpConnector";
        public string[] ConnectorAssemblies { get; set; }
        public string ClientId { get; set; }
        public string PrivateKeyFile { get; set; }
    }
}
