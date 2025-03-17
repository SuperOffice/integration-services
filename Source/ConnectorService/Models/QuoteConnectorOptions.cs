namespace ConnectorService.Models
{
    public class QuoteConnectorOptions
    {
        public const string QuoteConnector = "QuoteConnector";
        public string ClientId { get; set; }
        public string PrivateKeyFile { get; set; }
    }
}
