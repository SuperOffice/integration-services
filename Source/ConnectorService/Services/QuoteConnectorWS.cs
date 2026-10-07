

using ConnectorService.Authentication;
using ConnectorService.Models;
using Microsoft.Extensions.Options;
using SuperOffice.Connectors;
using SuperOffice.Online.IntegrationService;
using SuperOffice.Online.IntegrationService.Contract.V1;
using SuperOffice.Online.Tokens;
using SuperOffice.SuperID.Contracts;
using SuperOffice.SuperID.Contracts.V1;

namespace ConnectorService.Services
{
    public class QuoteConnectorWS : OnlineQuoteConnector<QuoteConnector>, IIntegrationServiceConnectorAuth
    {
        public const string Endpoint = "QuoteConnectorWS.svc";
        private readonly ConnectorServiceOptions _connectorServiceOptions;
        private readonly ISuperOfficeTokenValidator _superOfficeTokenValidator;
        private readonly IPartnerTokenIssuer _partnerTokenIssuer;
        private readonly ILogger<QuoteConnectorWS> _logger;

        public QuoteConnectorWS(
            IOptions<ConnectorServiceOptions> connectorServiceOptions,
            ILogger<QuoteConnectorWS> logger,
            ISuperOfficeTokenValidator superOfficeTokenValidator,
            IPartnerTokenIssuer partnerTokenIssuer
        ) : base(connectorServiceOptions.Value.ClientId, connectorServiceOptions.Value.PrivateKeyFile)
        {
            _logger = logger;
            _connectorServiceOptions = connectorServiceOptions.Value;
            _superOfficeTokenValidator = superOfficeTokenValidator;
            _partnerTokenIssuer = partnerTokenIssuer;
        }

        /// <summary>
        /// Authenticates an integration service request by validating the provided signed token and ensuring it matches the expected audience.
        /// Returns an authentication response indicating success or failure.
        /// </summary>
        /// <param name="request">The authentication request containing the signed token.</param>
        /// <returns>An AuthenticationResponse indicating the result of the authentication process.</returns>
        AuthenticationResponse IIntegrationServiceConnectorAuth.Authenticate(AuthenticationRequest request)
            => IntegrationServiceAuthentication.Authenticate(request, _superOfficeTokenValidator, _connectorServiceOptions.ClientId, _partnerTokenIssuer, _logger);

        /// <summary>
        /// Retrieves the inner typed quote connector based on the provided request.
        /// If the request originates from Online, it updates the connection configuration fields.
        /// </summary>
        /// <typeparam name="TRequest">The type of the request.</typeparam>
        /// <param name="request">The request containing connection configuration fields.</param>
        /// <returns>The inner typed quote connector.</returns>
        protected override QuoteConnector GetInnerTypedQuoteConnector<TRequest>(TRequest request)
        {
            // Check if the request comes from Online by inspecting the first property of ConnectionConfigFields
            if (request.ConnectionConfigFields.Keys.FirstOrDefault() == "ApplicationId")
            {
                // Update the original ConnectionConfigFields with the new values
                request.ConnectionConfigFields = RefactorConnectionConfigFields(request.ConnectionConfigFields);
            }

            var inner = base.GetInnerTypedQuoteConnector(request);
            return inner;
        }

        /// <summary>
        /// Refactors the connection configuration fields by updating or adding specific fields required by the ExcelQuoteConnector.
        /// </summary>
        /// <param name="requestConfigFields">The original connection configuration fields.</param>
        /// <returns>A new ConnectionConfigFields object with the updated values.</returns>
        private ConnectionConfigFields RefactorConnectionConfigFields(ConnectionConfigFields requestConfigFields)
        {
            // Create a new ConnectionConfigFields object to hold the updated values
            var updatedConnectionConfigFields = new ConnectionConfigFields();

            // Try to retrieve the file name from the connection config fields
            if (requestConfigFields.TryGetValue("#1", out var fileName))
            {
                updatedConnectionConfigFields.Add("#1", Path.Combine(Path.Combine(AppContext.BaseDirectory, "Resources"), fileName));
            }
            else
            {
                updatedConnectionConfigFields.Add("DefaultFileName", Path.Combine(Path.Combine(AppContext.BaseDirectory, "Resources"), "ExcelConnectorWithCapabilities.xlsx"));
            }

            // Add the rest of the connection config fields
            foreach (var entry in requestConfigFields)
            {
                updatedConnectionConfigFields.TryAdd(entry.Key, entry.Value);
            }

            return updatedConnectionConfigFields;
        }
    }
}
