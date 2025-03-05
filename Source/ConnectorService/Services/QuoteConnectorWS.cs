

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
    public class QuoteConnectorWS : OnlineQuoteConnector<ExcelQuoteConnector>, IIntegrationServiceConnectorAuth
    {
        public const string Endpoint = "QuoteConnectorWS.svc";
        private readonly SuperIdOptions _superIdOptions;
        private readonly ApplicationOptions _applicationOptions;
        private readonly ISuperOfficeTokenValidator _superOfficeTokenValidator;
        private readonly PartnerTokenIssuer _partnerTokenIssuer;

        public QuoteConnectorWS(
            IOptions<SuperIdOptions> superIdOptions,
            IOptions<ApplicationOptions> applicationOptions,
            ISuperOfficeTokenValidator superOfficeTokenValidator = null,
            IPartnerTokenIssuer partnerTokenIssuer = null
            ) : base
            (
                applicationOptions.Value.ClientId,
                GetPrivateKey(applicationOptions.Value.PrivateKeyFile)
            )
            {
                _superIdOptions = superIdOptions.Value;
                _applicationOptions = applicationOptions.Value;
                _superOfficeTokenValidator = superOfficeTokenValidator
                ?? new SuperOfficeTokenValidator(new LocalStoreSuperIdCertificateResolver(thumbbprint: _superIdOptions.Certificate)); // This certificate should be loaded from online's discovery document.
                _partnerTokenIssuer = new PartnerTokenIssuer(new PartnerCertificateResolver(() => PrivateKey));
            }

        /// <summary>
        /// Authenticates an integration service request by validating the provided signed token and ensuring it matches the expected audience.
        /// Returns an authentication response indicating success or failure.
        /// </summary>
        /// <param name="request">The authentication request containing the signed token.</param>
        /// <returns>An AuthenticationResponse indicating the result of the authentication process.</returns>
        AuthenticationResponse IIntegrationServiceConnectorAuth.Authenticate(AuthenticationRequest request)
        {
            var applicationIdentifier = _applicationOptions.ClientId;

            try
            {

                var token = ValidateSuperOfficeSignedToken(request.SignedToken);

                var audience = token.FindFirst("aud")?.Value;
                if (!string.Equals("spn:" + applicationIdentifier, audience, StringComparison.InvariantCultureIgnoreCase))
                {
                    return new AuthenticationResponse
                    {
                        Succeeded = false,
                        Reason = "Wrong audience, missmatch on application identifier"
                    };
                }

                // Get nonce and sign the partner token
                var nonce = token.GetNonce();
                if (string.IsNullOrEmpty(nonce))
                {
                    return new AuthenticationResponse
                    {
                        Succeeded = false,
                        Reason = "Failed to retrieve nonce from the token"
                    };
                }

                return new AuthenticationResponse
                {
                    Succeeded = true,
                    SignedApplicationToken = _partnerTokenIssuer.SignPartnerToken(token: nonce)
                };
            }
            catch
            {
                return new AuthenticationResponse
                {
                    Succeeded = false,
                    Reason = "Failed to validate authentication request"
                };
            }
        }

        /// <summary>
        /// Retrieves the inner typed quote connector based on the provided request.
        /// If the request originates from Online, it updates the connection configuration fields.
        /// </summary>
        /// <typeparam name="TRequest">The type of the request.</typeparam>
        /// <param name="request">The request containing connection configuration fields.</param>
        /// <returns>The inner typed quote connector.</returns>
        protected override ExcelQuoteConnector GetInnerTypedQuoteConnector<TRequest>(TRequest request)
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
        private static ConnectionConfigFields RefactorConnectionConfigFields(ConnectionConfigFields requestConfigFields)
        {
            // Create a new ConnectionConfigFields object to hold the updated values
            var updatedConnectionConfigFields = new ConnectionConfigFields();

            // Try to retrieve the file name from the connection config fields
            if (requestConfigFields.TryGetValue("#1", out var fileName))
            {
                updatedConnectionConfigFields.Add("#1", fileName);
            }
            else
            {
                // Set default fileName, as required by the ExcelQuoteConnector
                updatedConnectionConfigFields.Add("DefaultFileName", "ExcelConnectorWithCapabilities.xlsx");
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
