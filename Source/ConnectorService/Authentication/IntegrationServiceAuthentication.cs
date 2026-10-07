using SuperOffice.Online.Tokens;
using SuperOffice.SuperID.Contracts.V1;

namespace ConnectorService.Authentication
{
    public static class IntegrationServiceAuthentication
    {
        public static AuthenticationResponse Authenticate(
            AuthenticationRequest request,
            ISuperOfficeTokenValidator tokenValidator,
            string clientId,
            IPartnerTokenIssuer partnerTokenIssuer,
            ILogger logger)
        {
            try
            {
                var token = tokenValidator.ValidateToken(request.SignedToken, "spn:" + clientId);

                var nonce = token.GetNonce();
                if (string.IsNullOrEmpty(nonce))
                {
                    return Failed("Failed to retrieve nonce from the token");
                }

                return new AuthenticationResponse
                {
                    Succeeded = true,
                    SignedApplicationToken = partnerTokenIssuer.SignPartnerToken(nonce)
                };
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Failed to validate authentication request");
                return Failed("Failed to validate authentication request");
            }
        }

        private static AuthenticationResponse Failed(string reason) => new() { Succeeded = false, Reason = reason };
    }
}
