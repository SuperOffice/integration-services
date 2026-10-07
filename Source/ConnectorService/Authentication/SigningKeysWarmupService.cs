namespace ConnectorService.Authentication
{
    // Fetches the signing keys at startup so the first Authenticate call does not pay for it.
    public class SigningKeysWarmupService : IHostedService
    {
        private readonly SuperOfficeJwksTokenValidator _tokenValidator;

        public SigningKeysWarmupService(SuperOfficeJwksTokenValidator tokenValidator)
        {
            _tokenValidator = tokenValidator;
        }

        public Task StartAsync(CancellationToken cancellationToken) => _tokenValidator.GetSigningKeysAsync();

        public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
    }
}
