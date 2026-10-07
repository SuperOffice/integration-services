using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using SuperOffice.Online.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;

namespace ConnectorService.Authentication
{
    // Trusts the published keys of every Online environment, so the issuing environment needs no configuration.
    public class SuperOfficeJwksTokenValidator : ISuperOfficeTokenValidator
    {
        private const string ValidIssuer = "SuperOffice AS";
        private static readonly string[] Environments = ["sod", "qaonline", "online"];
        private static readonly TimeSpan FailureBackoff = TimeSpan.FromMinutes(1);
        private static readonly JwtSecurityTokenHandler TokenHandler = new();

        // A short timeout so an unreachable environment cannot stall Authenticate for the 100 s HttpClient default.
        private static readonly HttpDocumentRetriever DocumentRetriever = new(new HttpClient { Timeout = TimeSpan.FromSeconds(5) });

        private readonly ConfigurationManager<OpenIdConnectConfiguration>[] _configurationManagers = Environments
            .Select(environment => new ConfigurationManager<OpenIdConnectConfiguration>(
                $"https://{environment}.superoffice.com/login/.well-known/openid-configuration",
                new OpenIdConnectConfigurationRetriever(),
                DocumentRetriever))
            .ToArray();

        private readonly long[] _retryAfterTicks = new long[Environments.Length];
        private readonly Task<ICollection<SecurityKey>>[] _pendingFetches = new Task<ICollection<SecurityKey>>[Environments.Length];

        private readonly ILogger<SuperOfficeJwksTokenValidator> _logger;

        public SuperOfficeJwksTokenValidator(ILogger<SuperOfficeJwksTokenValidator> logger)
        {
            _logger = logger;
        }

        public ClaimsIdentity ValidateToken(string token, string audience)
        {
            try
            {
                return Validate(token, audience);
            }
            catch (SecurityTokenSignatureKeyNotFoundException)
            {
                // The signing key may have rotated; the refresh runs in the background, so a later request picks up the new key.
                foreach (var configurationManager in _configurationManagers)
                {
                    configurationManager.RequestRefresh();
                }
                throw;
            }
        }

        public async Task<List<SecurityKey>> GetSigningKeysAsync()
        {
            var keySets = await Task.WhenAll(_configurationManagers.Select((_, index) => GetEnvironmentSigningKeysAsync(index)));
            return keySets.SelectMany(keys => keys).ToList();
        }

        private ClaimsIdentity Validate(string token, string audience)
        {
            var tokenValidationParameters = new TokenValidationParameters
            {
                ValidIssuer = ValidIssuer,
                AudienceValidator = (audiences, _, _) => audiences.Contains(audience, StringComparer.OrdinalIgnoreCase),
                ValidAlgorithms = [SecurityAlgorithms.RsaSha256],
                // Authenticate is synchronous; this only blocks while an environment has never answered, later refreshes run in the background.
                IssuerSigningKeys = GetSigningKeysAsync().GetAwaiter().GetResult()
            };

            var principal = TokenHandler.ValidateToken(token, tokenValidationParameters, out _);
            return (ClaimsIdentity)principal.Identity;
        }

        private Task<ICollection<SecurityKey>> GetEnvironmentSigningKeysAsync(int index)
        {
            if (DateTime.UtcNow.Ticks < Interlocked.Read(ref _retryAfterTicks[index]))
            {
                return Task.FromResult<ICollection<SecurityKey>>([]);
            }

            // Callers arriving while a fetch is running share it instead of queuing for their own attempt.
            var pending = Volatile.Read(ref _pendingFetches[index]);
            if (pending is { IsCompleted: false })
            {
                return pending;
            }

            var fetch = FetchSigningKeysAsync(index);
            Volatile.Write(ref _pendingFetches[index], fetch);
            return fetch;
        }

        // One unreachable environment must not block tokens from the others.
        private async Task<ICollection<SecurityKey>> FetchSigningKeysAsync(int index)
        {
            var configurationManager = _configurationManagers[index];
            try
            {
                // Not cancellable: the fetch is shared between callers, so one caller's cancellation must not fail the others.
                return (await configurationManager.GetConfigurationAsync(CancellationToken.None).ConfigureAwait(false)).SigningKeys;
            }
            catch (Exception ex)
            {
                // Without this, every request would wait for the timeout while an environment has never been reachable.
                Interlocked.Exchange(ref _retryAfterTicks[index], DateTime.UtcNow.Add(FailureBackoff).Ticks);
                _logger.LogWarning(ex, "Failed to load SuperOffice signing keys from {MetadataAddress}", configurationManager.MetadataAddress);
                return [];
            }
        }
    }
}
