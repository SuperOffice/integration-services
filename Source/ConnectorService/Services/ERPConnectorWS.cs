using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ConnectorService.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Options;
using SuperOffice.CRM;
using SuperOffice.ErpSync;
using SuperOffice.ErpSync.ConnectorWS;
using SuperOffice.Factory;
using SuperOffice.Online.Tokens;
using SuperOffice.SuperID.Contracts;
using SuperOffice.SuperID.Contracts.V1;
using SuperOffice.Online.IntegrationService;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Reflection.Metadata;
using CoreWCF;

namespace ConnectorService.Services
{
    public class ErpConnectorWS : IErpConnectorWS, IIntegrationServiceConnectorAuth
    {
        public const string Endpoint = "ErpConnectorWS.svc";
        readonly HashSet<Assembly> _parsedAssemblies = new();
        private readonly ConnectorServiceOptions _connectorServiceOptions;
        private readonly SuperIdOptions _superIdOptions;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly ISuperOfficeTokenValidator _superOfficeTokenValidator;
        private readonly IPartnerTokenIssuer _partnerTokenIssuer;

        public ErpConnectorWS(
            IOptions<ConnectorServiceOptions> connectorServiceOptions,
            IOptions<ApplicationOptions> applicationOptions,
            IOptions<SuperIdOptions> superIdOptions,
            IWebHostEnvironment webHostEnvironment,
            ISuperOfficeTokenValidator superOfficeTokenValidator = null,
            IPartnerTokenIssuer partnerTokenIssuer = null
        )
        {
            _connectorServiceOptions = connectorServiceOptions.Value;
            _superIdOptions = superIdOptions.Value;
            _webHostEnvironment = webHostEnvironment;
            _superOfficeTokenValidator = superOfficeTokenValidator
                ?? new SuperOfficeTokenValidator(new LocalStoreSuperIdCertificateResolver(thumbbprint: _superIdOptions.Certificate)); // This certificate should be loaded from online's discovery document.
            _partnerTokenIssuer = partnerTokenIssuer ?? new PartnerTokenIssuer(new PartnerCertificateResolver(GetPrivateKey));
        }

        AuthenticationResponse IIntegrationServiceConnectorAuth.Authenticate(AuthenticationRequest request)
        {
            var applicationIdentifier = _connectorServiceOptions.ClientId;

            try
            {
                var token = ValidateSuperOfficeSignedToken(request.SignedToken);

                if (!string.Equals("spn:" + applicationIdentifier, token.FindFirst("aud").Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    return new AuthenticationResponse
                    {
                        Succeeded = false,
                        Reason = "Wrong audience, missmatch on application identifier"
                    };
                }

                return new AuthenticationResponse
                {
                    Succeeded = true,
                    SignedApplicationToken = _partnerTokenIssuer.SignPartnerToken(token.GetNonce())
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

        public string GetPrivateKey()
        {
            var fileName = _connectorServiceOptions.PrivateKeyFile;
            if (!Path.IsPathRooted(fileName))
                fileName = Path.Combine(_webHostEnvironment.ContentRootPath, path2: fileName);
            return File.ReadAllText(fileName);
        }

        public FieldMetadataInfoArrayPluginResponseWS WSGetConfigData()
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetConfigData();
                var result = ResponseHelper.CreateWSResponse<FieldMetadataInfoArrayPluginResponseWS>(impRes);
                result.FieldMetaDataObjects = impRes.FieldMetaDataObjects?.Select(r => r.FromPlugin()).ToArray();
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<FieldMetadataInfoArrayPluginResponseWS>(crash);
            }
        }

        private void ParseAssemblies()
        {
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (!AssemblyHelper.IsSystemAssembly(assembly) && !_parsedAssemblies.Contains(assembly))
                {
                    _parsedAssemblies.Add(assembly);
                    var types = assembly.GetTypes();

                    foreach (var type in types)
                    {
                        if (PluginInfo.IsValidPlugin(type, a => a.IsAutoDiscoverable == true))
                        {
                            PluginFactory.Add(PluginInfo.Create(type));
                        }
                    }
                }
            }
        }

        private IErpConnector GetImplementation()
        {
            // Parse assemblies and prime the NetServer class factory with IPlugin classes
            ParseAssemblies();
            var assemblyNames = _connectorServiceOptions.ConnectorAssemblies;
            // EisPluginLoader will parse assemblies and find plugins.

            var plugin = EisPluginLoader.Instance.GetConnector(OperationContext.Current.IncomingMessageHeaders.To, assemblyNames);

            // returns matching plugin (uses ?ConnectorName=xyz query string in URL), or throws NotFound exception
            return plugin;
        }

        public ConnectorResultBaseWS WSTestConfigData(TestConfigDataRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(implementation.TestConfigData(request.ConnectionInfo));
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(crash);
            }
        }

        public ConnectorResultBaseWS WSSaveConnection(SaveConnectionRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(implementation.SaveConnection(new Guid(request.ConnectionGuid), request.ConnectionInfo));
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(crash);
            }
        }

        public ConnectorResultBaseWS WSTestConnection(RequestBaseWS request)
        {
            try
            {
                var implementation = GetImplementation();
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(implementation.TestConnection(new Guid(request.ConnectionGuid)));
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(crash);
            }
        }

        public ConnectorResultBaseWS WSDeleteConnection(RequestBaseWS request)
        {
            try
            {
                var implementation = GetImplementation();
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(implementation.DeleteConnection(new Guid(request.ConnectionGuid)));
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ConnectorResultBaseWS>(crash);
            }
        }

        public StringArrayPluginResponseWS WSGetSupportedActorTypes(RequestBaseWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetSupportedActorTypes(new Guid(request.ConnectionGuid));
                var result = ResponseHelper.CreateWSResponse<StringArrayPluginResponseWS>(impRes);

                result.Items = impRes.Items;
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<StringArrayPluginResponseWS>(crash);
            }
        }

        public FieldMetadataInfoArrayPluginResponseWS WSGetSupportedActorTypeFields(GetSupportedActorTypeFieldsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetSupportedActorTypeFields(new Guid(request.ConnectionGuid), request.ActorType);
                var result = ResponseHelper.CreateWSResponse<FieldMetadataInfoArrayPluginResponseWS>(impRes);

                result.FieldMetaDataObjects = impRes.FieldMetaDataObjects == null ? null :
                    impRes.FieldMetaDataObjects.Select(r => r.FromPlugin()).ToArray();
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<FieldMetadataInfoArrayPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSGetActors(GetActorsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetActors(new Guid(request.ConnectionGuid), request.ActorType, request.ErpKeys, request.FieldKeys);
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        public StringArrayPluginResponseWS WSGetSearchableFields(GetSearchableFieldsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetSearchableFields(new Guid(request.ConnectionGuid), request.ActorType);
                var result = ResponseHelper.CreateWSResponse<StringArrayPluginResponseWS>(impRes);

                result.Items = impRes.Items;
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<StringArrayPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSSearchActorsAdvanced(SearchActorsAdvancedRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var restrictions = request.Restrictions.Select(r => r.ToPlugin()).ToArray();
                var impRes = implementation.SearchActorsAdvanced(new Guid(request.ConnectionGuid), request.ActorType, restrictions, request.FieldKeys);
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSSearchActors(SearchActorsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.SearchActors(new Guid(request.ConnectionGuid), request.ActorType, request.SearchText, request.FieldKeys);
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSSearchActorByParent(SearchActorByParentRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.SearchActorByParent(new Guid(request.ConnectionGuid), request.ActorType, request.SearchText,
                    request.ParentActorType, request.ParentActorErpKey, request.FieldKeys);
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        public ActorPluginResponseWS WSCreateActor(CreateActorRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.CreateActor(new Guid(request.ConnectionGuid), request.Actor.ToPlugin());
                var result = ResponseHelper.CreateWSResponse<ActorPluginResponseWS>(impRes);

                result.Actor = impRes.Actor.FromPlugin();
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSSaveActors(SaveActorsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.SaveActors(new Guid(request.ConnectionGuid), request.Actors.Select(a => a.ToPlugin()).ToArray());
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        public ListItemArrayPluginResponseWS WSGetList(GetListRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetList(new Guid(request.ConnectionGuid), request.ListName);
                var result = ResponseHelper.CreateWSResponse<ListItemArrayPluginResponseWS>(impRes);

                result.ListItems = impRes.ListItems;
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ListItemArrayPluginResponseWS>(crash);
            }
        }

        public ListItemArrayPluginResponseWS WSGetListItems(GetListItemsRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetListItems(new Guid(request.ConnectionGuid), request.ListName, request.ListItemKeys);
                var result = ResponseHelper.CreateWSResponse<ListItemArrayPluginResponseWS>(impRes);

                result.ListItems = impRes.ListItems;
                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ListItemArrayPluginResponseWS>(crash);
            }
        }

        public ActorArrayPluginResponseWS WSGetActorsByTimestamp(GetActorsByTimestampRequestWS request)
        {
            try
            {
                var implementation = GetImplementation();
                var impRes = implementation.GetActorsByTimestamp(new Guid(request.ConnectionGuid), request.UpdatedOnOrAfter, request.ActorType, request.FieldKeys);
                var result = ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(impRes);

                result.Actors = impRes.Actors == null ? null :
                    impRes.Actors.Select(a => a.FromPlugin()).ToArray();

                return result;
            }
            catch (Exception crash)
            {
                return ResponseHelper.CreateWSResponse<ActorArrayPluginResponseWS>(crash);
            }
        }

        private ClaimsIdentity ValidateSuperOfficeSignedToken(string token)
        {
            string ValidIssuer = "SuperOffice AS";

            var certificatePath = "App_Data/SuperOfficeFederatedLogin.crt";

            if (string.IsNullOrEmpty(certificatePath) || !File.Exists(certificatePath))
            {
                throw new FileNotFoundException($"Certificate file not found at {certificatePath}");
            }

            var tokenHandler = new JwtSecurityTokenHandler();
            var tokenValidationParameters = new TokenValidationParameters();
            tokenValidationParameters.ValidateAudience = false;
            tokenValidationParameters.ValidIssuer = ValidIssuer;
            tokenValidationParameters.IssuerSigningKey = new X509SecurityKey(new X509Certificate2(certificatePath));

            var principal = tokenHandler.ValidateToken(token, tokenValidationParameters, out var securityToken);
            return principal.Identities.OfType<ClaimsIdentity>().FirstOrDefault();
        }
    }
}
