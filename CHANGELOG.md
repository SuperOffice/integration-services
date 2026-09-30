# Changelog

All notable changes to this project are documented in this file, grouped by pull request (newest first).

The categories follow [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [#3](https://github.com/SuperOffice/integration-services/pull/3) Fix startup, Docker build and stale configuration (unreleased)

### Added

- GitHub Actions CI workflow that runs on pull requests and pushes to `main`. It builds `Connectors.slnx`, builds the Docker image, starts the container and checks that `/`, Swagger and both WSDL endpoints respond.

### Fixed

- The service no longer fails at startup with "API Explorer not registered in DI". The OpenAPI setup now lives only in `OpenApiExtensions.AddOpenApi()`.
- Azure Key Vault is only added when `VaultUri` is set, so the service runs locally without a vault.
- Fixed the Dockerfile, which referenced project folders that no longer exist. `Swagger/custom.js` is now included in the published output.
- `ConnectorAssemblies` now points to `ErpConnector.dll`, and the default `Application:Host` is `localhost`.
- Failed `Authenticate` calls in `ErpConnectorWS` and `QuoteConnectorWS` are now logged.
- Requests to `/Services/*` that CoreWCF doesn't handle (for example http instead of https) now return 404 instead of a misleading `401 Invalid API Key.`
- Updated the README configuration section to match the shared `ConnectorService` settings, Key Vault and user secrets.

## [#1](https://github.com/SuperOffice/integration-services/pull/1) Update NuGet packages, migrate to .slnx and add package source mapping (2026-09-30)

### Changed

- Migrated the solution from `Connectors.sln` to `Connectors.slnx`. Requires Visual Studio 17.14+ or .NET SDK 9.0.200+.
- Added `nuget.config` that maps all packages to nuget.org (package source mapping).
- Updated CoreWCF.Http and CoreWCF.Primitives from 1.6.0 to 1.9.1.
- Updated Microsoft.IdentityModel.Tokens and System.IdentityModel.Tokens.Jwt from 6.36.0 to 8.19.1 (required by CoreWCF 1.9.1).
- Updated System.ServiceModel.Primitives from 4.10.3 to 8.1.2.
- Updated SuperOffice.Crm.Online.IntegrationServices from 10.5.3.698 to 10.5.5.982.
- Updated NSwag.AspNetCore from 14.2.0 to 14.7.1.
- Updated Microsoft.VisualStudio.Azure.Containers.Tools.Targets from 1.21.0 to 1.23.0.
- Updated Microsoft.Extensions.DependencyInjection.Abstractions, System.Diagnostics.DiagnosticSource and System.Diagnostics.EventLog to their latest 8.0.x patches.
- Enabled central transitive package pinning, and pinned Microsoft.IdentityModel.Protocols.OpenIdConnect to 8.19.1 so all IdentityModel packages resolve to the same version.

### Security

- Fixed 7 CoreWCF.Primitives advisories, including a critical SAML token authentication bypass (GHSA-xjr9-gg9q-jx3v).
- Pinned System.Security.Cryptography.Xml to 8.0.4 (GHSA-g8r8-53c2-pm3f).
