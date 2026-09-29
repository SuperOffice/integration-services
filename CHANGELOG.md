# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

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
