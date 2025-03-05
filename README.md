# Connectors

This project is created as a sample for how to set up a QuoteConnector and ERPConnector.

## Architecture

An overview of the architecture can be seen in [Architecture.dsl](./Architecture.dsl).

## ConnectorService
This project is the wrapper which exposes the services for the QuoteConnector.

Since .net Core does not natively support WCF, this projects makes use of the community-driven project [coreWCF](https://github.com/CoreWCF/CoreWCF) to add this support. CoreWCF has been picked up by microsoft and is now part of their [support policy](https://dotnet.microsoft.com/en-us/platform/support/policy/corewcf), but SuperOffice does not have anything to do with this support directly.

### QuoteConnectorWS

This is the WCF service that is exposed to SuperOffice. It is a simple wrapper around the QuoteConnector, and is responsible for handling the incoming requests and returning the responses.

It requires a special snippet to check if the incoming request comes from SuperOffice Online. The snippet can be found in RefactorConnectionConfigFields() in [QuoteConnectorWS.cs](./ConnectorService/Services/QuoteConnectorWS.cs).

### ERPConnectorWS

## SuperOffice.ExcelQuoteConnector
Contains the implementation of the QuoteConnector, using a locally stored Excel file as the data source.

## Quickstart

To set up a new QuoteConnector the following steps needs to be completed:

1. [Create a quote connector][0]
2. [Register an ERP and quote sync app][1]
3. [Configure the application with service endpoints][2]
4. [Create a new Quote Connector in SuperOffice Admin][3]

### Appsettings.json

The Service needs a clientId/Application identifier and the private certificate that belongs to the application. By default the certificate is located in "AppData/PrivateKey.xml", and the clientId can be found in appsettings.json.

<!-- Reference links -->
[0]: https://docs.superoffice.com/en/api/netserver/plugins/quote-connectors/online-quote-connectors/index.html
[1]: https://docs.superoffice.com/en/developer-portal/create-app/sync-app.html
[2]: https://docs.superoffice.com/en/developer-portal/create-app/config/update-endpoints.html
[3]: https://docs.superoffice.com/en/quote/learn/admin/erp-connection-add.html
[4]: https://docs.superoffice.com/en/api/netserver/plugins/quote-connectors/set-up.html#pluginresponseinfo-testconnection--dictionarystring-string-connectiondata-connectiondata-
