# Connectors

This project is created as a sample for how to set up a QuoteConnector and ERPConnector.

## Architecture

An overview of the architecture can be seen in [Architecture.dsl](./Architecture.dsl).

## ConnectorService
This project is the wrapper which exposes the services for the QuoteConnector.

Since .net Core does not natively support WCF, this projects makes use of the community-driven project [coreWCF](https://github.com/CoreWCF/CoreWCF) to add this support. CoreWCF has been picked up by microsoft and is now part of their [support policy](https://dotnet.microsoft.com/en-us/platform/support/policy/corewcf), but SuperOffice does not have anything to do with this support directly.

### Minimalistic API

The ConnectorService project contains a minimalistic API for handling reading from and writing to the ExcelFiles. It also enables the user to upload a new Excel-file to the service, or download the [template-excelfile](./Source/ConnectorService/Resources/ExcelConnectorWithCapabilities.xlsx). 

### EPPLUS

The service uses the [EPPLUS](https://www.epplussoftware.com/) for reading from and writing to excelfiles. This is a licensed product, and since this project is a sample (and not to be used in production), the license is set to be [NonCommercial](./Source/ConnectorService/Utils/ExcelHandler.cs#L29).

For production use, you need to acquire a license from EPPLUS (or make sure you adhere to their licensing terms).

### QuoteConnectorWS

This is the WCF service that is exposed to SuperOffice. It is a simple wrapper around the QuoteConnector, and is responsible for handling the incoming requests and returning the responses.

Note: 
The service requires a special snippet to check if the incoming request comes from SuperOffice Online. The snippet can be found in RefactorConnectionConfigFields() in [QuoteConnectorWS.cs](./Source/ConnectorService/Services/QuoteConnectorWS.cs).
This snippet is not mandatory for all connectors, but the [implementation](./Source/SuperOffice.ExcelQuoteConnector/ExcelQuoteConnector.cs#L255) for the connector expects the fileName to be in the first position of the `connectionConfigFields` dictionary. 
This filename is used to load data from Excel in [ReadInData()](./Source/SuperOffice.ExcelQuoteConnector/ExcelQuoteConnector.cs#L878)
This workaround is necessary to support both onsite and online versions of SuperOffice.

### ERPConnectorWS

## SuperOffice.ExcelQuoteConnector
Contains the implementation of the QuoteConnector, using a locally stored Excel file as the data source. **This implementation is not intended for production use, but as a sample for how to implement a QuoteConnector.**

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
