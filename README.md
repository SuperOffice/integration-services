# Connectors

This project is created as a sample for how to set up a QuoteConnector and ERPConnector.

## Architecture

An overview of the architecture can be seen in [Architecture.dsl](./Architecture.dsl).

## ConnectorService
This project is the wrapper which exposes the services for the QuoteConnector.

Since .net Core does not natively support WCF, this projects makes use of the community-driven project [coreWCF](https://github.com/CoreWCF/CoreWCF) to add this support. CoreWCF has been picked up by microsoft and is now part of their [support policy](https://dotnet.microsoft.com/en-us/platform/support/policy/corewcf), but SuperOffice does not have anything to do with this support directly.

### Minimalistic API

The ConnectorService project contains a minimalistic API for handling reading from and writing to the ExcelFiles. It also enables the user to upload a new Excel-file to the service, or download the [template-excelfile](./Source/ConnectorService/Resources/ExcelConnectorWithCapabilities.xlsx). 

All endpoints configured for this sample can be found in [ExcelHandlerEndpoints.cs](./Source/ConnectorService/API/ExcelHandlerEndpoints.cs).

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
2. [Register an ERPSyncConector app][1]
3. [Register an QuoteConnector app][1]
4. [Configure the application with service endpoints][2]
5. [Host the application](#Hosting)
5. [Create a new Quote Connector in SuperOffice Admin][3]

Note: The ERPSyncConnector and QuoteConnector apps need different application_id, so you cannot share 1 application for both connectors.

### Hosting

SuperOffice does not require the service to be hosted in a specific way, but it needs to be accesible externally. One way to host the service is to use [Azure App Services][5], but any hosting provider that supports .Net Core should work.

Alternatively it can be run directly from Visual Studio by using the launcSettings.json for https:
![Vs Hosting](Resources/vs-hosting.png)

#### Ngrok

For local development you can use [ngrok](https://ngrok.com/) to expose your local service to the internet. This is useful for testing the service with SuperOffice Online.

Through Ngrok.exe you can run this command to create a tunnel for a locally runnig SSL secure endpoint: 

`ngrok http https://localhost:7128 --url=YOURDOMAIN.ngrok-free.app`

Adjust port 7128 to be the port for the internally running application, and YOURDOMAIN to be what ngrok gives you when setting it up through [ngrok.com](https://ngrok.com/)


### Appsettings.json

```json
  "Application": {
    "EnableHttpsEndpoints": true,
    "Host": "HOSTNAME",
    "ResourcesPath": "Resources"
  },
  "QuoteConnector": {
    "ClientId": "73ed40c8...",
    "PrivateKeyFile": "App_Data/Quote_PrivateKey.xml"
  },
  "ErpConnector": {
    "ClientId": "ecf27a4...",
    "PrivateKeyFile": "App_Data/PrivateKey.xml",
    "ConnectorAssemblies": [
      "SuperOffice.EIS.TestConnector.dll"
    ]
  }
```

The Service needs a clientId/Application identifier and the private certificate that belongs to the application. By default the certificate is located in "AppData/PrivateKey.xml", and the clientId can be found in appsettings.json.

Note that the ERPConnectod and QuoteConnector has different credentials.

If the service is hosted in Azure, the `Application.Host` should be set to the hostname of the Azure App Service. All of these settings can be found in Azure portal, and settings defined in the portal directly will override the settings defined in appsettings.json. Please refer to [Microsoft documentation][6] for how to configure your application.

If you are running this application locally, on localhost, you can set the hostname to be `localhost`. 

The reasoning for setting `Application.Host` specifically can be seen in [this discussion on github](https://github.com/CoreWCF/CoreWCF/discussions/1515). For a project running locally/outside of an App Service this setting can probably be omitted, but it is included in this POC for completeness.

## Where are the data used by the Connectors?

The data provided by the connectors are all located in [Resources](./Resources). 

* EIS_Connections.txt - Used by SuperOffice.EIS.TestConnector to store information about a connection that has been created. 
* ErpClient.xslm - Used by the SuperOffice.EIS.TestConnector to provide data back to SuperOffice.
* ExcelConnectorWithCapabilities.xlsx - Used by the ExcelQuoteConnector to provide data back to SuperOffice.

Editing these files will reflect the data seen inside of SuperOffice, and in SuperOffice Admin it is neccessary to point to the correct file to get the data. It is also possible to upload your own files, through the [Minimalistic API](#Minimalistic_API), or download one of the existing 'templates' above and re-upload new versions.

In a real-world scenario, the data would be fetched from an external system, and not stored in the project itself!

<!-- Reference links -->
[0]: https://docs.superoffice.com/en/api/netserver/plugins/quote-connectors/online-quote-connectors/index.html
[1]: https://docs.superoffice.com/en/developer-portal/create-app/sync-app.html
[2]: https://docs.superoffice.com/en/developer-portal/create-app/config/update-endpoints.html
[3]: https://docs.superoffice.com/en/quote/learn/admin/erp-connection-add.html
[4]: https://docs.superoffice.com/en/api/netserver/plugins/quote-connectors/set-up.html#pluginresponseinfo-testconnection--dictionarystring-string-connectiondata-connectiondata-
[5]: https://azure.microsoft.com/en-us/products/app-service
[6]: https://learn.microsoft.com/en-us/azure/app-service/configure-common?tabs=portal
