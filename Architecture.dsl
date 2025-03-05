workspace {
    
    !identifiers hierarchical
    
    model {
    
        softwareSystem = softwareSystem "SuperOffice Connectors" {

            online = group "Online" {
                onlineClient = container "Online Client" {
                    tags "Online" "Client"
                }
                developerPortal = container "Developer Portal" {
                    tags "Online" "Developer Portal"
                    onlineClient -> this "Uses client_id to get connector uri"
                }
            }

            services = group "Services" {
                connectorService = container "ConnectorService" {
                    tags "Connector" "Entrypoint" "Services"
                }
                quote = container "QuoteConnector.svc" {
                    tags "Services" "Quote"
                }
                erp = container "SyncConnector.svc" {
                    tags "Services" "ERP" "Sync"
                }

                connectorService -> quote "Uses"
                connectorService -> erp "Uses"

            }

            onsite = group "OnSite" {
                onsiteClient = container "OnSite Client" {
                    tags "Online" "Client"
                }
            }

            developerPortal -> connectorService
            onsiteClient -> connectorService
        }

    }
    
    views {
        container softwareSystem "Containers_All" {
            include *
            autolayout
        }

        container softwareSystem "Containers_online" {
            include ->softwareSystem.online->
            autolayout
        }

        container softwareSystem "Containers_onsite" {
            include ->softwareSystem.onsite->
            autolayout
        }

        container softwareSystem "Containers_services" {
            include ->softwareSystem.services->
            autolayout
        }

        styles {
            element "Services" {
                shape hexagon
                background #152192
            }

            element "Online" {
                background #91F0AE
            }
            element "OnSite" {
                background #EDF08C
            }
            
        }

    }

}
