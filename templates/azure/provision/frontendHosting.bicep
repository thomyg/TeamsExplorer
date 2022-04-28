@secure()
param provisionParameters object
param userAssignedIdentityId string
var resourceBaseName = provisionParameters.resourceBaseName
var sku = contains(provisionParameters, 'frontendHostingSku') ? provisionParameters['frontendHostingSku'] : 'F1'
var serverFarmsName = contains(provisionParameters, 'frontendHostingServerFarmsName') ? provisionParameters['frontendHostingServerFarmsName'] : '${resourceBaseName}tab'
var webAppName = contains(provisionParameters, 'frontendHostingWebAppName') ? provisionParameters['frontendHostingWebAppName'] : '${resourceBaseName}tab'

resource serverFarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: serverFarmsName
  location: resourceGroup().location
  sku: {
    name: sku
  }
  kind: 'app'
}

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  name: webAppName
  location: resourceGroup().location
  properties: {
    serverFarmId: serverFarms.id
    keyVaultReferenceIdentity: userAssignedIdentityId
    siteConfig: {
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
      ]
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${userAssignedIdentityId}': {}
    }
  }
}

var siteDomain = webApp.properties.defaultHostName

output resourceId string = webApp.id
output endpoint string = 'https://${siteDomain}'
output domain string = siteDomain
