@secure()
param provisionParameters object
param provisionOutputs object

var webAppCurrentAppSettings = list('${provisionOutputs.frontendHostingOutput.value.webAppResourceId}/config/appsettings', '2021-02-01').properties

module teamsFxFrontendHostingConfig './teamsFx/frontendHosting.bicep' = {
  name: 'addTeamsFxFrontendHostingConfiguration'
  params:{
    provisionParameters: provisionParameters
    provisionOutputs: provisionOutputs
    currentAppSettings: webAppCurrentAppSettings
  }
}
