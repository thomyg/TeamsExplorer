// Auto generated content, please customize files under provision folder

@secure()
param provisionParameters object
param provisionOutputs object
@secure()
param currentAppSettings object

var webAppName = split(provisionOutputs.frontendHostingOutput.value.webAppResourceId, '/')[8]

var m365ClientId = provisionParameters['m365ClientId']
var m365ClientSecret = provisionParameters['m365ClientSecret']
var m365TenantId = provisionParameters['m365TenantId']
var m365OauthAuthorityHost = provisionParameters['m365OauthAuthorityHost']
var oauthAuthority = uri(m365OauthAuthorityHost, m365TenantId)
var aadMetadataAddress = uri(m365OauthAuthorityHost, '${m365TenantId}/v2.0/.well-known/openid-configuration')
var tabAppDomain = provisionOutputs.frontendHostingOutput.value.domain
var tabAppEndpoint = provisionOutputs.frontendHostingOutput.value.endpoint 
var m365ApplicationIdUri = 'api://${tabAppDomain}/${m365ClientId}'
var initiateLoginEndpoint = uri(tabAppEndpoint, 'auth-start.html')

resource appSettings 'Microsoft.Web/sites/config@2021-02-01' = {
  name: '${webAppName}/appsettings'
  properties: union({
    AAD_METADATA_ADDRESS: aadMetadataAddress
    CLIENT_ID: m365ClientId
    CLIENT_SECRET: m365ClientSecret
    IDENTIFIER_URI: m365ApplicationIdUri
    IDENTITY_ID: provisionOutputs.identityOutput.value.identityClientId
    OAUTH_AUTHORITY: oauthAuthority
    TAB_APP_ENDPOINT: tabAppEndpoint
    TeamsFx__Authentication__ClientId: m365ClientId
    TeamsFx__Authentication__InitiateLoginEndpoint: initiateLoginEndpoint
    TeamsFx__Authentication__SimpleAuthEndpoint: tabAppEndpoint
  }, currentAppSettings)
}
