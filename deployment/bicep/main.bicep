@description('The base name for resources')
param baseName string = 'word-mcp'

@description('The environment (dev, test, prod)')
param environment string = 'dev'

@description('The location for all resources')
param location string = resourceGroup().location

@description('The Azure region code')
param regionCode string = 'weu'

@description('The name of the resource group')
param resourceGroupName string = 'rg-${baseName}-${environment}-${regionCode}'

@description('The MCP Server Auth Key')
@secure()
param mcpServerAuthKey string

@description('Azure Tenant ID for Graph API')
@secure()
param azureTenantId string

@description('Azure Client ID for Graph API')
@secure()
param azureClientId string

@description('Azure Client Secret for Graph API')
@secure()
param azureClientSecret string

@description('SharePoint Site URL')
param sharepointSiteUrl string

@description('SharePoint Template Folder Path')
param sharepointTemplateFolder string

@description('OneDrive User')
param onedriveUser string

@description('OneDrive Output Folder')
param onedriveOutputFolder string

// Define resource naming variables
var containerAppName = 'ca-${baseName}-${environment}-${regionCode}'
var containerAppEnvName = 'cae-${baseName}-${environment}-${regionCode}'
var logAnalyticsName = 'log-${baseName}-${environment}-${regionCode}'
var containerRegistryName = 'cr${replace(replace(baseName, \'-\', \'\'), \'_\', \'\')}${environment}${regionCode}'
var imageName = '${containerRegistryName}.azurecr.io/${baseName}:latest'

// Log analytics workspace
resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsName
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
    features: {
      enableLogAccessUsingOnlyResourcePermissions: true
    }
  }
}

// Container registry
resource containerRegistry 'Microsoft.ContainerRegistry/registries@2021-12-01-preview' = {
  name: containerRegistryName
  location: location
  sku: {
    name: 'Basic'
  }
  properties: {
    adminUserEnabled: true
  }
}

// Container app environment
resource containerAppEnvironment 'Microsoft.App/managedEnvironments@2022-10-01' = {
  name: containerAppEnvName
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalyticsWorkspace.properties.customerId
        sharedKey: logAnalyticsWorkspace.listKeys().primarySharedKey
      }
    }
  }
}

// Container app
resource containerApp 'Microsoft.App/containerApps@2022-10-01' = {
  name: containerAppName
  location: location
  properties: {
    managedEnvironmentId: containerAppEnvironment.id
    configuration: {
      activeRevisionsMode: 'Single'
      ingress: {
        external: true
        targetPort: 8080
        allowInsecure: false
        traffic: [
          {
            latestRevision: true
            weight: 100
          }
        ]
      }
      registries: [
        {
          server: '${containerRegistryName}.azurecr.io'
          username: containerRegistry.listCredentials().username
          passwordSecretRef: 'registry-password'
        }
      ]
      secrets: [
        {
          name: 'registry-password'
          value: containerRegistry.listCredentials().passwords[0].value
        }
        {
          name: 'mcp-server-auth-key'
          value: mcpServerAuthKey
        }
        {
          name: 'azure-tenant-id'
          value: azureTenantId
        }
        {
          name: 'azure-client-id'
          value: azureClientId
        }
        {
          name: 'azure-client-secret'
          value: azureClientSecret
        }
      ]
    }
    template: {
      containers: [
        {
          name: containerAppName
          image: imageName
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            {
              name: 'MCP_SERVER_AUTH_KEY'
              secretRef: 'mcp-server-auth-key'
            }
            {
              name: 'AZURE_TENANT_ID'
              secretRef: 'azure-tenant-id'
            }
            {
              name: 'AZURE_CLIENT_ID'
              secretRef: 'azure-client-id'
            }
            {
              name: 'AZURE_CLIENT_SECRET'
              secretRef: 'azure-client-secret'
            }
            {
              name: 'SHAREPOINT_SITE_URL'
              value: sharepointSiteUrl
            }
            {
              name: 'SHAREPOINT_TEMPLATE_FOLDER'
              value: sharepointTemplateFolder
            }
            {
              name: 'ONEDRIVE_USER'
              value: onedriveUser
            }
            {
              name: 'ONEDRIVE_OUTPUT_FOLDER'
              value: onedriveOutputFolder
            }
            {
              name: 'LOG_LEVEL'
              value: 'INFO'
            }
            {
              name: 'FILE_LOGGING'
              value: 'true'
            }
            {
              name: 'ENVIRONMENT'
              value: environment
            }
          ]
        }
      ]
      scale: {
        minReplicas: 1
        maxReplicas: 1
      }
    }
  }
}

// Outputs
output containerAppUrl string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output containerAppName string = containerAppName
output containerRegistryName string = containerRegistryName
output logAnalyticsName string = logAnalyticsName