@description('The location the resources should be created in. Leave as is.')
param location string = resourceGroup().location

@description('Optional tags to apply to the Azure resources during deployment.')
param tags object

@description('The name of the Azure Automation account.')
param automationAccount_name string

@description('The identity type that should be assigned to the Azure Automation Account. Possible values include: \'SystemAssigned\', \'UserAssigned\'.')
@allowed([
  'SystemAssigned'
  'UserAssigned'
])
param automationAccount_identity string = 'SystemAssigned'

@description('The Stock Keeping Unit (SKU) of the Azure Automation Account. Possible values include: \'Free\', \'Basic\', \'Standard\', \'Premium\'.')
@allowed([
  'Free'
  'Basic'
  'Standard'
  'Premium'
])
param automationAccount_sku string = 'Basic'

@description('The base time that is used for calculating the start time of the Automation runbook schedule.')
param baseTime string = utcNow('u')

// @description('The PowerShell modules URI.')
// #disable-next-line no-hardcoded-env-urls
// param psModulesUri string = 'https://devopsgallerystorage.blob.core.windows.net/packages/'

@description('The public URI from which Azure can GET the full script. For example: you can use a public GitHub repository or use ngrok.io to make the script available to the public during the deployment.')
param runbook_SyncSharePointUserProfilesFromAzureActiveDirectory_script_uri string

@description('The public URI from which Azure can GET the full script. For example: you can use a public GitHub repository or use ngrok.io to make the script available to the public during the deployment.')
param runbook_UpdateSharePointUserProfilesSyncJobStatus_script_uri string

@description('The name of the SharePoint connection.')
param connections_sharepointonline_name string = 'sharepointonline'

@description('The name of the Azure Automation Account connection.')
param connections_azureautomation_name string = 'azureautomation'

@description('The name of the Office 365 connection.')
param connections_office365_name string = 'office365'

@description('The client Id of the Azure AD application used to authenticate some connections.')
param credentials_clientId string

@description('The client secret of the Azure AD application used to authenticate some connections.')
@secure()
param credentials_clientSecret string

@description('The name of the Update SharePoint User Profiles Sync Job Status runbook.')
param runbook_UpdateSharePointUserProfilesSyncJobStatus_name string

@description('The name of the Sync SharePoint User Profiles From Azure Active Directory runbook.')
param runbook_SyncSharePointUserProfilesFromAzureActiveDirectory_name string

@description('The name of the Logic App.')
param logicApp_name string

@description('The identity type that should be assigned to the Azure Automation Account. Possible values include: \'SystemAssigned\', \'UserAssigned\'.')
@allowed([
  'SystemAssigned'
  'UserAssigned'
])
param logicApp_identity string = 'SystemAssigned'

@description('The default provisioned state of the Logic App workflow. Possible values include: \'Enabled\', \'Disabled\'.')
@allowed([
  'Enabled'
  'Disabled'
])
param logicApp_state string = 'Disabled'

@description('The comma-separated list of observers that should be notified of the flow run results.')
param logicApp_paramaters_processObservers string

@description('The URL of the SharePoint database site.')
param logicApp_sharePointDatabaseSiteUrl string

@description('The Id of the SharePoint database Jobs list.')
param logicApp_sharePointDatabaseJobsListId string

var connections = [
  {
    type: 'Microsoft.Web/connections'
    apiVersion: '2016-06-01'
    name: connections_office365_name
    location: location
    properties: {
      api: {
        name: connections_office365_name
        displayName: 'Office 365 Outlook'
        description: 'Microsoft Office 365 is a cloud-based service that is designed to help meet your organization\'s needs for robust security, reliability, and user productivity.'
        iconUri: 'https://connectoricons-prod.azureedge.net/releases/v1.0.1507/1.0.1507.2528/${connections_office365_name}/icon.png'
        brandColor: '#0078D4'
        id: '/subscriptions/${subscription().subscriptionId}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_office365_name}'
        type: 'Microsoft.Web/locations/managedApis'
      }
    }
  }
  {
    name: connections_sharepointonline_name
    properties: {
      nonSecretParameterValues: {
        'token:clientId': credentials_clientId
        'token:clientSecret': credentials_clientSecret
        'token:TenantId': subscription().tenantId
        'token:grantType': 'client_credentials'
      }
      api: {
        name: connections_sharepointonline_name
        displayName: 'SharePoint'
        description: 'SharePoint helps organizations share and collaborate with colleagues, partners, and customers. You can connect to SharePoint Online or to an on-premises SharePoint 2013 or 2016 farm using the On-Premises Data Gateway to manage documents and list items.'
        iconUri: 'https://connectoricons-prod.azureedge.net/releases/v1.0.1507/1.0.1507.2528/${connections_sharepointonline_name}/icon.png'
        brandColor: '#036C70'
        id: '/subscriptions/${subscription().subscriptionId}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_sharepointonline_name}'
        type: 'Microsoft.Web/locations/managedApis'
      }
    }
  }
  {
    name: connections_azureautomation_name
    properties: {
      nonSecretParameterValues: {
        'token:clientId': credentials_clientId
        'token:clientSecret': credentials_clientSecret
        'token:TenantId': subscription().tenantId
        'token:grantType': 'client_credentials'
      }
      api: {
        name: connections_azureautomation_name
        displayName: 'Azure Automation'
        description: 'Azure Automation provides tools to manage your cloud and on-premises infrastructure seamlessly.'
        iconUri: 'https://connectoricons-prod.azureedge.net/releases/v1.0.1538/1.0.1538.2619/${connections_azureautomation_name}/icon.png'
        brandColor: '#56A0D7'
        id: '/subscriptions/${subscription().subscriptionId}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_azureautomation_name}'
        type: 'Microsoft.Web/locations/managedApis'
      }
    }
  }
]
var runbooks = [
  {
    name: runbook_SyncSharePointUserProfilesFromAzureActiveDirectory_name
    properties: {
      runbookType: 'PowerShell7'
      logProgress: false
      logVerbose: true
      description: 'Script to synchronize SharePoint User Profiles from Azure Active Directory.'
      publishContentLink: {
        uri: runbook_SyncSharePointUserProfilesFromAzureActiveDirectory_script_uri
        version: '1.0.0.0'
      }
    }
  }
  {
    name: runbook_UpdateSharePointUserProfilesSyncJobStatus_name
    properties: {
      runbookType: 'PowerShell7'
      logProgress: false
      logVerbose: true
      description: 'Script to synchronize SharePoint User Profiles from Azure Active Directory.'
      publishContentLink: {
        uri: runbook_UpdateSharePointUserProfilesSyncJobStatus_script_uri
        version: '1.0.0.0'
      }
    }
  }
]
// var psModules = [
//   {
//     name: 'PnP.PowerShell'
//     url: uri(psModulesUri, 'pnp.powershell.1.12.0.nupkg')
//     version: '7.1'
//   }
// ]
var logicApp = {
  state: logicApp_state
  parameters: {
    '$connections': {
      defaultValue: {
      }
      type: 'Object'
    }
    parObservers: {
      defaultValue: logicApp_paramaters_processObservers
      type: 'String'
    }
  }
  triggers: {
    'When_an_item_is_created_-_Sync_job': {
      conditions: [
        {
          expression: '@or(or(equals(triggerBody()?[\'State\']?[\'Value\'], \'Submitted\'), equals(triggerBody()?[\'State\']?[\'Value\'], \'Error\')), equals(triggerBody()?[\'State\']?[\'Value\'], \'Skipped\'))'
        }
      ]
      inputs: {
        host: {
          connection: {
            name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
          }
        }
        method: 'get'
        path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${logicApp_sharePointDatabaseSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${logicApp_sharePointDatabaseJobsListId}\'))}/onnewitems'
      }
      recurrence: {
        frequency: 'Minute'
        interval: 5
      }
      splitOn: '@triggerBody()?[\'value\']'
      type: 'ApiConnection'
    }
  }
  actions: {
    'Initialize_variable_-_varBreakLoop': {
      inputs: {
        variables: [
          {
            name: 'varBreakLoop'
            type: 'boolean'
            value: '@false'
          }
        ]
      }
      runAfter: {
      }
      type: 'InitializeVariable'
    }
    'Initialize_variable_-_varMinutesToDelay': {
      inputs: {
        variables: [
          {
            name: 'varMinutesToDelay'
            type: 'integer'
            value: 5
          }
        ]
      }
      runAfter: {
        'Initialize_variable_-_varBreakLoop': [
          'Succeeded'
        ]
      }
      type: 'InitializeVariable'
    }
    'Switch_-_Job_state': {
      cases: {
        'Case_-_Skipped': {
          actions: {
          }
          case: 'Skipped'
        }
        'Case_-_Submitted': {
          actions: {
            'Until_-_Job_has_finished': {
              actions: {
                'Condition_-_If_runbook_succeeded': {
                  actions: {
                    'Condition_-_If_job_succeeded': {
                      actions: {
                        'Set_variable_-_varBreakLoop_1': {
                          inputs: {
                            name: 'varBreakLoop'
                            value: '@true'
                          }
                          runAfter: {
                          }
                          type: 'SetVariable'
                        }
                      }
                      else: {
                        actions: {
                          'Condition_-_If_job_is_still_ongoing': {
                            actions: {
                              'Compose_-_Calculate_new_delay_time': {
                                inputs: '@mul(variables(\'varMinutesToDelay\'), 1.25)'
                                runAfter: {
                                }
                                type: 'Compose'
                              }
                              'Set_variable_-_varMinutesToDelay': {
                                inputs: {
                                  name: 'varMinutesToDelay'
                                  value: '@outputs(\'Compose_-_Calculate_new_delay_time\')'
                                }
                                runAfter: {
                                  'Compose_-_Calculate_new_delay_time': [
                                    'Succeeded'
                                  ]
                                }
                                type: 'SetVariable'
                              }
                            }
                            else: {
                              actions: {
                                'Send_an_email_(V2)_-_Notify_observers': {
                                  inputs: {
                                    body: {
                                      Body: '<p>Hi,<br>\n<br>\nThe latest sync job failed. Please review the information below and act accordingly:<br>\nSolution: aad2spup<br>\nLogic App: @{workflow()[\'name\']}<br>\nJob Id: @{triggerBody()?[\'JobId\']}<br>\nStatus:: @{triggerBody()?[\'State\']?[\'Value\']}<br>\nSP List Item: <a href=\'@{triggerBody()?[\'{Link}\']}\' target=\'_blank\'>Go to list item</a><br>\nRun: <a href=\'@{concat(\'https://portal.azure.com/#view/Microsoft_Azure_EMA/LogicAppsMonitorBlade/runid/\', encodeUriComponent(workflow()?[\'run\'][\'id\']))}\' target=\'_blank\'>Go to run</a><br>\n</p>'
                                      Importance: 'High'
                                      Subject: 'aad2spup: latest sync job failed'
                                      To: '@parameters(\'parObservers\')'
                                    }
                                    host: {
                                      connection: {
                                        name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
                                      }
                                    }
                                    method: 'post'
                                    path: '/v2/Mail'
                                  }
                                  runAfter: {
                                    'Set_variable_-_varBreakLoop_2': [
                                      'Succeeded'
                                    ]
                                  }
                                  type: 'ApiConnection'
                                }
                                'Set_variable_-_varBreakLoop_2': {
                                  inputs: {
                                    name: 'varBreakLoop'
                                    value: '@true'
                                  }
                                  runAfter: {
                                  }
                                  type: 'SetVariable'
                                }
                              }
                            }
                            expression: {
                              or: [
                                {
                                  equals: [
                                    '@body(\'Parse_JSON_-_Create_dynamic_object_of_job_output\')?[\'State\']'
                                    'Submitted'
                                  ]
                                }
                                {
                                  equals: [
                                    '@body(\'Parse_JSON_-_Create_dynamic_object_of_job_output\')?[\'State\']'
                                    'Processing'
                                  ]
                                }
                                {
                                  equals: [
                                    '@body(\'Parse_JSON_-_Create_dynamic_object_of_job_output\')?[\'State\']'
                                    'Queued'
                                  ]
                                }
                              ]
                            }
                            runAfter: {
                            }
                            type: 'If'
                          }
                        }
                      }
                      expression: {
                        and: [
                          {
                            equals: [
                              '@body(\'Parse_JSON_-_Create_dynamic_object_of_job_output\')?[\'State\']'
                              'Succeeded'
                            ]
                          }
                        ]
                      }
                      runAfter: {
                        'Parse_JSON_-_Create_dynamic_object_of_job_output': [
                          'Succeeded'
                        ]
                      }
                      type: 'If'
                    }
                    Get_job_output: {
                      inputs: {
                        host: {
                          connection: {
                            name: '@parameters(\'$connections\')[\'azureautomation\'][\'connectionId\']'
                          }
                        }
                        method: 'get'
                        path: '/subscriptions/@{encodeURIComponent(\'${subscription().subscriptionId}\')}/resourceGroups/@{encodeURIComponent(\'${resourceGroup().name}\')}/providers/Microsoft.Automation/automationAccounts/@{encodeURIComponent(\'${automationAccount_name}\')}/jobs/@{encodeURIComponent(body(\'Create_job_-_run-UpdateSharePointUserProfilesSyncJobStatus\')?[\'properties\']?[\'jobId\'])}/output'
                        queries: {
                          'x-ms-api-version': '2015-10-31'
                        }
                      }
                      runAfter: {
                      }
                      type: 'ApiConnection'
                    }
                    'Parse_JSON_-_Create_dynamic_object_of_job_output': {
                      inputs: {
                        content: '@body(\'Get_job_output\')'
                        schema: {
                          properties: {
                            Error: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                            ErrorMessage: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                            JobId: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                            LogFolderUri: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                            SourceUri: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                            State: {
                              type: [
                                'null'
                                'string'
                              ]
                            }
                          }
                          type: 'object'
                        }
                      }
                      runAfter: {
                        Get_job_output: [
                          'Succeeded'
                        ]
                      }
                      type: 'ParseJson'
                    }
                  }
                  else: {
                    actions: {
                      'Send_an_email_(V2)_-_Notify_observers_2': {
                        inputs: {
                          body: {
                            Body: '<p>Hi,<br>\n<br>\nThe runbook that updates the status of the latest sync job failed. Please review the information below and act accordingly:<br>\nSolution: aad2spup<br>\nLogic App: @{workflow()[\'name\']}<br>\nJob Id: @{triggerBody()?[\'JobId\']}<br>\nStatus:: @{triggerBody()?[\'State\']?[\'Value\']}<br>\nSP List Item: <a href=\'@{triggerBody()?[\'{Link}\']}\' target=\'_blank\'>Go to list item</a><br>\nRun: <a href=\'@{concat(\'https://portal.azure.com/#view/Microsoft_Azure_EMA/LogicAppsMonitorBlade/runid/\', encodeUriComponent(workflow()?[\'run\'][\'id\']))}\' target=\'_blank\'>Go to run</a><br>\n</p>'
                            Importance: 'High'
                            Subject: 'aad2spup: runbook to update latest sync job status failed'
                            To: '@parameters(\'parObservers\')'
                          }
                          host: {
                            connection: {
                              name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
                            }
                          }
                          method: 'post'
                          path: '/v2/Mail'
                        }
                        runAfter: {
                        }
                        type: 'ApiConnection'
                      }
                      'Set_variable_-_varBreakLoop_3': {
                        inputs: {
                          name: 'varBreakLoop'
                          value: '@true'
                        }
                        runAfter: {
                          'Send_an_email_(V2)_-_Notify_observers_2': [
                            'Succeeded'
                          ]
                        }
                        type: 'SetVariable'
                      }
                    }
                  }
                  expression: {
                    and: [
                      {
                        equals: [
                          '@body(\'Create_job_-_run-UpdateSharePointUserProfilesSyncJobStatus\')?[\'properties\']?[\'provisioningState\']'
                          'Succeeded'
                        ]
                      }
                    ]
                  }
                  runAfter: {
                    'Create_job_-_run-UpdateSharePointUserProfilesSyncJobStatus': [
                      'Succeeded'
                      'TimedOut'
                      'Failed'
                    ]
                  }
                  type: 'If'
                }
                'Create_job_-_run-UpdateSharePointUserProfilesSyncJobStatus': {
                  inputs: {
                    body: {
                      properties: {
                        parameters: {
                          JobID: '@triggerBody()?[\'JobId\']'
                        }
                      }
                    }
                    host: {
                      connection: {
                        name: '@parameters(\'$connections\')[\'azureautomation\'][\'connectionId\']'
                      }
                    }
                    method: 'put'
                    path: '/subscriptions/@{encodeURIComponent(\'${subscription().subscriptionId}\')}/resourceGroups/@{encodeURIComponent(\'${resourceGroup().name}\')}/providers/Microsoft.Automation/automationAccounts/@{encodeURIComponent(\'${automationAccount_name}\')}/jobs'
                    queries: {
                      runbookName: runbooks[1].name
                      wait: true
                      'x-ms-api-version': '2015-10-31'
                    }
                  }
                  runAfter: {
                    Delay: [
                      'Succeeded'
                    ]
                  }
                  type: 'ApiConnection'
                }
                Delay: {
                  inputs: {
                    interval: {
                      count: '@variables(\'varMinutesToDelay\')'
                      unit: 'Minute'
                    }
                  }
                  runAfter: {
                  }
                  type: 'Wait'
                }
              }
              expression: '@equals(variables(\'varBreakLoop\'), true)'
              limit: {
                count: 3
                timeout: 'P7D'
              }
              runAfter: {
              }
              type: 'Until'
            }
          }
          case: 'Submitted'
        }
      }
      default: {
        actions: {
          'Send_an_email_(V2)_-_Notify_observers_3': {
            inputs: {
              body: {
                Body: '<p>Hi,<br>\n<br>\nThe latest sync job could not be started. Please review the information below and act accordingly:<br>\nSolution: aad2spup<br>\nLogic App: @{workflow()[\'name\']}<br>\nJob Id: @{triggerBody()?[\'JobId\']}<br>\nStatus:: @{triggerBody()?[\'State\']?[\'Value\']}<br>\nSP List Item: <a href=\'@{triggerBody()?[\'{Link}\']}\' target=\'_blank\'>Go to list item</a><br>\nRun: <a href=\'@{concat(\'https://portal.azure.com/#view/Microsoft_Azure_EMA/LogicAppsMonitorBlade/runid/\', encodeUriComponent(workflow()?[\'run\'][\'id\']))}\' target=\'_blank\'>Go to run</a><br>\n</p>'
                Importance: 'High'
                Subject: 'aad2spup: latest sync job could not be started'
                To: '@parameters(\'parObservers\')'
              }
              host: {
                connection: {
                  name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
                }
              }
              method: 'post'
              path: '/v2/Mail'
            }
            runAfter: {
            }
            type: 'ApiConnection'
          }
        }
      }
      expression: '@triggerBody()?[\'State\']?[\'Value\']'
      runAfter: {
        'Initialize_variable_-_varMinutesToDelay': [
          'Succeeded'
        ]
      }
      type: 'Switch'
    }
  }
}

var automationAccount_runbook_schedule_startTime = dateTimeAdd(baseTime, 'P1D', 'yyyy-MM-ddT00:00:00+01:00')

resource automationAccount 'Microsoft.Automation/automationAccounts@2021-06-22' = {
  name: automationAccount_name
  location: location
  identity: {
    type: automationAccount_identity
  }
  tags: tags
  properties: {
    publicNetworkAccess: true
    disableLocalAuth: false
    sku: {
      name: automationAccount_sku
    }
  }
}

// resource child_automationAccount_psModules 'Microsoft.Automation/automationAccounts/modules@2015-10-31' = [for item in psModules: {
//   parent: automationAccount
//   name: item.name
//   location: location
//   tags: {
//   }
//   properties: {
//     contentLink: {
//       uri: item.url
//       version: item.version
//     }
//   }
// }]

resource child_automationAccount_runbooks 'Microsoft.Automation/automationAccounts/runbooks@2019-06-01' = [for item in runbooks: {
  parent: automationAccount
  name: item.name
  location: location
  tags: tags
  properties: item.properties
}]

resource apiConnections 'Microsoft.Web/connections@2016-06-01' = [for item in connections: {
  name: item.name
  location: location
  properties: item.properties
}]

resource child_automationAccount_schedule 'Microsoft.Automation/automationAccounts/schedules@2020-01-13-preview' = {
  parent: automationAccount
  name: 'sched-SyncSharePointUserProfilesFromAzureActiveDirectory'
  properties: {
    description: 'Schedule for the SyncSharePointUserProfilesFromAzureActiveDirectory runbook'
    startTime: automationAccount_runbook_schedule_startTime
    expiryTime: '2999-12-31T23:59:59+00:00'
    interval: 1
    frequency: 'Week'
    timeZone: 'Europe/Amsterdam'
    advancedSchedule: {
      weekDays: [
        'Friday'
      ]
    }
  }
  dependsOn: [
    child_automationAccount_runbooks
  ]
}

resource child_automationAccount_jobSchedule 'Microsoft.Automation/automationAccounts/jobSchedules@2022-08-08' = {
  name: 'job-SyncSharePointUserProfilesFromAzureActiveDirectory'
  parent: automationAccount
  properties: {
    runbook: {
      name: runbook_SyncSharePointUserProfilesFromAzureActiveDirectory_name
    }
    schedule: {
      name: child_automationAccount_schedule.name
    }
  }
}

resource azureLogicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: logicApp_name
  location: location
  tags: tags
  identity: {
    type: logicApp_identity
  }
  properties: {
    state: logicApp.state
    definition: {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#'
      contentVersion: '1.0.0.0'
      parameters: logicApp.parameters
      triggers: logicApp.triggers
      actions: logicApp.actions
      outputs: {
      }
    }
    parameters: {
      '$connections': {
        value: {
          sharepointonline: {
            id: '${subscription().id}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_sharepointonline_name}'
            connectionId: resourceId('Microsoft.Web/connections', connections_sharepointonline_name)
            connectionName: connections_sharepointonline_name
          }
          azureautomation: {
            id: '${subscription().id}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_azureautomation_name}'
            connectionId: resourceId('Microsoft.Web/connections', connections_azureautomation_name)
            connectionName: connections_azureautomation_name
          }
          office365: {
            id: '${subscription().id}/providers/Microsoft.Web/locations/${location}/managedApis/${connections_office365_name}'
            connectionId: resourceId('Microsoft.Web/connections', connections_office365_name)
            connectionName: connections_office365_name
          }
        }
      }
    }
  }
  dependsOn: [
    automationAccount
    apiConnections
  ]
}

output LogicAppIdentityObjectId string = azureLogicApp.identity.principalId
output AutomationAccountIdentityObjectId string = automationAccount.identity.principalId
output AutomationAccountName string = automationAccount_name
output AutomationConnectionName string = connections_azureautomation_name
output SharePointConnectionName string = connections_sharepointonline_name
output Office365ConnectionName string = connections_office365_name
