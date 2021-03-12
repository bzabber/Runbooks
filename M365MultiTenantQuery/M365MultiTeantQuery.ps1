##################################################################################################
#- Name: M365MultiTenantQuery.ps1                                                               -#
#- Date: March 11, 2021                                                                         -#
#- Description: This script will leverage the Office 365 service communications API to pull     -# 
#-              service health and messages and feed into Log Analytics                         -#
#- Dependencies:                                                                                -#
#- 	- Azure Service Principal (Registered Client App) with API read permissions                 -#
#- 	- Log analytics workspace and key                                                           -#
#-      - OMSIngestionAPI v1.6.0 available from PowerShell Gallery:                             -# 
#-        https://www.powershellgallery.com/packages/OMSIngestionAPI/1.6.0                      -#
#-      -  M365Monitoring PowerShell Module                                                     -#
#- 	                                                                                            -#
##################################################################################################

#--- Include module to format and send request to OMS ---#
Import-Module OMSIngestionAPI
#--- Include the module to query M365 subscriptions.
Import-module M365Monitoring

#--- Get Log Analytics authentication info from variables ---#
$CustomerId = Get-AutomationVariable -Name 'OMSWorkSpaceID'
$SharedKey = Get-AutomationVariable -Name 'OMSPrimaryKey'

###################################################################################
#- See Section 4.6.1 and 4.6.2 of MSM14                                          -#
#- Array of tenants and the required parameters to authenticate with M365 tenant -#
#- You need to add a new array for each tenant that will be queried.             -#
#- Each Hashtable value will reference associated value in Azure Automation      -#
#- Variables.                                                                    -#
#- Each Azure Variable will use the following naming convention:                 -#
#- Dept Abbreviation<VariableName>                                               -#
#- Ex:                                                                           -#
#-     SSCTenantID                                                               -#
#-     SSCClientID                                                               -#
#-     SSCClientSecret                                                           -#
#-  TenantName is used to create a friendly name for the O365 subscription.      -#
#-  TenantName is used by the O365 Incident Workbook to scope the data to a      -#
#-  specific O365 subscription                                                   -#
###################################################################################
$tenants = @(
  [pscustomobject]@{
        TenantName   = "BryanZ-O365";
        TenantID     = Get-AutomationVariable -Name 'BZM365TenantID';
        ClientID     = Get-AutomationVariable -Name 'BZM365ClientID';
        ClientSecret = Get-AutomationVariable -Name 'BZM365ClientSecret';
    },
    [pscustomobject]@{
        TenantName   = "BrianK-O365";
        TenantID     = Get-AutomationVariable -Name 'BKM365TenantID';
        ClientID     = Get-AutomationVariable -Name 'BKClientID';
        ClientSecret = Get-AutomationVariable -Name 'BKClientSecret';
    },
    [pscustomobject]@{
        TenantName   = "SSC-SPC";
        TenantID     = Get-AutomationVariable -Name 'SSCM365TenantID';
        ClientID     = Get-AutomationVariable -Name 'SSCClientID';
        ClientSecret = Get-AutomationVariable -Name 'SSCClientSecret';
    },
    [pscustomobject]@{
        TenantName  = "IRB";
        TenantID    = Get-AutomationVariable -Name 'IRBTenantID';
        ClientID    = Get-AutomationVariable -Name 'IRBClientID';
        ClientSecret = Get-AutomationVariable -Name 'IRBClientSecret';
    },
    [pscustomobject]@{
        TenantName  = "ACOA";
        TenantID    = Get-AutomationVariable -Name 'ACOATenantID';
        ClientID    = Get-AutomationVariable -Name 'ACOAClientID';
        ClientSecret = Get-AutomationVariable -Name 'ACOAClientSecret';
    }
    <#

    
    #>

) 

#--- Query M365 Service Health Dashboard via O365 Services Communications API ---#
$Servicehealth = $tenants | foreach { get-M365ServiceHealth -TenantID $_.TenantID -ClientID $_.ClientID -ClientSecret $_.ClientSecret -TenantName $_.TenantName }
#write-output $Servicehealth

$JSON = $Servicehealth | ConvertTo-Json -Depth 10
#write-output $JSON

#--- Set the name of the log that will be created/appended to in Log Analytics. ---#
$LogType = "O365ServiceHealth"

#--- Submit the ServiceHealth Data to Log Analytics API endpoint. ---#

Send-OMSAPIIngestionFile -customerId $CustomerID -sharedKey $SharedKey -body $JSON -logType $LogType -Verbose

#--- Query M365 Message Center via O365 Services Communications API ---#
$Messages = $tenants | foreach { get-M365Messages -TenantID $_.TenantID -ClientID $_.ClientID -ClientSecret $_.ClientSecret -TenantName $_.TenantName }
write-output $Messages

#Convert Messages to JSON format before sending to Log Analytics.
$JSON = $Messages | ConvertTo-Json -Depth 10
#write-output $JSON

#--- Set the name of the log that will be created/appended to in Log Analytics. ---#
$LogType = "O365MessageCenter"

#--- Submit the ServiceHealth Data to Log Analytics API endpoint. ---#
Send-OMSAPIIngestionFile -customerId $CustomerID -sharedKey $SharedKey -body $JSON -logType $LogType -Verbose

