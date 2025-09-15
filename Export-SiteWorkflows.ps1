<#
.SYNOPSIS
    Exports SharePoint 2013 workflows (definitions and subscriptions) using CSOM via PnP.PowerShell.

.DESCRIPTION
    Connects to a SharePoint Online site using PnP.PowerShell and exports all workflow definitions (XAML) and workflow subscriptions (JSON) to a specified folder.

.NOTES
    No warranty. Use at your own risk.
    Author: Marko Rosberg (marko.rosberg@creativus.fi)
#>



param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$SiteURL = "https://company.sharepoint.com/sites/site",

    [Parameter(Mandatory=$false, Position=1)]
    [string]$FileFolder = "C:\Temp\Workflows\site",

    [Parameter(Mandatory=$false)]
    [string]$PnPAppId = "00000000-0000-0000-0000-000000000000"  # Replace with your PnP App ID (for interactive authentication()
)

# Ensure the output directory exists
if (-not (Test-Path -Path $fileFolder)) {
    New-Item -ItemType Directory -Path $fileFolder | Out-Null
}

# Connect interactively with PnP to obtain an access token for SharePoint
Write-Host "Connect to SharePoint using PnP interactive sign-in..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $PnPAppId -ErrorAction Stop


# Connect to Site
$ctx = (Get-PnPWeb).Context

# find the WorkflowServices assembly that PnP loaded
$wsAsm = [AppDomain]::CurrentDomain.GetAssemblies() |
    Where-Object { $_.GetName().Name -eq 'Microsoft.SharePoint.Client.WorkflowServices' } |
    Select-Object -First 1

if (-not $wsAsm) {
    throw "Microsoft.SharePoint.Client.WorkflowServices assembly not found in AppDomain. Ensure PnP.PowerShell loaded CSOM or install the CSOM package instead of Add-Type-ing PnP's DLLs."
}

# get the WorkflowServicesManager type from the same assembly and create an instance via Activator
$wsManagerType = $wsAsm.GetType('Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager')
if (-not $wsManagerType) {
    throw "Type 'Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager' not found in the loaded WorkflowServices assembly."
}

$WorkflowServicesManager = [System.Activator]::CreateInstance($wsManagerType, $ctx, $ctx.Web)

# get the WorkflowDeploymentService and WorkflowSubscriptionService
$WorkflowDeploymentService = $WorkflowServicesManager.GetWorkflowDeploymentService()
$WorkflowSubscriptionService = $WorkflowServicesManager.GetWorkflowSubscriptionService()

# get enumerators for definitions and subscriptions
$definitions = $WorkflowDeploymentService.EnumerateDefinitions($true)
$subscriptions = $WorkflowSubscriptionService.EnumerateSubscriptions()

# load and execute
$ctx.Load($definitions)
$ctx.Load($subscriptions)
$ctx.ExecuteQuery()

# export subscription details for each workflow subscription
foreach ($sub in $subscriptions) {   
    $fileName = "WorkflowSubscription_$($sub.Name)_$($sub.Id).json"
    $filePath = Join-Path -Path $fileFolder -ChildPath $fileName
    $sub | ConvertTo-Json -Depth 2 | Out-File -FilePath $filePath -Encoding UTF8
    Write-Host "Exported workflow subscription: $filePath" -ForegroundColor Green
}

# export xaml definitions for each workflow definition
foreach ($def in $definitions) {   
    $fileName = "WorkflowDefinition_$($def.DisplayName)_$($def.Id).xaml"
    $filePath = Join-Path -Path $fileFolder -ChildPath $fileName
    [System.IO.File]::WriteAllText($filePath, $def.Xaml)
    Write-Host "Exported workflow definition: $filePath" -ForegroundColor Green
}