<#
.SYNOPSIS
    Parse all personal workspaces in Power BI and display reports, datasets, and dashboards.

.DESCRIPTION
    This script connects to the Power BI service and retrieves all personal workspaces.
    For each workspace, it lists the reports, datasets, and dashboards contained within.
    Results are organized and displayed by workspace.
    Optionally, you can delete all content from personal workspaces.

.PARAMETER Delete
    When specified, deletes all content (reports, datasets, dashboards, dataflows) from personal workspaces.
    Requires confirmation unless -AutomationMode is used.

.PARAMETER WhatIf
    Shows what would be deleted without actually performing the deletion.

.PARAMETER AutomationMode
    Enables non-interactive mode for Azure Automation. Uses Managed Identity for authentication.
    Skips all prompts and confirmations. Use with -Delete for scheduled cleanup.

.NOTES
    Author: KYOS SA
    Date: 2025-12-17
    Requires: MicrosoftPowerBIMgmt PowerShell module
    Permissions: Power BI Admin or appropriate workspace access
    
    For Azure Automation with Managed Identity:
    1. Assign 'Fabric Administrator' role to the Managed Identity in Entra ID
    2. In Power BI Admin Portal > Tenant settings:
       - Enable 'Service principals can use Fabric APIs'
       - Enable 'Service principals can access read-only admin APIs'
       - Add the Managed Identity to both settings

.EXAMPLE
    .\pbi-personalworkspace.ps1
    Lists all personal workspaces and their contents.

.EXAMPLE
    .\pbi-personalworkspace.ps1 -WhatIf
    Shows what would be deleted from all personal workspaces without performing deletions.

.EXAMPLE
    .\pbi-personalworkspace.ps1 -Delete
    Deletes all content from personal workspaces after confirmation.

.EXAMPLE
    .\pbi-personalworkspace.ps1 -Delete -AutomationMode
    Deletes all content from personal workspaces without prompts (for Azure Automation).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Delete,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [switch]$AutomationMode
)

<# Requires -Modules MicrosoftPowerBIMgmt #>

# Check if the Power BI module is installed
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    if ($AutomationMode) {
        Write-Output "Installing MicrosoftPowerBIMgmt module..."
    } else {
        Write-Host "Installing MicrosoftPowerBIMgmt module..." -ForegroundColor Yellow
    }
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}

# Import the module
Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop

# Connect to Power BI Service
if ($AutomationMode) {
    Write-Output "Connecting to Power BI Service using Managed Identity..."
} else {
    Write-Host "Connecting to Power BI Service..." -ForegroundColor Cyan
}

try {
    if ($AutomationMode) {
        # Use Managed Identity for Azure Automation
        Connect-PowerBIServiceAccount -ServicePrincipal -Credential (New-Object System.Management.Automation.PSCredential(
            (Get-AutomationVariable -Name 'PowerBI_AppId' -ErrorAction SilentlyContinue) ?? $env:AZURE_CLIENT_ID,
            (ConvertTo-SecureString "dummy" -AsPlainText -Force)
        )) -TenantId ((Get-AutomationVariable -Name 'PowerBI_TenantId' -ErrorAction SilentlyContinue) ?? $env:AZURE_TENANT_ID) -ErrorAction SilentlyContinue
        
        # If service principal auth fails, try managed identity directly
        if (-not (Get-PowerBIAccessToken -ErrorAction SilentlyContinue)) {
            # For System-Assigned Managed Identity, use Connect-AzAccount first
            Connect-AzAccount -Identity -ErrorAction Stop
            $token = (Get-AzAccessToken -ResourceUrl "https://analysis.windows.net/powerbi/api").Token
            Connect-PowerBIServiceAccount -AccessToken $token -ErrorAction Stop
        }
        Write-Output "Successfully connected to Power BI Service using Managed Identity"
    }
    else {
        # Interactive login for manual execution
        Connect-PowerBIServiceAccount -ErrorAction Stop
        Write-Host "Successfully connected to Power BI Service" -ForegroundColor Green
    }
}
catch {
    if ($AutomationMode) {
        Write-Error "Failed to connect to Power BI Service: $_"
    } else {
        Write-Error "Failed to connect to Power BI Service: $_"
    }
    exit 1
}

# Function to invoke Power BI REST API with retry logic for rate limiting (429)
function Invoke-PowerBIRestMethodWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$Method,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 2
    )
    
    $retryCount = 0
    $delay = $InitialDelaySeconds
    
    while ($true) {
        try {
            $response = Invoke-PowerBIRestMethod -Url $Url -Method $Method -ErrorAction Stop
            return $response
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            # Check if it's a 429 (Too Many Requests) error
            if ($errorMessage -match "429" -or $errorMessage -match "Too Many Requests") {
                $retryCount++
                if ($retryCount -gt $MaxRetries) {
                    Write-Warning "Max retries ($MaxRetries) exceeded for $Url"
                    return $null
                }
                Write-Host "      Rate limited (429). Waiting $delay seconds before retry $retryCount/$MaxRetries..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $delay
                $delay = $delay * 2  # Exponential backoff
            }
            else {
                # Not a rate limit error, throw it
                throw $_
            }
        }
    }
}

# Function to delete workspace contents using Admin REST API
function Remove-WorkspaceContents {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$WorkspaceData,
        [Parameter(Mandatory = $false)]
        [switch]$WhatIfMode
    )

    $deletedItems = @{
        Reports = 0
        Datasets = 0
        Dashboards = 0
        Dataflows = 0
    }

    $workspaceId = $WorkspaceData.WorkspaceId
    $workspaceName = $WorkspaceData.WorkspaceName

    # Delete Reports
    foreach ($report in $WorkspaceData.Reports) {
        $reportName = if ($report.name) { $report.name } else { $report.Name }
        $reportId = if ($report.id) { $report.id } else { $report.Id }
        
        if ($WhatIfMode) {
            Write-Host "    [WhatIf] Would delete report: $reportName ($reportId)" -ForegroundColor Yellow
            $deletedItems.Reports++
        }
        else {
            try {
                Invoke-PowerBIRestMethod -Url "admin/reports/$reportId" -Method Delete -ErrorAction Stop
                Write-Host "    Deleted report: $reportName" -ForegroundColor Red
                $deletedItems.Reports++
            }
            catch {
                Write-Warning "    Failed to delete report '$reportName': $($_.Exception.Message)"
            }
        }
    }

    # Delete Dashboards
    foreach ($dashboard in $WorkspaceData.Dashboards) {
        $dashboardName = if ($dashboard.name) { $dashboard.name } else { $dashboard.Name }
        $dashboardId = if ($dashboard.id) { $dashboard.id } else { $dashboard.Id }
        
        if ($WhatIfMode) {
            Write-Host "    [WhatIf] Would delete dashboard: $dashboardName ($dashboardId)" -ForegroundColor Yellow
            $deletedItems.Dashboards++
        }
        else {
            try {
                Invoke-PowerBIRestMethod -Url "admin/dashboards/$dashboardId" -Method Delete -ErrorAction Stop
                Write-Host "    Deleted dashboard: $dashboardName" -ForegroundColor Red
                $deletedItems.Dashboards++
            }
            catch {
                Write-Warning "    Failed to delete dashboard '$dashboardName': $($_.Exception.Message)"
            }
        }
    }

    # Delete Dataflows
    foreach ($dataflow in $WorkspaceData.Dataflows) {
        $dataflowName = if ($dataflow.name) { $dataflow.name } else { $dataflow.Name }
        $dataflowId = if ($dataflow.objectId) { $dataflow.objectId } else { $dataflow.ObjectId }
        
        if ($WhatIfMode) {
            Write-Host "    [WhatIf] Would delete dataflow: $dataflowName ($dataflowId)" -ForegroundColor Yellow
            $deletedItems.Dataflows++
        }
        else {
            try {
                Invoke-PowerBIRestMethod -Url "groups/$workspaceId/dataflows/$dataflowId" -Method Delete -ErrorAction Stop
                Write-Host "    Deleted dataflow: $dataflowName" -ForegroundColor Red
                $deletedItems.Dataflows++
            }
            catch {
                Write-Warning "    Failed to delete dataflow '$dataflowName': $($_.Exception.Message)"
            }
        }
    }

    # Delete Datasets (must be deleted last as reports depend on them)
    foreach ($dataset in $WorkspaceData.Datasets) {
        $datasetName = if ($dataset.name) { $dataset.name } else { $dataset.Name }
        $datasetId = if ($dataset.id) { $dataset.id } else { $dataset.Id }
        
        if ($WhatIfMode) {
            Write-Host "    [WhatIf] Would delete dataset: $datasetName ($datasetId)" -ForegroundColor Yellow
            $deletedItems.Datasets++
        }
        else {
            try {
                Invoke-PowerBIRestMethod -Url "admin/datasets/$datasetId" -Method Delete -ErrorAction Stop
                Write-Host "    Deleted dataset: $datasetName" -ForegroundColor Red
                $deletedItems.Datasets++
            }
            catch {
                Write-Warning "    Failed to delete dataset '$datasetName': $($_.Exception.Message)"
            }
        }
    }

    return $deletedItems
}

# Function to get workspace contents using Admin REST API
function Get-WorkspaceContents {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceName,
        [Parameter(Mandatory = $false)]
        [switch]$UseAdminScope
    )

    $workspaceData = [PSCustomObject]@{
        WorkspaceName = $WorkspaceName
        WorkspaceId   = $WorkspaceId
        Reports       = @()
        Datasets      = @()
        Dashboards    = @()
        Dataflows     = @()
    }

    # Use Admin API endpoints for accessing other users' workspaces
    if ($UseAdminScope) {
        # Get Reports via Admin API
        try {
            $reportsResponse = Invoke-PowerBIRestMethodWithRetry -Url "admin/groups/$WorkspaceId/reports" -Method Get
            if ($reportsResponse) {
                $reports = ($reportsResponse | ConvertFrom-Json).value
                if ($reports) {
                    $workspaceData.Reports = $reports | Select-Object name, id, webUrl, datasetId
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve reports for workspace: $WorkspaceName - $($_.Exception.Message)"
        }

        # Get Datasets via Admin API
        try {
            $datasetsResponse = Invoke-PowerBIRestMethodWithRetry -Url "admin/groups/$WorkspaceId/datasets" -Method Get
            if ($datasetsResponse) {
                $datasets = ($datasetsResponse | ConvertFrom-Json).value
                if ($datasets) {
                    $workspaceData.Datasets = $datasets | Select-Object name, id, configuredBy, isRefreshable, isOnPremGatewayRequired
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve datasets for workspace: $WorkspaceName - $($_.Exception.Message)"
        }

        # Get Dashboards via Admin API
        try {
            $dashboardsResponse = Invoke-PowerBIRestMethodWithRetry -Url "admin/groups/$WorkspaceId/dashboards" -Method Get
            if ($dashboardsResponse) {
                $dashboards = ($dashboardsResponse | ConvertFrom-Json).value
                if ($dashboards) {
                    $workspaceData.Dashboards = $dashboards | Select-Object name, id, embedUrl, isReadOnly
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve dashboards for workspace: $WorkspaceName - $($_.Exception.Message)"
        }

        # Get Dataflows via Admin API
        try {
            $dataflowsResponse = Invoke-PowerBIRestMethodWithRetry -Url "admin/groups/$WorkspaceId/dataflows" -Method Get
            if ($dataflowsResponse) {
                $dataflows = ($dataflowsResponse | ConvertFrom-Json).value
                if ($dataflows) {
                    $workspaceData.Dataflows = $dataflows | Select-Object name, objectId, configuredBy
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve dataflows for workspace: $WorkspaceName - $($_.Exception.Message)"
        }
    }
    else {
        # Use standard cmdlets for Individual scope
        try {
            $reports = Get-PowerBIReport -WorkspaceId $WorkspaceId -ErrorAction SilentlyContinue
            if ($reports) {
                $workspaceData.Reports = $reports | Select-Object Name, Id, WebUrl, DatasetId
            }
        }
        catch {
            Write-Warning "Could not retrieve reports for workspace: $WorkspaceName"
        }

        try {
            $datasets = Get-PowerBIDataset -WorkspaceId $WorkspaceId -ErrorAction SilentlyContinue
            if ($datasets) {
                $workspaceData.Datasets = $datasets | Select-Object Name, Id, ConfiguredBy, IsRefreshable, IsOnPremGatewayRequired
            }
        }
        catch {
            Write-Warning "Could not retrieve datasets for workspace: $WorkspaceName"
        }

        try {
            $dashboards = Get-PowerBIDashboard -WorkspaceId $WorkspaceId -ErrorAction SilentlyContinue
            if ($dashboards) {
                $workspaceData.Dashboards = $dashboards | Select-Object Name, Id, EmbedUrl, IsReadOnly
            }
        }
        catch {
            Write-Warning "Could not retrieve dashboards for workspace: $WorkspaceName"
        }

        try {
            $dataflows = Get-PowerBIDataflow -WorkspaceId $WorkspaceId -ErrorAction SilentlyContinue
            if ($dataflows) {
                $workspaceData.Dataflows = $dataflows | Select-Object Name, ObjectId, ConfiguredBy
            }
        }
        catch {
            Write-Warning "Could not retrieve dataflows for workspace: $WorkspaceName"
        }
    }

    return $workspaceData
}

# Function to display workspace contents
function Show-WorkspaceContents {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$WorkspaceData
    )

    Write-Host "`n" -NoNewline
    Write-Host "=" * 80 -ForegroundColor DarkGray
    Write-Host "WORKSPACE: $($WorkspaceData.WorkspaceName)" -ForegroundColor Cyan
    Write-Host "ID: $($WorkspaceData.WorkspaceId)" -ForegroundColor DarkCyan
    Write-Host "=" * 80 -ForegroundColor DarkGray

    # Display Reports
    Write-Host "`n  REPORTS ($($WorkspaceData.Reports.Count)):" -ForegroundColor Yellow
    if ($WorkspaceData.Reports.Count -gt 0) {
        $WorkspaceData.Reports | ForEach-Object {
            Write-Host "    - $($_.Name)" -ForegroundColor White
            Write-Host "      ID: $($_.Id)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "    (No reports found)" -ForegroundColor DarkGray
    }

    # Display Datasets
    Write-Host "`n  DATASETS ($($WorkspaceData.Datasets.Count)):" -ForegroundColor Yellow
    if ($WorkspaceData.Datasets.Count -gt 0) {
        $WorkspaceData.Datasets | ForEach-Object {
            Write-Host "    - $($_.Name)" -ForegroundColor White
            Write-Host "      ID: $($_.Id) | Configured By: $($_.ConfiguredBy)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "    (No datasets found)" -ForegroundColor DarkGray
    }

    # Display Dashboards
    Write-Host "`n  DASHBOARDS ($($WorkspaceData.Dashboards.Count)):" -ForegroundColor Yellow
    if ($WorkspaceData.Dashboards.Count -gt 0) {
        $WorkspaceData.Dashboards | ForEach-Object {
            Write-Host "    - $($_.Name)" -ForegroundColor White
            Write-Host "      ID: $($_.Id)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "    (No dashboards found)" -ForegroundColor DarkGray
    }

    # Display Dataflows
    Write-Host "`n  DATAFLOWS ($($WorkspaceData.Dataflows.Count)):" -ForegroundColor Yellow
    if ($WorkspaceData.Dataflows.Count -gt 0) {
        $WorkspaceData.Dataflows | ForEach-Object {
            Write-Host "    - $($_.Name)" -ForegroundColor White
            Write-Host "      ID: $($_.ObjectId) | Configured By: $($_.ConfiguredBy)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "    (No dataflows found)" -ForegroundColor DarkGray
    }
}

# Main script execution
if ($AutomationMode) {
    Write-Output "Retrieving Power BI Workspaces..."
} else {
    Write-Host "`nRetrieving Power BI Workspaces..." -ForegroundColor Cyan
}

# Determine if WhatIf mode is active (using our manual switch)
$isWhatIfMode = $WhatIf.IsPresent

# Get all workspaces - Personal workspaces have type "PersonalGroup" or are "My Workspace"
# Using -Scope Organization requires admin permissions, -Scope Individual gets user's accessible workspaces
$useAdminScope = $false
try {
    # Try to get all workspaces as admin first
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All -ErrorAction Stop
    if ($AutomationMode) {
        Write-Output "Retrieved workspaces using Organization scope (Admin mode)"
    } else {
        Write-Host "Retrieved workspaces using Organization scope (Admin mode)" -ForegroundColor Green
    }
    $useAdminScope = $true
}
catch {
    if ($AutomationMode) {
        Write-Output "Admin scope not available, using Individual scope..."
    } else {
        Write-Host "Admin scope not available, using Individual scope..." -ForegroundColor Yellow
    }
    $workspaces = Get-PowerBIWorkspace -Scope Individual -All
}

# Filter for personal workspaces (Type = "PersonalGroup" only)
# Personal workspaces in Power BI have Type = "PersonalGroup"
$personalWorkspaces = $workspaces | Where-Object { 
    $_.Type -eq "PersonalGroup"
}

# If no personal workspaces found with filter, show all workspaces
if ($personalWorkspaces.Count -eq 0) {
    if ($AutomationMode) {
        Write-Output "No personal workspaces found with strict filter. Showing all accessible workspaces..."
    } else {
        Write-Host "No personal workspaces found with strict filter. Showing all accessible workspaces..." -ForegroundColor Yellow
    }
    $personalWorkspaces = $workspaces
}

if ($AutomationMode) {
    Write-Output "Found $($personalWorkspaces.Count) workspace(s) to process"
} else {
    Write-Host "`nFound $($personalWorkspaces.Count) workspace(s) to process" -ForegroundColor Cyan
}

# Check if Delete mode is enabled
$isDeleteMode = $Delete -or $isWhatIfMode

if ($isDeleteMode -and -not $isWhatIfMode -and -not $AutomationMode) {
    Write-Host "`n" -NoNewline
    Write-Host "!" * 80 -ForegroundColor Red
    Write-Host "WARNING: You are about to DELETE all content from personal workspaces!" -ForegroundColor Red
    Write-Host "!" * 80 -ForegroundColor Red
    $confirmation = Read-Host "`nType 'YES' to confirm deletion"
    if ($confirmation -ne 'YES') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Disconnect-PowerBIServiceAccount
        exit 0
    }
}

if ($isDeleteMode -and -not $isWhatIfMode -and $AutomationMode) {
    Write-Output "WARNING: AutomationMode enabled - proceeding with deletion without confirmation"
}

if ($isWhatIfMode) {
    if ($AutomationMode) {
        Write-Output "[WhatIf Mode] Showing what would be deleted..."
    } else {
        Write-Host "`n[WhatIf Mode] Showing what would be deleted..." -ForegroundColor Yellow
    }
}

# Store all workspace data for export
$allWorkspaceData = @()

# Track deletion statistics
$totalDeleted = @{
    Reports = 0
    Datasets = 0
    Dashboards = 0
    Dataflows = 0
}

# Process each workspace
$counter = 0
foreach ($workspace in $personalWorkspaces) {
    $counter++
    Write-Progress -Activity "Processing Workspaces" -Status "$counter of $($personalWorkspaces.Count): $($workspace.Name)" -PercentComplete (($counter / $personalWorkspaces.Count) * 100)
    
    # Use admin scope if available to access other users' workspace contents
    if ($useAdminScope) {
        $workspaceContents = Get-WorkspaceContents -WorkspaceId $workspace.Id -WorkspaceName $workspace.Name -UseAdminScope
    } else {
        $workspaceContents = Get-WorkspaceContents -WorkspaceId $workspace.Id -WorkspaceName $workspace.Name
    }
    $allWorkspaceData += $workspaceContents
    
    # Display the contents
    Show-WorkspaceContents -WorkspaceData $workspaceContents
    
    # Delete contents if requested
    if ($isDeleteMode) {
        $hasContent = ($workspaceContents.Reports.Count -gt 0) -or 
                      ($workspaceContents.Datasets.Count -gt 0) -or 
                      ($workspaceContents.Dashboards.Count -gt 0) -or 
                      ($workspaceContents.Dataflows.Count -gt 0)
        
        if ($hasContent) {
            if ($isWhatIfMode) {
                Write-Host "`n  [WhatIf] Items that would be deleted:" -ForegroundColor Yellow
            } else {
                Write-Host "`n  Deleting content..." -ForegroundColor Red
            }
            
            $deleted = Remove-WorkspaceContents -WorkspaceData $workspaceContents -WhatIfMode:$isWhatIfMode
            $totalDeleted.Reports += $deleted.Reports
            $totalDeleted.Datasets += $deleted.Datasets
            $totalDeleted.Dashboards += $deleted.Dashboards
            $totalDeleted.Dataflows += $deleted.Dataflows
        }
    }
}

Write-Progress -Activity "Processing Workspaces" -Completed

# Summary
if ($AutomationMode) {
    Write-Output "================================================================================"
    Write-Output "SUMMARY"
    Write-Output "================================================================================"
} else {
    Write-Host "`n" -NoNewline
    Write-Host "=" * 80 -ForegroundColor DarkGray
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor DarkGray
}

$totalReports = ($allWorkspaceData | ForEach-Object { $_.Reports.Count } | Measure-Object -Sum).Sum
$totalDatasets = ($allWorkspaceData | ForEach-Object { $_.Datasets.Count } | Measure-Object -Sum).Sum
$totalDashboards = ($allWorkspaceData | ForEach-Object { $_.Dashboards.Count } | Measure-Object -Sum).Sum
$totalDataflows = ($allWorkspaceData | ForEach-Object { $_.Dataflows.Count } | Measure-Object -Sum).Sum

if ($AutomationMode) {
    Write-Output "Total Workspaces: $($allWorkspaceData.Count)"
    Write-Output "Total Reports: $totalReports"
    Write-Output "Total Datasets: $totalDatasets"
    Write-Output "Total Dashboards: $totalDashboards"
    Write-Output "Total Dataflows: $totalDataflows"
} else {
    Write-Host "`nTotal Workspaces: $($allWorkspaceData.Count)" -ForegroundColor White
    Write-Host "Total Reports: $totalReports" -ForegroundColor White
    Write-Host "Total Datasets: $totalDatasets" -ForegroundColor White
    Write-Host "Total Dashboards: $totalDashboards" -ForegroundColor White
    Write-Host "Total Dataflows: $totalDataflows" -ForegroundColor White
}

# Show deletion summary if applicable
if ($isDeleteMode) {
    $actionWord = if ($isWhatIfMode) { "Would delete" } else { "Deleted" }
    
    if ($AutomationMode) {
        Write-Output "================================================================================"
        if ($isWhatIfMode) {
            Write-Output "DELETION PREVIEW (WhatIf)"
        } else {
            Write-Output "DELETION SUMMARY"
        }
        Write-Output "================================================================================"
        Write-Output "$actionWord Reports: $($totalDeleted.Reports)"
        Write-Output "$actionWord Datasets: $($totalDeleted.Datasets)"
        Write-Output "$actionWord Dashboards: $($totalDeleted.Dashboards)"
        Write-Output "$actionWord Dataflows: $($totalDeleted.Dataflows)"
    } else {
        Write-Host "`n" -NoNewline
        Write-Host "=" * 80 -ForegroundColor DarkGray
        if ($isWhatIfMode) {
            Write-Host "DELETION PREVIEW (WhatIf)" -ForegroundColor Yellow
        } else {
            Write-Host "DELETION SUMMARY" -ForegroundColor Red
        }
        Write-Host "=" * 80 -ForegroundColor DarkGray
        Write-Host "$actionWord Reports: $($totalDeleted.Reports)" -ForegroundColor $(if ($isWhatIfMode) { 'Yellow' } else { 'Red' })
        Write-Host "$actionWord Datasets: $($totalDeleted.Datasets)" -ForegroundColor $(if ($isWhatIfMode) { 'Yellow' } else { 'Red' })
        Write-Host "$actionWord Dashboards: $($totalDeleted.Dashboards)" -ForegroundColor $(if ($isWhatIfMode) { 'Yellow' } else { 'Red' })
        Write-Host "$actionWord Dataflows: $($totalDeleted.Dataflows)" -ForegroundColor $(if ($isWhatIfMode) { 'Yellow' } else { 'Red' })
    }
}

# Export option
# Export option (skip in AutomationMode unless explicitly needed)
if (-not $AutomationMode) {
    $exportPath = Join-Path -Path $PSScriptRoot -ChildPath "PowerBI_Workspace_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $exportChoice = Read-Host "`nWould you like to export the results to JSON? (Y/N)"

    if ($exportChoice -eq "Y" -or $exportChoice -eq "y") {
        $allWorkspaceData | ConvertTo-Json -Depth 10 | Out-File -FilePath $exportPath -Encoding UTF8
        Write-Host "Results exported to: $exportPath" -ForegroundColor Green
    }
}

# Disconnect from Power BI Service
Disconnect-PowerBIServiceAccount
if ($AutomationMode) {
    Write-Output "Disconnected from Power BI Service"
    Write-Output "Script completed successfully at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
} else {
    Write-Host "`nDisconnected from Power BI Service" -ForegroundColor Cyan
}
