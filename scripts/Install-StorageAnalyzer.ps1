<#
.SYNOPSIS
    Adds the Smart Storage Analyzer app to many SharePoint Online sites at once.

.DESCRIPTION
    The package provisions a "Storage Analyzer" page through a site-scoped feature, so it cannot be
    deployed tenant-wide. Each site needs the app added to it. This script does that in bulk:
    it enumerates sites (or uses the list you give it) and runs Install-PnPApp on each one,
    optionally upgrading sites that already have an older version.

    Prerequisites:
      * PnP.PowerShell module   (Install-Module PnP.PowerShell -Scope CurrentUser)
      * smart-storage-analyzer.sppkg already uploaded to the TENANT app catalog and deployed
        (click "Deploy" when prompted; do NOT tick "Make this solution available to all sites")
      * SharePoint administrator rights, plus an Entra ID app registration for PnP interactive
        sign-in (pass its client ID with -ClientId)

.PARAMETER AdminUrl
    Tenant admin URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER ClientId
    Client ID of the Entra ID app registration used for PnP PowerShell interactive sign-in.

.PARAMETER SiteUrls
    Install only on these sites. Use this OR -AllSites (one of the two is required).

.PARAMETER AllSites
    Install on every eligible site in the tenant: communication, team and classic team sites
    (OneDrive, search, app catalog and redirect sites are skipped). Use this OR -SiteUrls.
    Combine with -ExcludeUrls to leave specific sites out, and run with -WhatIf first.

.PARAMETER ExcludeUrls
    Site URLs to skip (matched exactly, case-insensitive).

.PARAMETER Update
    Also upgrade sites that already have an older version installed.

.PARAMETER AppId
    The solution ID from config/package-solution.json. Defaults to this project's ID.

.EXAMPLE
    # Preview what would happen on every eligible site (nothing is changed)
    .\Install-StorageAnalyzer.ps1 -AdminUrl https://contoso-admin.sharepoint.com -ClientId <guid> -AllSites -WhatIf

.EXAMPLE
    # Install on two specific sites
    .\Install-StorageAnalyzer.ps1 -AdminUrl https://contoso-admin.sharepoint.com -ClientId <guid> `
        -SiteUrls https://contoso.sharepoint.com/sites/finance, https://contoso.sharepoint.com/sites/legal

.EXAMPLE
    # Install on all sites and upgrade sites that have an older version
    .\Install-StorageAnalyzer.ps1 -AdminUrl https://contoso-admin.sharepoint.com -ClientId <guid> -AllSites -Update
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Specific')]
param(
    [Parameter(Mandatory = $true)] [string] $AdminUrl,
    [Parameter(Mandatory = $true)] [string] $ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'Specific')] [string[]] $SiteUrls,
    [Parameter(Mandatory = $true, ParameterSetName = 'All')] [switch] $AllSites,
    [string[]] $ExcludeUrls = @(),
    [switch] $Update,
    [string] $AppId = '4c10f106-ac82-4037-9553-c3cfdfaa443b'
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'The PnP.PowerShell module is required. Run: Install-Module PnP.PowerShell -Scope CurrentUser'
}

# Site templates that can host the app. Everything else (OneDrive, search, app catalog, redirect) is skipped.
$eligibleTemplatePrefixes = @('GROUP#0', 'SITEPAGEPUBLISHING#0', 'STS#3', 'STS#0')

Write-Host "Connecting to $AdminUrl ..."
Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId

if ($PSCmdlet.ParameterSetName -eq 'Specific') {
    $targets = $SiteUrls
} else {
    $targets = Get-PnPTenantSite | Where-Object {
        $t = $_.Template
        $eligibleTemplatePrefixes | Where-Object { $t -like "$_*" }
    } | Select-Object -ExpandProperty Url
}
$targets = @($targets | Where-Object { $ExcludeUrls -notcontains $_ } | Sort-Object -Unique)

Write-Host "$($targets.Count) site(s) to process."

$results = foreach ($url in $targets) {
    $status = 'Unknown'
    try {
        if ($PSCmdlet.ShouldProcess($url, 'Install or update Smart Storage Analyzer')) {
            Connect-PnPOnline -Url $url -Interactive -ClientId $ClientId
            $app = Get-PnPApp -Identity $AppId -Scope Tenant
            if (-not $app) {
                $status = 'Error: app not found in tenant app catalog'
            } elseif ($app.InstalledVersion) {
                if ($Update -and ([version]$app.InstalledVersion -lt [version]$app.AppCatalogVersion)) {
                    Update-PnPApp -Identity $AppId -Scope Tenant
                    $status = "Updated to $($app.AppCatalogVersion)"
                } else {
                    $status = "Skipped (already installed: $($app.InstalledVersion))"
                }
            } else {
                Install-PnPApp -Identity $AppId -Scope Tenant -Wait
                $status = 'Installed'
            }
        } else {
            $status = 'WhatIf (no change made)'
        }
    } catch {
        $status = "Error: $($_.Exception.Message)"
    }
    Write-Host ("{0,-70} {1}" -f $url, $status)
    [pscustomobject]@{ Url = $url; Status = $status }
}

$logPath = Join-Path (Get-Location) ("StorageAnalyzer-install-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
$results | Export-Csv -Path $logPath -NoTypeInformation
Write-Host "Done. Results written to $logPath"
