<#
.SYNOPSIS
    Reports Restricted Content Discovery (RCD) status for all SharePoint Online sites.

.DESCRIPTION
    Enumerates every site in the tenant, reads the RestrictContentOrgWideSearch
    property (the site-level "Restricted Content Discovery" setting), and exports
    a CSV that can be filtered on sites where it is Enabled vs Not Enabled.

    Reference: https://learn.microsoft.com/en-us/sharepoint/restricted-content-discovery

    Notes:
    - Restricted Content Discovery is only supported on SharePoint sites, so
      OneDrive personal sites are excluded by default.
    - Requires the SharePoint Online Management Shell module and a SharePoint
      Administrator (or Global Administrator) account.
    - Per Microsoft docs, enabling/disabling RCD requires SharePoint Advanced
      Management + a Microsoft 365 Copilot license, but READING the property
      works regardless.

.PARAMETER AdminUrl
    SharePoint admin center URL, e.g. https://contoso-admin.sharepoint.com
    If omitted, you will be prompted for the tenant name.

.PARAMETER OutputPath
    Full path for the CSV output. Defaults to
    .\RestrictedContentDiscovery-Report-<yyyyMMdd-HHmm>.csv

.PARAMETER Fast
    Skips the per-site re-query and trusts the property values returned by the
    bulk 'Get-SPOSite -Limit All' call. Much faster on large tenants, but some
    properties can be stale/null in bulk mode. Without -Fast each site is
    re-read individually for accuracy.

.PARAMETER IncludeOneDrive
    Also enumerate OneDrive personal sites. RCD cannot be applied to OneDrive,
    so these will always report as not enabled - included only for completeness.

.PARAMETER UseSystemBrowser
    Hand authentication off to the system default browser instead of WAM.
    MFA-friendly and works in PowerShell 7 - this is the recommended auth method:
      Connect-SPOService -Url <admin> -ModernAuth $true `
        -AuthenticationUrl https://login.microsoftonline.com/organizations -UseSystemBrowser $true

.EXAMPLE
    .\Get-RestrictedContentDiscoveryReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Get-RestrictedContentDiscoveryReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com -Fast -OutputPath C:\Reports\RCD.csv
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$AdminUrl,

    [Parameter()]
    [string]$OutputPath = ".\RestrictedContentDiscovery-Report-$(Get-Date -Format 'yyyyMMdd-HHmm').csv",

    [Parameter()]
    [switch]$Fast,

    [Parameter()]
    [switch]$IncludeOneDrive,

    [Parameter()]
    [switch]$UseSystemBrowser
)

# ---------------------------------------------------------------------------
# 1. Prerequisites
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
    Write-Error "The SharePoint Online Management Shell module is not installed. Install it with:`n    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
    return
}
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop

if (-not $AdminUrl) {
    $tenant = Read-Host "Enter tenant name (e.g. 'contoso' for contoso-admin.sharepoint.com)"
    $AdminUrl = "https://$tenant-admin.sharepoint.com"
}

# ---------------------------------------------------------------------------
# 2. Connect
# ---------------------------------------------------------------------------
Write-Host "Connecting to $AdminUrl ..." -ForegroundColor Cyan
$connectParams = @{ Url = $AdminUrl; ErrorAction = 'Stop' }
$availableParams = (Get-Command Connect-SPOService).Parameters.Keys
if ($UseSystemBrowser) {
    if ($availableParams -contains 'UseSystemBrowser') {
        # Browser-based auth: MFA happens in the browser, not in PowerShell
        $connectParams['UseSystemBrowser'] = $true
        if ($availableParams -contains 'ModernAuth')        { $connectParams['ModernAuth'] = $true }
        if ($availableParams -contains 'AuthenticationUrl') { $connectParams['AuthenticationUrl'] = 'https://login.microsoftonline.com/organizations' }
    }
    else {
        Write-Warning "'-UseSystemBrowser' is not supported by this module version. Run: Update-Module Microsoft.Online.SharePoint.PowerShell"
    }
}
try {
    Connect-SPOService @connectParams
}
catch {
    Write-Error "Failed to connect to SharePoint Online: $($_.Exception.Message)"
    return
}

# ---------------------------------------------------------------------------
# 3. Enumerate sites
# ---------------------------------------------------------------------------
Write-Host "Retrieving site list (this can take a while on large tenants)..." -ForegroundColor Cyan
$getParams = @{ Limit = 'All' }
if ($IncludeOneDrive) { $getParams['IncludePersonalSite'] = $true }

try {
    $sites = Get-SPOSite @getParams -ErrorAction Stop
}
catch {
    Write-Error "Failed to enumerate sites: $($_.Exception.Message)"
    return
}

Write-Host ("Found {0} site(s). Checking Restricted Content Discovery status..." -f $sites.Count) -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 4. Read the RestrictContentOrgWideSearch property per site
# ---------------------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()
$errors  = [System.Collections.Generic.List[object]]::new()
$i = 0

foreach ($site in $sites) {
    $i++
    Write-Progress -Activity "Checking Restricted Content Discovery" `
                   -Status "$i of $($sites.Count): $($site.Url)" `
                   -PercentComplete (($i / $sites.Count) * 100)

    try {
        if ($Fast) {
            # Use the property from the bulk call as-is
            $rcdValue = $site.RestrictContentOrgWideSearch
        }
        else {
            # Re-read the individual site for an accurate property value
            $detail = Get-SPOSite -Identity $site.Url -ErrorAction Stop
            $rcdValue = $detail.RestrictContentOrgWideSearch
        }

        $results.Add([PSCustomObject]@{
            SiteUrl                     = $site.Url
            Title                       = $site.Title
            Template                    = $site.Template
            Owner                       = $site.Owner
            RestrictedContentDiscovery  = if ($rcdValue -eq $true) { 'Enabled' } else { 'Not Enabled' }
            IsOneDrive                  = ($site.Template -like 'SPSPERS*' -or $site.Url -like '*-my.sharepoint.com/personal/*')
            StorageUsedMB               = [math]::Round($site.StorageUsageCurrent, 0)
            LastContentModified         = $site.LastContentModifiedDate
        })
    }
    catch {
        $errors.Add([PSCustomObject]@{
            SiteUrl = $site.Url
            Error   = $_.Exception.Message
        })
    }
}
Write-Progress -Activity "Checking Restricted Content Discovery" -Completed

# ---------------------------------------------------------------------------
# 5. Export + summary
# ---------------------------------------------------------------------------
$results | Sort-Object RestrictedContentDiscovery, SiteUrl | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$enabledCount = ($results | Where-Object { $_.RestrictedContentDiscovery -eq 'Enabled' }).Count

Write-Host "`n================= SUMMARY =================" -ForegroundColor Green
Write-Host ("Total sites checked : {0}" -f $results.Count)
Write-Host ("RCD Enabled         : {0}" -f $enabledCount) -ForegroundColor ($enabledCount -gt 0 ? 'Yellow' : 'Green')
Write-Host ("RCD Not Enabled     : {0}" -f ($results.Count - $enabledCount))
if ($errors.Count -gt 0) {
    Write-Host ("Errors              : {0} (see below)" -f $errors.Count) -ForegroundColor Red
    $errors | Format-Table -AutoSize
}
Write-Host "==========================================="
Write-Host "Report saved to: $(Resolve-Path $OutputPath)" -ForegroundColor Cyan
Write-Host "Tip: filter the 'RestrictedContentDiscovery' column in Excel to split Enabled vs Not Enabled."
