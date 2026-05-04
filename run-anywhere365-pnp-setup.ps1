#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Tenant,
    [string]$SiteUrl,
    [string]$ApplicationName = "Anywhere365-OneUCC-PnPApp",
    [string]$CertificateAppId,
    [switch]$SkipFullControl
)

$ErrorActionPreference = 'Stop'

function Read-RequiredValue {
    param(
        [Parameter(Mandatory)] [string]$Prompt,
        [string]$CurrentValue
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return $CurrentValue
    }

    do {
        $value = Read-Host $Prompt
    } while ([string]::IsNullOrWhiteSpace($value))

    return $value.Trim()
}

Write-Host "=== Anywhere365 PnP setup gestart ===" -ForegroundColor Cyan

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell module niet gevonden. Installeer met: Install-Module PnP.PowerShell -Scope CurrentUser"
}

$Tenant = Read-RequiredValue -Prompt "Voer tenant in (bijv. contoso.onmicrosoft.com of tenant GUID)" -CurrentValue $Tenant
$SiteUrl = Read-RequiredValue -Prompt "Voer SharePoint site URL in (bijv. https://contoso.sharepoint.com/sites/Anywhere365)" -CurrentValue $SiteUrl
$CertificateAppId = Read-RequiredValue -Prompt "Voer App ID met certificaat in" -CurrentValue $CertificateAppId

Write-Host "\n[1/5] Registreren van PnP Entra ID app: $ApplicationName" -ForegroundColor Yellow
$registration = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $ApplicationName -Tenant $Tenant -Interactive

$clientId = $null
if ($registration -and $registration.PSObject.Properties.Name -contains 'AppId') {
    $clientId = $registration.AppId
}
if ([string]::IsNullOrWhiteSpace($clientId) -and $registration -and $registration.PSObject.Properties.Name -contains 'ClientId') {
    $clientId = $registration.ClientId
}
if ([string]::IsNullOrWhiteSpace($clientId)) {
    $clientId = Read-Host "Kon ClientId niet automatisch bepalen. Vul ClientId handmatig in"
}

Write-Host "\n[2/5] Verbinden met SharePoint via Connect-PnPOnline" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $clientId

Write-Host "\n[3/5] Huidige app-permissies opvragen voor site" -ForegroundColor Yellow
Get-PnPAzureADAppSitePermission -Site $SiteUrl | Format-Table -AutoSize

Write-Host "\n[4/5] Write permissie toekennen aan app met certificaat" -ForegroundColor Yellow
$grantRights = Grant-PnPAzureADAppSitePermission -AppId $CertificateAppId -DisplayName "Anywhere365AppOnly" -Permissions Write -Site $SiteUrl

if (-not $SkipFullControl) {
    Write-Host "\n[5/5] Permissie verhogen naar FullControl" -ForegroundColor Yellow
    Set-PnPAzureADAppSitePermission -PermissionId $grantRights.Id -Permissions FullControl -Site $SiteUrl
} else {
    Write-Host "\n[5/5] Overslaan van FullControl (SkipFullControl is gezet)." -ForegroundColor DarkYellow
}

Write-Host "\nKlaar. Samenvatting:" -ForegroundColor Green
Write-Host "Tenant: $Tenant"
Write-Host "SiteUrl: $SiteUrl"
Write-Host "ClientId: $clientId"
Write-Host "CertificateAppId: $CertificateAppId"
Write-Host "PermissionId: $($grantRights.Id)"
