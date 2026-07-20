<#
  register-app.ps1 — ONE-TIME publisher setup (run by the maintainer, not end users).

  Creates the shared Microsoft Entra app registration the build ships with. The app is a
  PUBLIC / native client: it authenticates with PKCE and NO client secret, so the resulting
  "Application (client) id" is NOT a secret and is safe to embed in the build
  (src/config.ts -> DEFAULT_CLIENT_ID) and distribute.

  End users NEVER run this — a personal Microsoft account has no tenant to register an app in.
  This turns the ~8-screen portal walkthrough (docs/installer.md) into one command.

  Requirements:
    - Azure CLI (`az`) signed in to a tenant you own:  az login
    - Permission to create app registrations in that tenant.

  Usage:
    pwsh ./scripts/register-app.ps1 [-DisplayName "Claude Hotmail Connector"]
#>
[CmdletBinding()]
param(
  [string]$DisplayName = "Claude Hotmail Connector"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
  throw "Azure CLI (az) not found. Install it and run 'az login' first."
}

# Microsoft Graph resource app id + well-known DELEGATED permission (scope) ids.
$graphAppId = "00000003-0000-0000-c000-000000000000"
$scopes = @(
  "37f7f235-527c-4136-accd-4a02d197296e", # openid
  "14dad69e-099b-42c9-810b-d002981feec1", # profile
  "64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0", # email
  "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", # offline_access
  "e1fe6dd8-ba31-4d61-89e7-88639da4683d", # User.Read
  "570282fd-fa5c-430d-a7fd-fc8dc98a9dca", # Mail.Read
  "1ec239c2-d7c9-4623-a91a-a9775856bb36"  # Calendars.ReadWrite
)

$rra = @(
  @{
    resourceAppId  = $graphAppId
    resourceAccess = @($scopes | ForEach-Object { @{ id = $_; type = "Scope" } })
  }
)
$rraFile = New-TemporaryFile
$rra | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $rraFile -Encoding utf8

Write-Host "Creating public-client app registration '$DisplayName'..." -ForegroundColor Cyan
$app = az ad app create `
  --display-name $DisplayName `
  --sign-in-audience AzureADandPersonalMicrosoftAccount `
  --is-fallback-public-client true `
  --public-client-redirect-uris "http://localhost" `
  --required-resource-accesses "@$rraFile" | ConvertFrom-Json

Remove-Item -LiteralPath $rraFile -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done. Application (client) id:" -ForegroundColor Green
Write-Host "  $($app.appId)"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Put this id in src/config.ts -> DEFAULT_CLIENT_ID (and/or your .env)."
Write-Host "  2. (Recommended before wide distribution) complete Microsoft Publisher Verification"
Write-Host "     and set consent-screen branding (logo, privacy/ToS URLs) so users don't see the"
Write-Host "     'unverified app' warning at sign-in."
Write-Host "  3. Confirm under Authentication: 'Allow public client flows' = Yes, redirect http://localhost."
