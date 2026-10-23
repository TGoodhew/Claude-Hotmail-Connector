<#
  register-app.ps1 — create YOUR OWN Microsoft Entra app registration for the connector.

  Each user brings their own registration (spec v2 section 8.1); the connector does not ship a shared
  one. This turns the README's portal walkthrough ("Register a Microsoft Entra application")
  into one command. The app is a PUBLIC / native client: it authenticates with PKCE and NO
  client secret, so the resulting "Application (client) ID" is not a secret.

  Requirements:
    - A Microsoft Entra tenant. A personal Microsoft account (hotmail/outlook/live) has none by
      default: create a free one first (typically by signing up for a free Azure account), see
      https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-create-new-tenant
    - Azure CLI (`az`) signed in to that tenant:  az login
    - Permission to create app registrations in that tenant (you have it in your own tenant).

  Usage:
    pwsh ./scripts/register-app.ps1 [-DisplayName "Claude Hotmail Connector"] [-Audience Personal|Multitenant]

  -Audience Personal (default): personal Microsoft accounts only (hotmail/outlook/live), signed in via
  the `consumers` endpoint, the connector's default. Work/school accounts are already served by
  Claude's official Microsoft 365 connector.
  -Audience Multitenant: also allow work/school accounts; then set MICROSOFT_TENANT=common.
#>
[CmdletBinding()]
param(
  [string]$DisplayName = "Claude Hotmail Connector",
  [ValidateSet("Personal", "Multitenant")]
  [string]$Audience = "Personal"
)

$signInAudience = if ($Audience -eq "Personal") { "PersonalMicrosoftAccount" } else { "AzureADandPersonalMicrosoftAccount" }

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
# -InputObject, not a pipe: piping a one-element array unwraps it, and az then
# silently ignores the object it gets instead of an array (no permissions set).
ConvertTo-Json -InputObject $rra -Depth 6 | Set-Content -LiteralPath $rraFile -Encoding utf8

Write-Host "Creating public-client app registration '$DisplayName' ($Audience accounts)..." -ForegroundColor Cyan
# Personal-account apps require v2 access tokens; harmless for multitenant ones.
$app = az ad app create `
  --display-name $DisplayName `
  --sign-in-audience $signInAudience `
  --requested-access-token-version 2 `
  --is-fallback-public-client true `
  --public-client-redirect-uris "http://localhost" `
  --required-resource-accesses "@$rraFile" | ConvertFrom-Json

$created = $LASTEXITCODE
Remove-Item -LiteralPath $rraFile -Force -ErrorAction SilentlyContinue
if ($created -ne 0 -or -not $app.appId) {
  throw "az ad app create failed (exit code $created); see the error above."
}
# Check the Graph permissions landed: az doesn't fail when it ignores them.
$granted = @($app.requiredResourceAccess | Where-Object { $_.resourceAppId -eq $graphAppId } |
  ForEach-Object { $_.resourceAccess })
if ($granted.Count -ne $scopes.Count) {
  throw "The app '$($app.appId)' was created without its $($scopes.Count) Microsoft Graph " +
    "permissions. Add them under API permissions, or delete the app and run this script again."
}

Write-Host ""
Write-Host "Done. Application (client) id:" -ForegroundColor Green
Write-Host "  $($app.appId)"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Give the connector this id, either:"
Write-Host "       - Claude Desktop: install hotmail-connector.mcpb and paste it into the extension's"
Write-Host "         'Application (client) ID' setting, or"
Write-Host "       - Claude Code / source: set MICROSOFT_CLIENT_ID=$($app.appId) in the server's env."
Write-Host "  2. Ask Claude something that needs your mailbox; a Microsoft sign-in opens in your browser."
Write-Host "  3. If sign-in fails, confirm under Authentication: 'Allow public client flows' = Yes,"
Write-Host "     and the 'Mobile and desktop applications' redirect URI is http://localhost."
if ($Audience -eq "Multitenant") {
  Write-Host "  4. Multitenant registration: set MICROSOFT_TENANT=common so work/school accounts can sign in."
}
