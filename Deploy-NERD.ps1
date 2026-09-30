$ErrorActionPreference = 'Continue'
Set-Location $PSScriptRoot
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'Node.js is required. Use the Node installation already used for NERD.' }
$siteOrigin = Read-Host 'Paste the Azure Static Web App URL (https://...azurestaticapps.net)'
& node scripts/create-manifest.cjs $siteOrigin
if ($LASTEXITCODE -ne 0) { throw 'Manifest creation failed.' }
$secureToken = Read-Host 'Paste the deployment token from Azure (hidden input)' -AsSecureString
$tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
$previousToken = $env:SWA_CLI_DEPLOYMENT_TOKEN
try {
  $env:SWA_CLI_DEPLOYMENT_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer)
  & npx --yes --package '@azure/static-web-apps-cli' swa deploy ./dist --api-location ./api --api-language node --api-version 22 --env production --verbose info 2>&1 | ForEach-Object { ([string]$_).Replace($env:SWA_CLI_DEPLOYMENT_TOKEN, "[REDACTED]") }
  if ($LASTEXITCODE -ne 0) { throw 'Azure deployment failed. Save the error message, excluding any token.' }
} finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
  $env:SWA_CLI_DEPLOYMENT_TOKEN = $previousToken
  $secureToken.Dispose()
}
Write-Host 'Deployment completed. manifest-azure.xml is ready for the Office add-in deployment step.'
