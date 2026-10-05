#Requires -Version 7.4
#Requires -Modules PnP.PowerShell
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
 [Parameter(Mandatory)][ValidatePattern('^https://[a-z0-9-]+\.sharepoint\.com/sites/DORK$')][string]$SiteUrl,
 [Parameter(Mandatory)][guid]$ProvisioningClientId,
 [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference='Stop'
# Explicitly require an existing dedicated site; no tenant-wide grants or site creation here.
Connect-PnPOnline -Url $SiteUrl -ClientId $ProvisioningClientId -Interactive
$documents=Get-PnPList -Identity 'Documents' -Includes RootFolder,ItemCount
if($documents.RootFolder.ServerRelativeUrl -ne '/sites/DORK/Shared Documents'){throw 'Documents library must use /sites/DORK/Shared Documents.'}
if($documents.ItemCount -gt 0){throw 'Preview provisioning refuses populated Documents libraries.'}
if(-not $PSCmdlet.ShouldProcess($SiteUrl,'Create DORK libraries, lists and metadata columns')){return}
function Ensure-List($Title,$Url,$Template){
 $existing=Get-PnPList | Where-Object Title -eq $Title
 if(-not $existing){New-PnPList -Title $Title -Url $Url -Template $Template -OnQuickLaunch | Out-Null}
 Get-PnPList -Identity $Title
}
function Ensure-Field($List,$Name,$Internal,$Type,$Choices){
 $existing=Get-PnPField -List $List | Where-Object InternalName -eq $Internal
 if($existing){if($existing.TypeAsString -ne $Type){throw "Incompatible existing field $Internal"};return}
 $args=@{List=$List;DisplayName=$Name;InternalName=$Internal;Type=$Type;AddToDefaultView=$true}
 if($Choices){$args.Choices=$Choices}
 Add-PnPField @args | Out-Null
}
$templates=Ensure-List 'DORK-Templates' 'DORK-Templates' 'DocumentLibrary'
$registry=Ensure-List 'DORK ID Registry' 'Lists/DORKIDRegistry' 'GenericList'
$needs=Ensure-List 'Documentation Needs' 'Lists/DocumentationNeeds' 'GenericList'
foreach($name in @('DocumentKey','DorkID','CurrentLocation')){Ensure-Field $registry.Id $name $name 'Text'}
Ensure-Field $registry.Id 'AllocationState' 'AllocationState' 'Choice' @('Reserved','Applied')
foreach($name in @('DocumentKey','DorkID')){Set-PnPField -List $registry.Id -Identity $name -Values @{Indexed=$true;EnforceUniqueValues=$true} | Out-Null}
Ensure-Field $documents.Id 'DORK ID' 'DORK_x0020_ID' 'Text'
Ensure-Field $documents.Id 'Information Source' 'Information_x0020_Source' 'Text'
foreach($name in @('Owner','Responsible','Secondary')){Ensure-Field $documents.Id $name $name 'User'}
foreach($pair in @(@('Last Reviewed','Last_x0020_Reviewed'),@('Next Review','Next_x0020_Review'))){Ensure-Field $documents.Id $pair[0] $pair[1] 'DateTime'}
Ensure-Field $documents.Id 'Classification' 'Classification' 'Choice' @('Public','Internal','Confidential','Restricted')
Ensure-Field $documents.Id 'Audience' 'Audience' 'Choice' @('IT Team','Staff','Administrators','Public')
Ensure-Field $documents.Id 'Lifecycle' 'Lifecycle' 'Choice' @('Draft','Active','Superseded','Retired')
Ensure-Field $needs.Id 'Details' 'Details' 'Note'
Ensure-Field $needs.Id 'Need Type' 'NeedType' 'Choice' @('New Document','Update Document','Review Document')
Ensure-Field $needs.Id 'Priority' 'Priority' 'Choice' @('Low','Normal','High','Critical')
Ensure-Field $needs.Id 'Status' 'Status' 'Choice' @('New','Planned','In Progress','Completed','Cancelled')
Ensure-Field $needs.Id 'Assigned To' 'AssignedTo' 'User'
Ensure-Field $needs.Id 'Target Date' 'TargetDate' 'DateTime'
@{siteUrl=$SiteUrl;documentsListId=$documents.Id.ToString();registryListId=$registry.Id.ToString();templatesListId=$templates.Id.ToString();needsListId=$needs.Id.ToString();complete=$false;remaining=@('Taxonomy and managed metadata','Content types and template bindings','Register views','Dashboard page','Entra/Azure setup','Framework bootstrap')} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutputPath -Encoding utf8
Write-Warning 'Schema scaffold created. Installation is incomplete; follow START-HERE-CHATGPT.md.'
