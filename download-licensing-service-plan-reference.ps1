#Use: Downloading licensing-service-plan-reference for later use.
#Version: 0.1
#Author: Simon Ingildsen - simon-ing@hotmail.com
#Maintainer: Simon Ingildsen - simon-ing@hotmail.com

#Load Modules
$ErrorOccured = $false
$modules = @("ImportExcel")

foreach ($module in $modules) {

try {import-module $module -ErrorAction Stop}
catch {$ErrorOccured = $true}

if (!$ErrorOccured) {
    write-host "PS module"$module" imported"
} else {
    write-host "Trying to install PS module"$module
    try {install-module $module -ErrorAction Stop}
    catch {write-host "Failed to install PS module"$module", are you running this as administrator?"}
}
}

# Variables

$url = "https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference"

# Work

$htmlTable = Get-HtmlTable -Url $url


# Fixing Microsoft mistakes

# 1. https://docs.microsoft.com/en-us/microsoft-365/security/office-365-security/office-365-atp?view=o365-worldwide#microsoft-defender-for-office-365-plan-1-and-plan-2
# Microsoft Defender for Office 365 Plan 1 is included in Microsoft 365 Business Premium.

$Serviceplansincluded = ([Environment]::NewLine) + "ATP_ENTERPRISE (f20fedf3-f3c3-43c3-8267-2bfdd51c0939)"
$ServiceplansincludedFriendlyNames = ([Environment]::NewLine) + "Office 365 Advanced Threat Protection (Plan 1) (f20fedf3-f3c3-43c3-8267-2bfdd51c0939)"

($htmlTable | Where-Object {$_.Productname -eq "Microsoft 365 Business Premium"}).Serviceplansincluded += $Serviceplansincluded
($htmlTable | Where-Object {$_.Productname -eq "Microsoft 365 Business Premium"})."Serviceplansincluded(friendlynames)" += $ServiceplansincludedFriendlyNames


# Export

$htmlTable | Export-Csv -Path .\licensing-service-plan-reference-modifed.csv
$htmlTable | Export-Excel -Path .\licensing-service-plan-reference-modifed.xlsx