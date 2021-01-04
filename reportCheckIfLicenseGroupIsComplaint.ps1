#Use: MS365ComplainsCAL. Check if your Microsoft 365 Suite is compaint.
#Version: 0.1
#Author: Simon Ingildsen - simon-ing@hotmail.com
#Maintainer: Simon Ingildsen - simon-ing@hotmail.com

$licensingServicePlanReference = Import-Csv -path .\licensing-service-plan-reference-modifed.csv

# Service Active?

# Compliance impractical or not possible

$complianceImpractical = @{
    'dataLossPrevention' = $true # Suite E3
    'legaleDiscoveryHold' = $true # Suite E3
    'eDiscoveryexport' = $true # Suite E3
    'office365AdvancedCompliance' = $false # Suite E5
    'aIPPlan2AddsAutomaticdataclassficationAIPScanner' = $false # Suite E5
    'azureADPremiumP1' = $true # Suite E3
    'azureADPremiumP2AddsPrivIDmanagementPIMIdentityProtectionAccessReviews' = $false # Suite E5
    'azureAdvancedTreatProtection' = $false # Suite E5
    'office365AdvancedThreatProtectionP2' = $false # Suite E5
    'microsoftDefenderAdvancedThreatProtection' = $false # Suite E5
}

# Compliance possible manuel configuration
$compliancePossible = @{
    'mSCloudAppSecurity' = $false # Suite E5
    'office365AdvancedThreatProtectionP1' = $true # Suite E5
}

$servicesNeeded = @{
    'dataLossPrevention' = ("EXCHANGE_S_ENTERPRISE","EXCHANGE_S_ARCHIVE","EXCHANGE_S_ARCHIVE_ADDON")
    'LegaleDiscoveryHold' = ("EXCHANGE_S_ENTERPRISE","EXCHANGE_S_ARCHIVE","EXCHANGE_S_ARCHIVE_ADDON")
    'eDiscoveryexport' = ("EXCHANGE_S_ENTERPRISE","EXCHANGE_S_ARCHIVE","EXCHANGE_S_ARCHIVE_ADDON")
    'office365AdvancedCompliance' = ('EQUIVIO_ANALYTICS') # Is this right license?
    'aIPPlan2AddsAutomaticdataclassficationAIPScanner' = ('RMS_S_PREMIUM2') # Is this right license?
    'azureADPremiumP1' = ("AAD_PREMIUM")
    'azureADPremiumP2AddsPrivIDmanagementPIMIdentityProtectionAccessReviews' = ('AAD_PREMIUM_P2')
    'azureAdvancedTreatProtection' = ('ATA')
    'microsoftDefenderAdvancedThreatProtection' = ('WINDEFATP')
    'mSCloudAppSecurity' = ('ADALLOM_S_STANDALONE')
    'office365AdvancedThreatProtectionP1' = ("ATP_ENTERPRISE")
    'office365AdvancedThreatProtectionP2' = ('THREAT_INTELLIGENCE')
}

$servicesEnabled = @()

Foreach ($complianceImpracticalService in $complianceImpractical.keys) {
    if ($complianceImpractical[$complianceImpracticalService]) {
        $servicesEnabled += $complianceImpracticalService
    }
}

Foreach ($compliancePossibleService in $compliancePossible.keys) {
    if ($compliancePossible[$compliancePossibleService]) {
        $servicesEnabled += $compliancePossibleService
    }
}


# Current license groups and pieces

$groupsPieces = @{
	'group1' = 114
	'group2' = 20
}

$groupsSkus = @{
	'group1' = ("Microsoft 365 E3","Office 365 Advanced Threat Protection (Plan 1)")
	'group2' = ("EXCHANGE ONLINE (PLAN 1)","Office 365 Advanced Threat Protection (Plan 1)","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE","AZURE ACTIVE DIRECTORY PREMIUM P1")
}

#

$groupsData = @()

foreach ($group in $groupsPieces.keys) {

    $getGroupsSkus = $groupsSkus[$group]

    $services = @()

    foreach ($groupsSku in $getGroupsSkus) {
        $productnameDetails = $licensingServicePlanReference | Where-Object {$_.Productname -eq $groupsSku}
        $services += $productnameDetails.Serviceplansincluded.Split([Environment]::NewLine) | Where-Object {$_ -ne ""}
    }

    $serviceNames = @()

    Foreach ($serviceName in $services) {
        $serviceNames += $serviceName.Split(" ")[0]
    }

    $servicesUnique = $serviceNames | Select-Object -Unique

    $groupsDetails = [pscustomobject]@{
        groupName = $group
        pieces = $groupsPieces[$group]
        skus = $groupsSkus[$group]
        services = $servicesUnique
        servicesCount = $servicesUnique.count 
    }
        $groupsData += $groupsDetails

}
$groupsData | ft

# Checking groups

Foreach ($groupData in $groupsData) {
    Write-Host $groupData.groupName

    Foreach ($serviceEnabled in $servicesEnabled) {
        $serviceComplaint = $false

        Foreach ($serviceNeeded in $servicesNeeded[$serviceEnabled]) {
            $serviceFound = $groupData.services | Where-Object {$_ -eq $serviceNeeded}
            if ($serviceFound) {
                $serviceComplaint = $true
            } else {
            }
        }

        if ($serviceComplaint) {
            Write-Host $serviceEnabled" is complaint"
        } else {
            Write-Host $serviceEnabled" is not complaint"
        }
    }
}


