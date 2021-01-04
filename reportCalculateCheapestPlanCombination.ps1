#Use: Returns the cheapest MS365 product combination for the selected license group
#Version: 0.1
#Author: Simon Ingildsen - simon-ing@hotmail.com
#Maintainer: Simon Ingildsen - simon-ing@hotmail.com

$runTime = Get-Date

# Loading needed modules

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

# Work

$licensingServicePlanReference = Import-Csv -path .\licensing-service-plan-reference-modifed.csv

$prices = Import-Excel -Path .\ms365prices.xlsx

$pricesWithoutMonthly = $prices | Where-Object {$_.Productname -notlike "*Month to Month*"}

$productsFound = @()

Foreach ($price in $pricesWithoutMonthly ) {
    $priceReference = $licensingServicePlanReference | Where-Object {$_.Productname -eq $price.Productname}
    if ($priceReference) {
        $productsFound += $price.Productname
    } else {
    }
}

Write-Host $productsFound.count"of"$pricesWithoutMonthly.count"products from the price list was found in microsoft licensing-service-plan-reference.csv"


# Crating a array with PSObejcts, with all the information needed to continue

$productsData = @()

ForEach ($productFound in $productsFound) {

    [decimal]$productPrice = ($pricesWithoutMonthly | Where-Object {$_.Productname -eq $productFound})."Price DKK user/month"

    $productFoundServices = ($licensingServicePlanReference | Where-Object {$_.Productname -eq $productFound}).Serviceplansincluded

    $productFoundServicesSplit = $productFoundServices.Split([Environment]::NewLine) | Where-Object {$_ -ne ""}

    $productFoundServicesName = @()

    ForEach ($productFoundService in $productFoundServicesSplit) {
        $productFoundServicesName += $productFoundService.Split(" ")[0]
    }

    $productDetails = [pscustomobject]@{
        productname = $productFound
        price = $productPrice
        services = $productFoundServicesName 

    }

    $productsData += $productDetails

}


# whatDoYouWant - Variables

$pieces = 2000
$microsoft365AppsForAny = $true # Yes, No
$microsoft365AppsForBusiness = $false # Yes, No
$microsoft365AppsForEnterprise = $false # Yes, No
$intune = $true # Yes, No
$ExchangeAny = $true # Yes, No
$ExchangeOnlinePlan1 = $false # Yes, No
$ExchangeOnlinePlan2 = $false # Yes, No
$aadP1 = $true # Yes, No
$office365ATP = $false # Yes, No

# Work

# Creating Key Value Pair for variables

$whatDoYouWant = @{
    'microsoft365AppsForAny' = $microsoft365AppsForAny 
    'microsoft365AppsForBusiness' = $microsoft365AppsForBusiness
    'microsoft365AppsForEnterprise' = $microsoft365AppsForEnterprise
    'intune' = $intune
    'ExchangeAny' = $ExchangeAny
    'ExchangeOnlinePlan1' = $ExchangeOnlinePlan1
    'ExchangeOnlinePlan2' = $ExchangeOnlinePlan2
    'aadP1' = $aadP1
    'office365ATP' = $office365ATP
}

# Creating Key Value Pair for service and servicesku matching

$servicesMatching = @{
    'intune' = ("INTUNE_A")
    'microsoft365AppsForAny' = ("OFFICE_BUSINESS","OFFICESUBSCRIPTION")
    'microsoft365AppsForBusiness' = ("OFFICE_BUSINESS")
    'microsoft365AppsForEnterprise' = ("OFFICESUBSCRIPTION")
    'ExchangeAny' = ("EXCHANGE_S_STANDARD","EXCHANGE_S_ENTERPRISE")
    'ExchangeOnlinePlan1' = ("EXCHANGE_S_STANDARD")
    'ExchangeOnlinePlan2' = ("EXCHANGE_S_ENTERPRISE")
    'aadP1' = ("AAD_PREMIUM","AAD_SMB")
    'office365ATP' = ("ATP_ENTERPRISE")
}

# Creating Key Value Pair for produkt user limitations
# Source https://docs.microsoft.com/en-us/office365/servicedescriptions/office-365-platform-service-description/office-365-plan-options

$plansWithLimits = @{
    'Microsoft 365 Business Basic' = 300
    'Microsoft 365 Business Standard' = 300
    'Microsoft 365 Business Premium' = 300
    'Microsoft 365 Apps for business' = 300
}

# Creating array with needed services

ForEach ($productname in $productsData) {

    $licenseNeededData = @()

    ForEach ($option in $whatDoYouWant.keys) {
        if ($whatDoYouWant[$option]) {
            $licenseNeededDetails = [PSCustomObject]@{
                Name = $servicesMatching[$option]
            }
            $licenseNeededData += $licenseNeededDetails 
        }
    }
}

$powersetData = Import-Csv -Path .\powersetDataSumLT402.csv
$powersetData = $powersetData | Select-Object Name,@{ label = "Sum";  Expression = {[decimal]($_.Sum)}},Services

# Creating a array which only contains product combinations with all or equivalent services present

$leftOvers = $powersetData

ForEach ($licItem in $licenseNeededData) {
    $licItemName = $licItem.name
    if ($licItem.name.count -gt 1) {
        $command = '$leftOvers = $leftOvers | Where-Object {'
        $i = 0
        ForEach ($name in $licItem.name) {
            $i = $i + 1
            $like = '$_.Services -like "*'+$name+'*"'
            if ($i -lt $licItem.name.count) {
                $or = " -or "
                $like = $like + $or
            }
            $command = $command + $like
        }
        $command = $command + '}'
        invoke-expression -command $command
    } else {
        $leftOvers = $leftOvers | Where-Object {$_.Services -like "*$licItemName*" }
    }
}

# Splits users into groups if plans with limits is used.

$resultData = @()

$limitPlansUsed = @()

$piecesLeft = $pieces

DO
{
    ForEach ($plan in $limitPlansUsed) {
        $leftOvers = $leftOvers | Where-Object {$_.Name -notlike "*$plan*" }
    }

    $cheapestOption = $leftOvers | Sort-Object -Property Sum | Select-Object -first 1
    $cheapestOptionName = $cheapestOption.Name.Split(",")

    $containsPlansWithLimits = $false

    ForEach ($productname in $cheapestOptionName) {
        if ($plansWithLimits[$productname]) {
            $containsPlansWithLimits = $true
            If ($piecesLeft -gt $plansWithLimits[$productname]) {
                $limitPlansUsed += $productname
                $piecesLeft = $piecesLeft - $plansWithLimits[$productname]
                $piecesUsed = $plansWithLimits[$productname]
            } else {
                $piecesUsed = $piecesLeft
                $piecesLeft = ""
            }
        } else {

        }
    }

    if (!($containsPlansWithLimits)) {
        $piecesUsed = $piecesLeft
        $piecesLeft = ""
    }

    $resultDetailsName = "Group"+ ($resultData.count + 1)
    $total = $piecesUsed * $cheapestOption.Sum

    $resultDetails = [PSCustomObject]@{
        Name        = $resultDetailsName
        Combination = $cheapestOption.Name
        Pieces      = $piecesUsed
        containsPlansWithLimits = $containsPlansWithLimits
        Price = $cheapestOption.Sum
        Total = $total
    }
    $resultData += $resultDetails
} While ($piecesLeft -gt 1)


# Display result to end user

$resultData | Format-Table | Out-String | Write-Host

[decimal]$totalMonthlySum = 0

ForEach ($result in $resultData) {
    $totalMonthlySum = $totalMonthlySum + $result.Total
}

$totalYearlySum = $totalMonthlySum * 12

Write-Host "For"$pieces" users this is "$totalMonthlySum" DKK per month and "$totalYearlySum" DKK per year"


# Time info
$endTime = Get-Date
$time = ($endTime - $runTime).TotalSeconds

Write-Host "This script took"$time" seconds to run."
