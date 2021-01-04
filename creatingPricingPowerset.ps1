#Use: Creating PowerSet from the price list.
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

# Loading needed functions

# Functions
# https://github.com/dfinke/powershell-algorithms/tree/master/src/algorithms/sets/power-set

function powerSet($originalSet) {
    $subSets = @()

    # We will have 2^n possible combinations (where n is a length of original set).
    # It is because for every element of original set we will decide whether to include
    # it or not (2 options for each set element).

    #$numberOfCombinations = [Math]::Pow(2, $originalSet.Count)
    $numberOfCombinations = 1 -shl $originalSet.Count

    # Each number in binary representation in a range from 0 to 2^n does exactly what we need:
    # it shoes by its bits (0 or 1) whether to include related element from the set or not.
    # For example, for the set {1, 2, 3} the binary number of 010 would mean that we need to
    # include only "2" to the current set.
    for ($combinationIndex = 0; $combinationIndex -lt $numberOfCombinations; $combinationIndex += 1) {
        $subSet = @()

        for ($setElementIndex = 0; $setElementIndex -lt $originalSet.Count; $setElementIndex += 1) {
            if ( ($combinationIndex -band (1 -shl $setElementIndex)) -gt 0) {
                $subSet += $originalSet[$setElementIndex]
            }
        }

        # Add current subset to the list of all subsets.
        $subSets += ,$subSet
    }

    return $subSets
}

$pricesFilePath = ".\ms365prices.xlsx"
$stateFilePath = ".\creatingPricingPowerset_state"
$licensingServicePlanReferenceFilePath = ".\licensing-service-plan-reference-modifed.csv"

$pricesFilePath = ""
$lastPricesFileHash = ""

If (Test-Path $pricesFilePath) {
    $pricesFileHash = (Get-FileHash -Path $pricesFilePath).Hash
}

If (Test-Path $stateFilePath) {
    $lastPricesFileHash = Get-Content -Path $stateFilePath
}

If ($pricesFileHash -eq $lastPricesFileHash) {
    Write-Host "The prices file contains same hash as last run. Please delete state file to re-run this script."
    Write-Error "ErrorAction" -ErrorAction Stop
}

$pricesFileHash | Out-File $stateFilePath

$licensingServicePlanReference = Import-Csv -path $licensingServicePlanReferenceFilePath

$prices = Import-Excel -Path $pricesFilePath

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


# Creating a PowerSet from possiable product combinations
Write-Host "Creating PowerSet"
$productCollectioPowersets = @()
$productCollectioPowersets = powerSet $productsFound
Write-Host "Finished creating PowerSet"

Write-Host "Creating PowerSet Array"
$powersetData = @()

ForEach ($productCollectioPowerset in $productCollectioPowersets) {

    $productPricingArray = @()

    ForEach ($product in $productCollectioPowerset) {
        $productPricingArray += ($productsData | Where-Object {$_.Productname -eq $product}).price
    }

    $sum = ($productPricingArray | Measure-Object -Sum).Sum

    $productServicesName = @()

    ForEach ($product in $productCollectioPowerset) {

        $productServices = ($licensingServicePlanReference | Where-Object {$_.Productname -eq $product}).Serviceplansincluded

        $productServicesSplit = $productServices.Split([Environment]::NewLine) | Where-Object {$_ -ne ""}

        ForEach ($productService in $productServicesSplit) {
            $productServicesName += $productService.Split(" ")[0]
        }
    }
    
    $productServicesNameUnique = $productServicesName | Sort-Object -Unique


    $powersetDetails = [PSCustomObject]@{
        Name = $productCollectioPowerset
        Sum = $sum
        Services = $productServicesNameUnique
    }

    $powersetData += $powersetDetails
}
Write-Host "Finished creating PowerSet Array"

$powersetDataJoin = $powersetData | Select-Object @{ label='Name'; expression={ $_.Name -join ","}}, Sum,@{ label='Services'; expression={ $_.Services -join ","}} 

$powersetDataExport = $powersetDataJoin | Select-Object Name,@{ label = "Sum";  Expression = {($_.Sum).Replace(",",".")}},Services
$powersetDataExport | Export-Csv -Encoding UTF8 -Path .\creatingPricingPowerset.csv


# Reduces the 262144 combinations to 28111, this is expected on MS 365 E5 including everything you need.
$powersetDataSortbySum = $powersetDataJoin | Sort-Object -Property Sum
$powersetDataSumLT402 = $powersetDataSortbySum | Where-Object {$_.Sum -lt 402 }

$powersetDataSumLT402Export = $powersetDataSumLT402 | Select-Object Name,@{ label = "Sum";  Expression = {($_.Sum.ToString()).Replace(",",".")}},Services
$powersetDataSumLT402Export | Export-Csv -Path .\powersetDataSumLT402.csv


# Time info
$endTime = Get-Date
$time = ($endTime - $runTime).TotalSeconds

Write-Host "This script took"$time" seconds to run."
