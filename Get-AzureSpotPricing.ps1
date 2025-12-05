<#
.SYNOPSIS
    Queries Azure Cost API for VM spot pricing, eviction rates, and pay-as-you-go costs.

.DESCRIPTION
    This script retrieves Azure VM pricing information including spot pricing per hour,
    eviction rates, and standard pay-as-you-go costs for specified regions and SKUs.

.PARAMETER Region
    The Azure regions to query. Default is "uksouth". Can specify multiple regions.

.PARAMETER SkuFilter
    The SKU filter patterns. Default is "Standard_D*s_v5" (Cobalt SKUs). Can specify multiple patterns.

.EXAMPLE
    .\Get-AzureSpotPricing.ps1
    Uses default values (UK South, Cobalt SKUs)

.EXAMPLE
    .\Get-AzureSpotPricing.ps1 -Region "eastus" -SkuFilter "Standard_E*s_v5"
    Queries East US region for E-series v5 SKUs

.EXAMPLE
    .\Get-AzureSpotPricing.ps1 -Region "swedencentral","uksouth" -SkuFilter "Standard_D*ps_v5","Standard_E*s_v5"
    Queries multiple regions and SKU patterns
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$Region = @("uksouth"),
    
    [Parameter(Mandatory = $false)]
    [string[]]$SkuFilter = @("Standard_D*s_v5")
)

# Check if user is logged into Azure
try {
    $context = Get-AzContext -ErrorAction Stop
    if (-not $context) {
        Write-Host "Please login to Azure first using Connect-AzAccount" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "Please login to Azure first using Connect-AzAccount" -ForegroundColor Red
    exit 1
}

Write-Host "`nQuerying Azure Pricing for:" -ForegroundColor Cyan
Write-Host "  Regions: $($Region -join ', ')" -ForegroundColor Yellow
Write-Host "  SKU Filters: $($SkuFilter -join ', ')" -ForegroundColor Yellow
Write-Host "`nFetching pricing data...`n" -ForegroundColor Cyan

function Get-AzurePricing {
    param(
        [string]$FilterQuery
    )
    
    $baseUri = "https://prices.azure.com/api/retail/prices"
    $results = @()
    $encodedFilter = [System.Web.HttpUtility]::UrlEncode($FilterQuery)
    $nextPageLink = "$baseUri`?`$filter=$encodedFilter"
    
    while ($nextPageLink) {
        try {
            $response = Invoke-RestMethod -Uri $nextPageLink -Method Get -ErrorAction Stop
            $results += $response.Items
            $nextPageLink = $response.NextPageLink
            
            # Limit to prevent too many pages
            if ($results.Count -gt 10000) {
                Write-Verbose "Reached result limit of 10000 items"
                break
            }
        } catch {
            $errorMessage = $_.Exception.Message
            if ($_.ErrorDetails.Message) {
                $errorMessage = $_.ErrorDetails.Message
            }
            Write-Warning "Error fetching pricing data: $errorMessage"
            break
        }
    }
    
    return $results
}

function Get-SpotEvictionRates {
    param(
        [string]$RegionName,
        [array]$SkuList
    )
    
    if ($SkuList.Count -eq 0) {
        return @{}
    }
    
    Write-Host "  Fetching Spot eviction rates..." -ForegroundColor Gray
    
    try {
        $evictionRates = @{}
        $failedSkus = @()
        
        # Initial delay to let API stabilize
        Start-Sleep -Milliseconds 500
        
        foreach ($sku in $SkuList) {
            $retryCount = 0
            $maxRetries = 4
            $success = $false
            
            while (-not $success -and $retryCount -le $maxRetries) {
                try {
                    $desiredSize = @(@{sku = $sku})
                    $response = Invoke-AzSpotPlacementScore -Location $RegionName -DesiredCount 10 -DesiredLocation @($RegionName) -DesiredSize $desiredSize -ErrorAction Stop -WarningAction SilentlyContinue
                    
                    if ($response.PlacementScore -and $response.PlacementScore.Count -gt 0 -and $response.PlacementScore[0].Score) {
                        $score = $response.PlacementScore[0]
                        $evictionRates[$sku] = @{
                            Risk = switch ($score.Score) {
                                'High' { 'Very Low' }
                                'Medium' { 'Low-Medium' }
                                'Low' { 'Medium-High' }
                                'None' { 'High' }
                                default { 'N/A' }
                            }
                            Percent = switch ($score.Score) {
                                'High' { '0-5%' }
                                'Medium' { '5-15%' }
                                'Low' { '15-30%' }
                                'None' { '30%+' }
                                default { 'N/A' }
                            }
                        }
                        Write-Verbose "✓ $sku : $($score.Score)"
                        $success = $true
                    }
                    
                } catch {
                    $retryCount++
                    if ($retryCount -le $maxRetries) {
                        # Progressive backoff: 1s, 2s, 3s, 4s
                        Start-Sleep -Seconds $retryCount
                    } else {
                        $failedSkus += $sku
                    }
                }
            }
            
            # Consistent delay between requests
            if (-not $success) {
                Start-Sleep -Milliseconds 800
            } else {
                Start-Sleep -Milliseconds 400
            }
        }
        
        # Second pass for failed SKUs with longer delays
        if ($failedSkus.Count -gt 0 -and $failedSkus.Count -le 5) {
            Write-Verbose "Second pass: retrying $($failedSkus.Count) failed SKUs..."
            Start-Sleep -Seconds 5
            
            foreach ($sku in $failedSkus) {
                try {
                    $desiredSize = @(@{sku = $sku})
                    $response = Invoke-AzSpotPlacementScore -Location $RegionName -DesiredCount 10 -DesiredLocation @($RegionName) -DesiredSize $desiredSize -ErrorAction Stop -WarningAction SilentlyContinue
                    
                    if ($response.PlacementScore -and $response.PlacementScore.Count -gt 0 -and $response.PlacementScore[0].Score) {
                        $score = $response.PlacementScore[0]
                        $evictionRates[$sku] = @{
                            Risk = switch ($score.Score) {
                                'High' { 'Very Low' }
                                'Medium' { 'Low-Medium' }
                                'Low' { 'Medium-High' }
                                'None' { 'High' }
                                default { 'N/A' }
                            }
                            Percent = switch ($score.Score) {
                                'High' { '0-5%' }
                                'Medium' { '5-15%' }
                                'Low' { '15-30%' }
                                'None' { '30%+' }
                                default { 'N/A' }
                            }
                        }
                        Write-Verbose "✓ $sku : $($score.Score) (retry)"
                    }
                    Start-Sleep -Seconds 2
                } catch {
                    # Silent fail on second pass
                }
            }
        }
        
        return $evictionRates
    } catch {
        Write-Warning "Could not retrieve eviction rates: $($_.Exception.Message)"
        return @{}
    }
}

# Combined results from all regions and SKU filters
$allPricingTable = @{}

# Loop through each region and SKU filter combination
foreach ($currentRegion in $Region) {
    foreach ($currentSkuFilter in $SkuFilter) {
        Write-Host "Processing Region: $currentRegion, SKU Filter: $currentSkuFilter" -ForegroundColor Cyan

        # Fetch pricing data
        Write-Host "  Fetching Pay-As-You-Go pricing..." -ForegroundColor Gray
        $payAsYouGoData = Get-AzurePricing -FilterQuery "serviceName eq 'Virtual Machines' and armRegionName eq '$currentRegion' and priceType eq 'Consumption' and currencyCode eq 'USD'"

        Write-Host "  Fetching Spot pricing..." -ForegroundColor Gray
        # Spot pricing doesn't have a separate priceType, so fetch all Consumption and filter by meterName
        $spotData = $payAsYouGoData | Where-Object { $_.meterName -like '*Spot*' }

        # Filter for Windows and Linux, and match SKU pattern
        Write-Host "  Filtering $($payAsYouGoData.Count) Pay-As-You-Go and $($spotData.Count) Spot items..." -ForegroundColor Gray

        $payAsYouGoData = $payAsYouGoData | Where-Object { 
            $_.armSkuName -like $currentSkuFilter -and 
            $_.productName -match "Virtual Machines.*Series" -and
            $_.unitOfMeasure -eq "1 Hour" -and
            $_.meterName -notlike '*Spot*' -and
            $_.meterName -notlike '*Low Priority*'
        }

        $spotData = $spotData | Where-Object { 
            $_.armSkuName -like $currentSkuFilter -and 
            $_.productName -match "Virtual Machines.*Series" -and
            $_.unitOfMeasure -eq "1 Hour"
        }

        Write-Host "  Found $($payAsYouGoData.Count) Pay-As-You-Go prices and $($spotData.Count) Spot prices" -ForegroundColor Gray

        # Create a combined dataset for this region/SKU combo
        $pricingTable = @{}

        # Process Pay-As-You-Go data
        foreach ($item in $payAsYouGoData) {
            $sku = $item.armSkuName
            $os = if ($item.productName -match "Windows") { "Windows" } else { "Linux" }
            $regionKey = $item.armRegionName
            $key = "$regionKey-$sku-$os"
            
            Write-Verbose "PayAsYouGo: $key - $($item.meterName) - Price: $($item.retailPrice)"
            
            if (-not $pricingTable.ContainsKey($key)) {
                $pricingTable[$key] = @{
                    SKU = $sku
                    OS = $os
                    Region = $item.armRegionName
                    PayAsYouGo = $item.retailPrice
                    SpotPrice = $null
                    EvictionRate = "N/A"
                    EvictionPercent = "N/A"
                    Savings = $null
                }
            } else {
                $pricingTable[$key].PayAsYouGo = $item.retailPrice
            }
        }

        # Process Spot data
        foreach ($item in $spotData) {
            $sku = $item.armSkuName
            $os = if ($item.productName -match "Windows") { "Windows" } else { "Linux" }
            $regionKey = $item.armRegionName
            $key = "$regionKey-$sku-$os"
            
            Write-Verbose "Spot: $key - $($item.meterName) - Price: $($item.retailPrice)"
            
            if (-not $pricingTable.ContainsKey($key)) {
                $pricingTable[$key] = @{
                    SKU = $sku
                    OS = $os
                    Region = $item.armRegionName
                    PayAsYouGo = $null
                    SpotPrice = $item.retailPrice
                    EvictionRate = "N/A"
                    EvictionPercent = "N/A"
                    Savings = $null
                }
            } else {
                $pricingTable[$key].SpotPrice = $item.retailPrice
            }
        }

        # Calculate savings percentage
        foreach ($key in $pricingTable.Keys) {
            $entry = $pricingTable[$key]
            if ($entry.PayAsYouGo -and $entry.SpotPrice -and $entry.PayAsYouGo -gt 0) {
                $savingsPercent = (($entry.PayAsYouGo - $entry.SpotPrice) / $entry.PayAsYouGo) * 100
                $entry.Savings = "{0:N1}%" -f $savingsPercent
            } else {
                $entry.Savings = "N/A"
            }
        }

        # Get eviction rates using Spot Placement Score API
        $uniqueSkus = $pricingTable.Keys | ForEach-Object { $pricingTable[$_].SKU } | Select-Object -Unique
        
        if ($uniqueSkus.Count -gt 0) {
            $evictionRates = Get-SpotEvictionRates -RegionName $currentRegion -SkuList $uniqueSkus
        } else {
            $evictionRates = @{}
        }

        # Update eviction rates in pricing table
        foreach ($key in $pricingTable.Keys) {
            $sku = $pricingTable[$key].SKU
            if ($evictionRates.ContainsKey($sku)) {
                $pricingTable[$key].EvictionRate = $evictionRates[$sku].Risk
                $pricingTable[$key].EvictionPercent = $evictionRates[$sku].Percent
            }
            # If API didn't return data, leave as N/A (no estimation)
        }

        # Merge into combined results
        foreach ($key in $pricingTable.Keys) {
            $allPricingTable[$key] = $pricingTable[$key]
        }
        
        Write-Host "" -ForegroundColor Gray
    }
}

Write-Host "`n=== Azure VM Spot Pricing Table ===" -ForegroundColor Green
Write-Host "" 

# Convert to array and sort, filter to show only rows with complete data or spot pricing
$results = $allPricingTable.Values | Where-Object { 
    ($_.SpotPrice -ne $null) -or ($_.PayAsYouGo -ne $null -and $_.SpotPrice -ne $null) 
} | Sort-Object { 
    # Sort by whether it has complete data (both prices), then by SKU
    if ($_.SpotPrice -and $_.PayAsYouGo) { 0 } else { 1 }
}, Region, SKU, OS

# Display results
if ($results.Count -eq 0) {
    Write-Host "No pricing data found for the specified regions and SKU filters." -ForegroundColor Red
    Write-Host "`nTry adjusting your parameters:" -ForegroundColor Yellow
    Write-Host "  - Verify the region names (e.g., 'eastus', 'westeurope', 'uksouth')" -ForegroundColor Gray
    Write-Host "  - Check the SKU filter patterns (e.g., 'Standard_D*s_v5', 'Standard_E*')" -ForegroundColor Gray
} else {
    $tableData = $results | ForEach-Object {
        [PSCustomObject]@{
            'VM SKU' = $_.SKU
            'OS' = $_.OS
            'Region' = $_.Region
            'Spot ($/hr)' = if ($_.SpotPrice) { "`${0:N4}" -f $_.SpotPrice } else { "N/A" }
            'PayAsYouGo ($/hr)' = if ($_.PayAsYouGo) { "`${0:N4}" -f $_.PayAsYouGo } else { "N/A" }
            'Savings' = $_.Savings
            'Eviction Risk' = $_.EvictionRate
            'Eviction Rate %' = $_.EvictionPercent
        }
    }
    
    $tableData | Format-Table -AutoSize
    
    Write-Host "`nTotal SKUs found: $($results.Count)" -ForegroundColor Cyan
    Write-Host "Total regions queried: $($Region.Count)" -ForegroundColor Cyan
    
    # Auto-export to Excel
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $excelPath = Join-Path $PSScriptRoot "AzureSpotPricing_$timestamp.xlsx"
    
    try {
        # Check if ImportExcel module is available
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "`nInstalling ImportExcel module..." -ForegroundColor Yellow
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        }
        
        Import-Module ImportExcel -ErrorAction Stop
        
        # Export to Excel with formatting
        $excelData = $results | ForEach-Object {
            [PSCustomObject]@{
                'VM SKU' = $_.SKU
                'OS' = $_.OS
                'Region' = $_.Region
                'Spot ($/hr)' = if ($_.SpotPrice) { $_.SpotPrice } else { $null }
                'PayAsYouGo ($/hr)' = if ($_.PayAsYouGo) { $_.PayAsYouGo } else { $null }
                'Savings' = $_.Savings
                'Eviction Risk' = $_.EvictionRate
                'Eviction Rate %' = $_.EvictionPercent
            }
        }
        
        $excelData | Export-Excel -Path $excelPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
            -WorksheetName "Spot Pricing" -TableName "SpotPricing" -TableStyle Medium2
        
        Write-Host "`n✓ Exported to Excel: $excelPath" -ForegroundColor Green
        
    } catch {
        Write-Host "`nFailed to export to Excel: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Exporting to CSV instead..." -ForegroundColor Gray
        
        $csvPath = Join-Path $PSScriptRoot "AzureSpotPricing_$timestamp.csv"
        $tableData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "✓ Exported to CSV: $csvPath" -ForegroundColor Green
    }
}
