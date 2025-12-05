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
    [string[]]$SkuFilter = @("Standard_D*s_v5"),
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipEvictionRates
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

function Get-StaticEvictionEstimate {
    param(
        [string]$VmSize
    )
    
    # Extract vCPU count from VM size (e.g., D16ps_v6 = 16 vCPUs)
    if ($VmSize -match '_(\d+)') {
        $vCpuCount = [int]$Matches[1]
    } else {
        $vCpuCount = 2  # Default fallback
    }
    
    # Larger VMs generally have lower eviction rates
    # Based on Azure's historical spot patterns
    if ($vCpuCount -ge 64) {
        return @{
            Risk = 'Very Low'
            Percent = '0-5% (Est)'
            IsEstimate = $true
        }
    } elseif ($vCpuCount -ge 32) {
        return @{
            Risk = 'Low'
            Percent = '5-10% (Est)'
            IsEstimate = $true
        }
    } elseif ($vCpuCount -ge 16) {
        return @{
            Risk = 'Low-Medium'
            Percent = '10-15% (Est)'
            IsEstimate = $true
        }
    } elseif ($vCpuCount -ge 8) {
        return @{
            Risk = 'Medium'
            Percent = '15-20% (Est)'
            IsEstimate = $true
        }
    } else {
        return @{
            Risk = 'Medium-High'
            Percent = '20-30% (Est)'
            IsEstimate = $true
        }
    }
}

function Get-SpotEvictionViaRestAPI {
    param(
        [string]$SubscriptionId,
        [string]$Location,
        [string]$VmSize
    )
    
    try {
        # Get access token (suppress warning)
        $token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com" -WarningAction SilentlyContinue).Token
        
        # Try multiple API versions - the 2024-03-01 might not be available in all regions
        $apiVersions = @("2024-07-01", "2024-03-01", "2023-09-01")
        
        foreach ($apiVersion in $apiVersions) {
            try {
                # Build REST API request
                $uri = "https://management.azure.com/subscriptions/$SubscriptionId/providers/Microsoft.Compute/locations/$Location/spotPlacementScores?api-version=$apiVersion"
                
                $body = @{
                    location = $Location
                    desiredCount = 100
                    desiredSizes = @(
                        @{ sku = $VmSize }
                    )
                    desiredLocations = @($Location)
                } | ConvertTo-Json -Depth 5
                
                $headers = @{
                    "Authorization" = "Bearer $token"
                    "Content-Type" = "application/json"
                }
                
                $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ErrorAction Stop
                
                if ($response.placementScore -and $response.placementScore.Count -gt 0) {
                    $score = $response.placementScore[0]
                    return @{
                        Success = $true
                        Score = $score.score
                        Risk = switch ($score.score) {
                            'High' { 'Very Low' }
                            'Medium' { 'Low-Medium' }
                            'Low' { 'Medium-High' }
                            'None' { 'High' }
                            default { 'Unknown' }
                        }
                        Percent = switch ($score.score) {
                            'High' { '0-5%' }
                            'Medium' { '5-15%' }
                            'Low' { '15-25%' }
                            'None' { '>25%' }
                            default { 'Unknown' }
                        }
                    }
                }
                
                # If we got here, try next API version
            } catch {
                # Continue to next API version
                continue
            }
        }
        
        return @{ Success = $false; Error = "No placement score in response from any API version" }
        
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
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
        $batchSize = 2  # Reduce to 2 SKUs at once for better stability
        
        # Initial stabilization delay
        Start-Sleep -Milliseconds 1000
        
        # Split SKUs into batches
        for ($i = 0; $i -lt $SkuList.Count; $i += $batchSize) {
            $batch = $SkuList[$i..[Math]::Min($i + $batchSize - 1, $SkuList.Count - 1)]
            
            # Start parallel jobs for this batch
            $jobs = @()
            foreach ($sku in $batch) {
                $job = Start-Job -ScriptBlock {
                    param($RegionName, $Sku)
                    
                    $retryCount = 0
                    $maxRetries = 4
                    $lastError = $null
                    
                    # Random initial delay to stagger requests
                    Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 800)
                    
                    while ($retryCount -le $maxRetries) {
                        try {
                            $desiredSize = @(@{sku = $Sku})
                            $response = Invoke-AzSpotPlacementScore -Location $RegionName -DesiredCount 10 -DesiredLocation @($RegionName) -DesiredSize $desiredSize -ErrorAction Stop -WarningAction SilentlyContinue
                            
                            if ($response.PlacementScore -and $response.PlacementScore.Count -gt 0 -and $response.PlacementScore[0].Score) {
                                return @{
                                    Sku = $Sku
                                    Score = $response.PlacementScore[0].Score
                                    Success = $true
                                }
                            }
                        } catch {
                            $lastError = $_.Exception.Message
                            $retryCount++
                            if ($retryCount -le $maxRetries) {
                                # Progressive delay with randomness: 2s, 3s, 4s, 5s
                                Start-Sleep -Seconds ($retryCount + 1)
                            }
                        }
                    }
                    
                    return @{Sku = $Sku; Success = $false; Error = $lastError}
                } -ArgumentList $RegionName, $sku
                
                $jobs += $job
            }
            
            # Wait for batch to complete (max 25 seconds per job for more retries)
            $completedJobs = $jobs | Wait-Job -Timeout 25
            $results = $completedJobs | Receive-Job
            
            # Handle timed-out jobs
            $timedOutJobs = $jobs | Where-Object { $_.State -eq 'Running' }
            foreach ($job in $timedOutJobs) {
                # Extract SKU from job (it's in the arguments)
                $sku = $job.ChildJobs[0].Runspace.SessionStateProxy.GetVariable('Sku')
                $results += @{Sku = $sku; Success = $false; Error = "Timeout"}
            }
            
            $jobs | Remove-Job -Force
            
            # Process results
            foreach ($result in $results) {
                if ($result.Success -and $result.Score) {
                    $evictionRates[$result.Sku] = @{
                        Risk = switch ($result.Score) {
                            'High' { 'Very Low' }
                            'Medium' { 'Low-Medium' }
                            'Low' { 'Medium-High' }
                            'None' { 'High' }
                            default { 'N/A' }
                        }
                        Percent = switch ($result.Score) {
                            'High' { '0-5%' }
                            'Medium' { '5-15%' }
                            'Low' { '15-30%' }
                            'None' { '30%+' }
                            default { 'N/A' }
                        }
                        Error = $null
                    }
                    Write-Verbose "✓ $($result.Sku) : $($result.Score)"
                } elseif ($result.Error) {
                    # Store error information
                    $errorType = if ($result.Error -eq "Timeout") { "Job Timeout" }
                                elseif ($result.Error -match "Expected.*Was String") { "Parser Error" } 
                                elseif ($result.Error -match "throttl|rate") { "Rate Limited" }
                                elseif ($result.Error -match "BadRequest") { "Bad Request" }
                                else { "API Failed" }
                    
                    $evictionRates[$result.Sku] = @{
                        Risk = $errorType
                        Percent = $errorType
                        Error = $errorType
                    }
                }
            }
            
            # Longer delay between batches
            if ($i + $batchSize -lt $SkuList.Count) {
                Start-Sleep -Seconds 2
            }
        }
        
        # Third pass: aggressive retry on all failures
        $failedSkus = $SkuList | Where-Object { -not $evictionRates.ContainsKey($_) -or $evictionRates[$_].Error }
        if ($failedSkus.Count -gt 0 -and $failedSkus.Count -le 8) {
            Write-Verbose "Final retry pass: $($failedSkus.Count) failed SKUs..."
            Start-Sleep -Seconds 5
            
            foreach ($sku in $failedSkus) {
                # Try twice with long delays
                for ($attempt = 1; $attempt -le 2; $attempt++) {
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
                                Error = $null
                            }
                            Write-Verbose "✓ $sku : $($score.Score) (final retry)"
                            break
                        }
                        Start-Sleep -Seconds 3
                    } catch {
                        if ($attempt -eq 2) {
                            # Final failure - keep error
                            $errorMsg = $_.Exception.Message
                            $evictionRates[$sku] = @{
                                Risk = if ($errorMsg -match "Expected.*Was String") { "Parser Error" } 
                                      elseif ($errorMsg -match "throttl|rate") { "Rate Limited" }
                                      elseif ($errorMsg -match "BadRequest") { "Bad Request" }
                                      else { "API Failed" }
                                Percent = if ($errorMsg -match "Expected.*Was String") { "Parser Error" } 
                                        elseif ($errorMsg -match "throttl|rate") { "Rate Limited" }
                                        elseif ($errorMsg -match "BadRequest") { "Bad Request" }
                                        else { "API Failed" }
                                Error = if ($errorMsg -match "Expected.*Was String") { "Parser Error" } 
                                       elseif ($errorMsg -match "throttl|rate") { "Rate Limited" }
                                       elseif ($errorMsg -match "BadRequest") { "Bad Request" }
                                       else { "API Failed" }
                            }
                        } else {
                            Start-Sleep -Seconds 4
                        }
                    }
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

        # Don't query eviction rates yet - do it later after all pricing is collected
        
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

# NOW query eviction rates one-by-one for each unique SKU/Region combo AFTER all pricing is done
if (-not $SkipEvictionRates) {
    Write-Host "Fetching eviction rates for all SKUs (querying individually)..." -ForegroundColor Cyan
    $allEvictionRates = @{}

    # Group by region to minimize context switching
    $regionGroups = $results | Group-Object Region

    foreach ($regionGroup in $regionGroups) {
        $regionName = $regionGroup.Name
        $uniqueSkusInRegion = $regionGroup.Group | Select-Object -ExpandProperty SKU -Unique
    
    Write-Host "  Region: $regionName ($($uniqueSkusInRegion.Count) SKUs)" -ForegroundColor Gray
    
    # Try batch query first (up to 5 SKUs at once per API limit)
    if ($uniqueSkusInRegion.Count -le 5) {
        try {
            Write-Verbose "Attempting batch query for all $($uniqueSkusInRegion.Count) SKUs..."
            $desiredSizes = $uniqueSkusInRegion | ForEach-Object { @{sku = $_} }
            $response = Invoke-AzSpotPlacementScore -Location $regionName -DesiredCount 100 -DesiredLocation @($regionName) -DesiredSize $desiredSizes -ErrorAction Stop -WarningAction SilentlyContinue
        
        if ($response.PlacementScore -and $response.PlacementScore.Count -gt 0) {
            # Match scores to SKUs by index (they're returned in same order as input)
            for ($i = 0; $i -lt $response.PlacementScore.Count; $i++) {
                if ($i -lt $uniqueSkusInRegion.Count) {
                    $score = $response.PlacementScore[$i]
                    $scoreSku = $uniqueSkusInRegion[$i]
                    $key = "$regionName-$scoreSku"
                    
                    $allEvictionRates[$key] = @{
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
                    Write-Verbose "✓ $scoreSku : $($score.Score)"
                }
            }
            
            # If batch succeeded for all SKUs, skip individual queries
            if ($allEvictionRates.Keys.Where({$_ -like "$regionName-*"}).Count -eq $uniqueSkusInRegion.Count) {
                Write-Verbose "Batch query succeeded for all SKUs"
                continue
            }
        }
    } catch {
        Write-Verbose "Batch query failed, falling back to individual queries: $($_.Exception.Message)"
    }
    } else {
        Write-Verbose "Skipping batch query (>5 SKUs), using individual queries"
    }
    
    # Fallback: Query each SKU individually
    foreach ($sku in $uniqueSkusInRegion) {
        $key = "$regionName-$sku"
        
        # Skip if already got data from batch query
        if ($allEvictionRates.ContainsKey($key)) {
            continue
        }
        
        # Delay between individual SKU queries with random jitter
        Start-Sleep -Milliseconds (1500 + (Get-Random -Minimum 0 -Maximum 500))
        
        $retryCount = 0
        $maxRetries = 4
        $success = $false
        
        while (-not $success -and $retryCount -le $maxRetries) {
            try {
                # Try with higher DesiredCount to get more reliable results
                $desiredSize = @(@{sku = $sku})
                $response = Invoke-AzSpotPlacementScore -Location $regionName -DesiredCount 100 -DesiredLocation @($regionName) -DesiredSize $desiredSize -ErrorAction Stop -WarningAction SilentlyContinue
                
                if ($response.PlacementScore -and $response.PlacementScore.Count -gt 0 -and $response.PlacementScore[0].Score) {
                    $score = $response.PlacementScore[0]
                    $allEvictionRates[$key] = @{
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
                    Write-Verbose "✓ $sku in $regionName : $($score.Score)"
                    $success = $true
                }
            } catch {
                $retryCount++
                $errorMsg = $_.Exception.Message
                
                # Try REST API fallback on cmdlet failure
                if ($retryCount -eq 2) {
                    Write-Verbose "  Trying REST API fallback for $sku..."
                    $context = Get-AzContext
                    $restResult = Get-SpotEvictionViaRestAPI -SubscriptionId $context.Subscription.Id -Location $regionName -VmSize $sku
                    
                    if ($restResult.Success) {
                        $allEvictionRates[$key] = @{
                            Risk = $restResult.Risk
                            Percent = $restResult.Percent
                        }
                        Write-Verbose "✓ $sku in $regionName : $($restResult.Score) (REST API)"
                        $success = $true
                        continue
                    }
                }
                
                if ($retryCount -le $maxRetries) {
                    Write-Verbose "Retry $retryCount for $sku..."
                    # Longer progressive retry delays with jitter
                    Start-Sleep -Seconds (5 + ($retryCount * 2) + (Get-Random -Minimum 0 -Maximum 3))
                } else {
                    # Final fallback: Use static estimate based on VM size
                    $estimate = Get-StaticEvictionEstimate -VmSize $sku
                    $allEvictionRates[$key] = @{
                        Risk = $estimate.Risk
                        Percent = $estimate.Percent
                    }
                    Write-Verbose "⚠ $sku in $regionName : Using static estimate (API unavailable)"
                }
            }
        }
    }
}

    # Update results with eviction data
    foreach ($result in $results) {
        $key = "$($result.Region)-$($result.SKU)"
        if ($allEvictionRates.ContainsKey($key)) {
            $result.EvictionRate = $allEvictionRates[$key].Risk
            $result.EvictionPercent = $allEvictionRates[$key].Percent
        }
    }
} else {
    Write-Host "Skipping eviction rate queries (use without -SkipEvictionRates to include)" -ForegroundColor Yellow
}


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
