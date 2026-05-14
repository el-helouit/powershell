<#
.SYNOPSIS
Bulk removal of Microsoft 365 licences using Microsoft Graph PowerShell.

.DESCRIPTION
This script reads a CSV file containing UserId and SkuPartNumber values,
and attempts to remove the specified licence from each user.

The script includes:
- Dry run capability (no changes applied)
- Proper error handling using -ErrorAction Stop
- Explicit handling of group-based licensing scenarios
- Structured logging for audit and review purposes

.NOTES
Author:     Jake El-Helou
Company:    El-Helou IT
Location:   Sydney, Australia

Created:    2026-05-14
Version:    1.0

REQUIREMENTS:
- Microsoft.Graph PowerShell module
- Permissions:
    User.ReadWrite.All
    Directory.ReadWrite.All

AUTHENTICATION:
Recommended connection method:

    $env:AZURE_IDENTITY_DISABLE_WAM = "true"
    Connect-MgGraph `
        -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" `
        -UseDeviceAuthentication

INPUT:
- CSV file with headers:
    UserId,SkuPartNumber

OUTPUT:
- CSV file with per-user execution results:
    TimestampUtc,UserId,SkuPartNumber,Status,Reason

STATUS VALUES:
- Success
- Failed
- Skipped - Group Assigned
- DryRun

LIMITATIONS:
- Licences assigned via group-based licensing cannot be removed directly
- These will be logged as:
    "Skipped - Group Assigned"

CHANGE CONTROL:
- Always run in dry-run mode prior to execution
- Review output CSV before applying changes
- Retain output for audit tracking

#>


# =========================
# CONFIG
# =========================
$csvPath = "C:\Path\To\Your\File.csv"
$outputPath = "C:\Path\To\Output.csv"

$dryRun = $true   # $true = no changes, $false = execute
$forceCleanOutputFile = $true

# Optional: If you want quieter console output, set to $true
$quiet = $false

# =========================
# HELPER: Extract a useful error message (Graph sometimes returns JSON in ErrorDetails.Message)
# =========================
function Get-GraphErrorMessage {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    $msg = $ErrorRecord.Exception.Message

    # Sometimes Graph puts structured JSON inside ErrorDetails.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $raw = $ErrorRecord.ErrorDetails.Message

        # Try parse JSON if it looks like JSON
        if ($raw.Trim().StartsWith("{")) {
            try {
                $json = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($json.error -and $json.error.message) {
                    return $json.error.message
                }
            } catch {
                # Fall back to raw if JSON parse fails
                return $raw
            }
        }

        # Otherwise return the raw detail message
        return $raw
    }

    return $msg
}

# =========================
# PRE-FLIGHT: Clean output file to avoid duplicate headers / corrupt CSV
# =========================
if ($forceCleanOutputFile -and (Test-Path $outputPath)) {
    Remove-Item $outputPath -Force
}

# =========================
# LOAD CSV
# =========================
$users = Import-Csv -Path $csvPath

if (-not $users -or $users.Count -eq 0) {
    throw "CSV appears empty or could not be read: $csvPath"
}

# Validate required headers exist
$headers = $users[0].PSObject.Properties.Name
if ($headers -notcontains "UserId" -or $headers -notcontains "SkuPartNumber") {
    throw "CSV must contain headers: UserId, SkuPartNumber. Found: $($headers -join ', ')"
}

# =========================
# GET ALL SKUS ONCE
# =========================
$allSkus = Get-MgSubscribedSku -All

# Build lookup hashtable for quick resolution (case-insensitive keys)
$skuLookup = @{}
foreach ($s in $allSkus) {
    if ($s.SkuPartNumber -and -not $skuLookup.ContainsKey($s.SkuPartNumber.ToUpper())) {
        $skuLookup[$s.SkuPartNumber.ToUpper()] = $s
    }
}

# =========================
# RESULTS
# =========================
$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $users) {

    # Trim input values
    $userId = ($row.UserId | ForEach-Object { "$_".Trim() })
    $skuPartNumber = ($row.SkuPartNumber | ForEach-Object { "$_".Trim() })

    # Skip empty rows
    if ([string]::IsNullOrWhiteSpace($userId) -and [string]::IsNullOrWhiteSpace($skuPartNumber)) {
        if (-not $quiet) { Write-Warning "Skipping blank row" }
        continue
    }

    # Validate required values
    if ([string]::IsNullOrWhiteSpace($userId)) {
        if (-not $quiet) { Write-Warning "Skipping row with empty UserId (SKU: $skuPartNumber)" }
        $results.Add([PSCustomObject]@{
            TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
            UserId        = $row.UserId
            SkuPartNumber = $skuPartNumber
            Status        = "Failed"
            Reason        = "Missing UserId in CSV row"
        })
        continue
    }

    if ([string]::IsNullOrWhiteSpace($skuPartNumber)) {
        if (-not $quiet) { Write-Warning "Skipping row with empty SkuPartNumber (User: $userId)" }
        $results.Add([PSCustomObject]@{
            TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
            UserId        = $userId
            SkuPartNumber = $row.SkuPartNumber
            Status        = "Failed"
            Reason        = "Missing SkuPartNumber in CSV row"
        })
        continue
    }

    if (-not $quiet) { Write-Host "Processing: $userId | SKU: $skuPartNumber" }

    # Resolve SKU via lookup
    $lookupKey = $skuPartNumber.ToUpper()
    $sku = $null
    if ($skuLookup.ContainsKey($lookupKey)) {
        $sku = $skuLookup[$lookupKey]
    }

    if (-not $sku) {
        if (-not $quiet) { Write-Warning "SKU not found in tenant: $skuPartNumber" }

        $results.Add([PSCustomObject]@{
            TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
            UserId        = $userId
            SkuPartNumber = $skuPartNumber
            Status        = "Failed"
            Reason        = "SKU not found in tenant"
        })
        continue
    }

    # =========================
    # EXECUTION
    # =========================
    try {
        if ($dryRun) {
            if (-not $quiet) { Write-Host "[DRY RUN] Would remove SKU $skuPartNumber from $userId" }

            $results.Add([PSCustomObject]@{
                TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
                UserId        = $userId
                SkuPartNumber = $skuPartNumber
                Status        = "DryRun"
                Reason        = "No change executed"
            })
        }
        else {
            # IMPORTANT: -ErrorAction Stop makes Graph errors catchable (fixes false 'Success' logging)
            Set-MgUserLicense `
                -UserId $userId `
                -RemoveLicenses @($sku.SkuId) `
                -AddLicenses @() `
                -ErrorAction Stop

            $results.Add([PSCustomObject]@{
                TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
                UserId        = $userId
                SkuPartNumber = $skuPartNumber
                Status        = "Success"
                Reason        = ""
            })
        }
    }
    catch {
        $errorMessage = Get-GraphErrorMessage -ErrorRecord $_

        # Classify group-based licensing explicitly
        $status =
            if ($errorMessage -like "*inherited from a group*" -or $errorMessage -like "*cannot be removed directly from the user*") {
                "Skipped - Group Assigned"
            }
            else {
                "Failed"
            }

        if (-not $quiet) { Write-Warning "$status for $userId : $errorMessage" }

        $results.Add([PSCustomObject]@{
            TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
            UserId        = $userId
            SkuPartNumber = $skuPartNumber
            Status        = $status
            Reason        = $errorMessage
        })
    }
}

# =========================
# EXPORT RESULTS
# =========================
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "`nResults exported cleanly to: $outputPath"

# =========================
# SUMMARY
# =========================
Write-Host "`n===== SUMMARY ====="
$results |
    Group-Object Status |
    Sort-Object Count -Descending |
    Select-Object Name, Count |
    Format-Table -AutoSize
