<#
.SYNOPSIS
Yearly mailbox storage estimate (MB & GB) for an Exchange Online mailbox.

.DESCRIPTION
This script connects to Exchange Online, gets mailbox folder statistics,
parses folder size correctly, and estimates storage usage by year (based on newest item date).

.How to Run this Script
.\Get-MailboxYearlyUsage.ps1 -UserPrincipalName admin@M365MCP29765884.onmicrosoft.com
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName
)

# --- Connect ---
Connect-ExchangeOnline -ShowBanner:$false | Out-Null
Write-Host "Fetching mailbox folder statistics for $UserPrincipalName..."

# Get folder stats with oldest/newest dates
$items = Get-ExoMailboxFolderStatistics -Identity $UserPrincipalName -IncludeOldestAndNewestItems |
    Select-Object Name, FolderPath, FolderSize, ItemsInFolder, OldestItemReceivedDate, NewestItemReceivedDate

# --- Helpers ---
function Get-BytesFromSizeString {
    param([string]$SizeString)

    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }

    # Normalize and strip NBSP
    $s = ($SizeString -replace [char]0x00A0, ' ').Trim()

    # 1) Look for "(n bytes)" with commas
    if ($s -match '\(([\d,]+)\s*bytes\)') {
        $num = ($matches[1] -replace '[^0-9]', '')
        if ($num) { return [double]$num }
    }

    # 2) Fallback parse "n.nn UNIT"
    if ($s -match '([0-9]*[.,]?[0-9]+)\s*(B|KB|MB|GB|TB)') {
        $val  = [double](($matches[1] -replace ',', '.'))
        $unit = $matches[2].ToUpper()
        switch ($unit) {
            'B'  { return $val }
            'KB' { return $val * 1KB }
            'MB' { return $val * 1MB }
            'GB' { return $val * 1GB }
            'TB' { return $val * 1TB }
            default { return 0 }
        }
    }

    return 0
}

# --- Aggregate by year ---
$yearly = @{}

foreach ($f in $items) {
    if (-not $f.NewestItemReceivedDate -or $f.ItemsInFolder -le 0) { continue }

    $year = $f.NewestItemReceivedDate.Year
    $bytes = Get-BytesFromSizeString $f.FolderSize

    if ($bytes -le 0) { continue }

    if (-not $yearly.ContainsKey($year)) { $yearly[$year] = 0.0 }
    $yearly[$year] += $bytes
}

# --- Build report ---
$report = $yearly.Keys |
    Sort-Object |
    ForEach-Object {
        $bytes = [double]$yearly[$_]
        [PSCustomObject]@{
            Year                  = $_
            'Estimated Size (MB)' = [math]::Round($bytes / 1MB, 2)
            'Estimated Size (GB)' = [math]::Round($bytes / 1GB, 3)
        }
    }

if (-not $report) {
    Write-Warning "No sizes could be parsed. Run with -Verbose to see raw folder sizes."
    $items | Select-Object -First 10 Name, FolderSize, ItemsInFolder, NewestItemReceivedDate | Format-Table -AutoSize
} else {
    $report | Format-Table -AutoSize
    $csv = ".\MailboxYearlyUsage_$($UserPrincipalName).csv"
    $report | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csv
    Write-Host "Saved: $csv"
}

Disconnect-ExchangeOnline

