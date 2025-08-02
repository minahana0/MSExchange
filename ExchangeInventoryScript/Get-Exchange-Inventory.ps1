<#
.SYNOPSIS
Exchange Inventory Script
🏆What This Exchange Inventory Script Does
🟢 Creates an output folder: C:\Exchange_Inventory
🟢 Exports each section to a separate CSV file
🟢 Uses Get-ExchangeServer, Get-MailboxDatabase, Get-SendConnector, Get-    ReceiveConnector, Get-ExchangeCertificate, Get-  HybridConfiguration, and Get-ADSite
🟢 Resolves IP address for each Exchange server using .NET DNS lookup
🟢 Checks and exports Hybrid config if it exists

⚙️ How To Run
🌐 Save the script as Get-Exchange-Inventory.ps1.
⚙️ Run it in the Exchange Management Shell, or use remote PowerShell with Exchange Online if hybrid is involved.
🧱Make sure you have the required permissions (Organization Management).
⚙️Adjust $OutputPath if needed.

.DESCRIPTION
This script collects comprehensive information about:
1. Exchange Sites
2. Exchange Servers (name, version, OS, site, IP, role)
3. Databases and file paths
4. Send Connectors
5. Receive Connectors
6. Public Certificates assigned to IIS
7. Hybrid Configuration

.EXPORTS
Each section exports results to a separate CSV file in C:\Exchange_Inventory

#>

# Make sure you are running in Exchange Management Shell or connected via remote PowerShell

# Output folder
$OutputPath = "C:\Exchange_Inventory"
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath
}

Write-Host "Starting Exchange Inventory Collection..." -ForegroundColor Cyan

# 1. List all Exchange Sites
try {
    $Sites = Get-ADSite | Select-Object Name
    $Sites | Export-Csv "$OutputPath\Exchange_Sites.csv" -NoTypeInformation
    Write-Host "Exported Exchange Sites"
} catch {
    Write-Host "Failed to get AD Sites: $_"
}

# 2. List all Exchange Servers with details
try {
    $Servers = Get-ExchangeServer | ForEach-Object {
        $IPAddress = try {
            [System.Net.Dns]::GetHostAddresses($_.Fqdn) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
        } catch {
            $null
        }
        [PSCustomObject]@{
            Name             = $_.Name
            ExchangeVersion  = $_.AdminDisplayVersion
            Edition          = $_.Edition
            Site             = $_.Site
            Roles            = $_.ServerRole
            FQDN             = $_.Fqdn
            OS               = $_.OperatingSystem
            IPAddress        = $IPAddress?.IPAddressToString
        }
    }
    $Servers | Export-Csv "$OutputPath\Exchange_Servers.csv" -NoTypeInformation
    Write-Host "Exported Exchange Servers"
} catch {
    Write-Host "Failed to get Exchange Servers: $_"
}

# 3. List all databases and file paths
try {
    $Databases = Get-MailboxDatabase | Select-Object Name, EdbFilePath, LogFolderPath, Server
    $Databases | Export-Csv "$OutputPath\Exchange_Databases.csv" -NoTypeInformation
    Write-Host "Exported Mailbox Databases"
} catch {
    Write-Host "Failed to get Databases: $_"
}

# 4. List all Send Connectors
try {
    $SendConnectors = Get-SendConnector | Select-Object Name, AddressSpaces, SmartHosts, DNSRoutingEnabled, SourceTransportServers, Enabled, Cost, IsScopedConnector
    $SendConnectors | Export-Csv "$OutputPath\Exchange_SendConnectors.csv" -NoTypeInformation
    Write-Host "Exported Send Connectors"
} catch {
    Write-Host "Failed to get Send Connectors: $_"
}

# 5. List all Receive Connectors
try {
    $ReceiveConnectors = Get-ReceiveConnector | Select-Object Name, Server, Bindings, RemoteIPRanges, AuthMechanism, PermissionGroups, TransportRole, ProtocolLoggingLevel
    $ReceiveConnectors | Export-Csv "$OutputPath\Exchange_ReceiveConnectors.csv" -NoTypeInformation
    Write-Host "Exported Receive Connectors"
} catch {
    Write-Host "Failed to get Receive Connectors: $_"
}

# 6. List of public certificates assigned to IIS
try {
    $Certificates = Get-ExchangeCertificate | Where-Object {$_.Services -like "*IIS*"} | Select-Object Thumbprint, Subject, Issuer, Services, NotAfter, FriendlyName
    $Certificates | Export-Csv "$OutputPath\Exchange_IIS_Certificates.csv" -NoTypeInformation
    Write-Host "Exported IIS Certificates"
} catch {
    Write-Host "Failed to get Exchange Certificates: $_"
}

# 7. Check for Hybrid Configuration
try {
    $HybridConfig = Get-HybridConfiguration
    if ($HybridConfig) {
        $HybridConfig | Select-Object * | Export-Csv "$OutputPath\Exchange_HybridConfiguration.csv" -NoTypeInformation
        Write-Host "Hybrid Configuration found and exported"
    } else {
        Write-Host "No Hybrid Configuration found"
    }
} catch {
    Write-Host "Failed to get Hybrid Configuration: $_"
}

Write-Host "Exchange Inventory collection completed. Files saved to: $OutputPath" -ForegroundColor Green

