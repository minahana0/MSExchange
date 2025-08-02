# MSExchange
This Repo contains scripts for MS Exchange Server and Exchange Online

📂 What This Script Does
✅ Creates an output folder: C:\Exchange_Inventory
✅ Exports each section to a separate CSV file
✅ Uses Get-ExchangeServer, Get-MailboxDatabase, Get-SendConnector, Get-ReceiveConnector, Get-ExchangeCertificate, Get-HybridConfiguration, and Get-ADSite
✅ Resolves IP address for each Exchange server using .NET DNS lookup
✅ Checks and exports Hybrid config if it exists

⚙️ How To Run
Save the script as Get-Exchange-Inventory.ps1.
Run it in the Exchange Management Shell, or use remote PowerShell with Exchange Online if hybrid is involved.
Make sure you have the required permissions (Organization Management).
Adjust $OutputPath if needed.