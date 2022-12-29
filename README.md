# eDiscovery-Export-Tool

Powershell script to create and download eDiscovery export jobs from the Microsoft Purview Compliance Portal. Supports assigning export jobs by region and downloading via the command line.

### Prerequisites

- [Microsoft eDiscovery Export Tool](https://learn.microsoft.com/en-us/microsoft-365/compliance/ediscovery-configure-edge-to-export-search-results?view=o365-worldwide)
- [ExchangeOnlineManagement Module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/)
- [Figlet Module](https://www.powershellgallery.com/packages/Figlet/)
- Admin access to change WinRM registry setting to allow basic authentication (script will prompt if needed)

	[About the Exchange Online PowerShell module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#updates-for-version-300-the-exo-v3-module) - Microsoft

	> Currently, no cmdlets in Security & Compliance PowerShell cmdlets are backed by the REST API. All cmdlets in Security & Compliance PowerShell still rely on the remote PowerShell session, so PowerShell on your client computer requires Basic authentication in WinRM to successfully use the Connect-IPPSSession cmdlet.

	[Basic auth - Connect to Security & Compliance PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/basic-auth-connect-to-scc-powershell?view=exchange-ps) - Microsoft

	> WinRM needs to allow Basic authentication (it's enabled by default). We don't send the username and password combination, but the Basic authentication header is required to send the session's OAuth token, since the client-side WinRM implementation has no support for OAuth. 

### Usage

1. Install Microsoft eDiscovery Export Tool (ClickOnce App)

	- [Microsoft eDiscovery Export Tool](https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application) 


2. Install Prerequisite Modules

	```
	PS> Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop -Scope CurrentUser
	PS> Install-Module Figlet -AllowClobber -Force -ErrorAction Stop -Scope CurrentUser
	```

3. Create a new eDiscovery search within Purview Compliance Center

	- [Search for content in a eDiscovery (Standard) case](https://learn.microsoft.com/en-us/microsoft-365/compliance/ediscovery-search-for-content?source=recommendations&view=o365-worldwide)
	<p>

4. Launch eDiscovery-Export-Tool

	```
	PS> .\eDiscovery-Export-Tool.ps1
	```

5. Authenticate to Compliance Center and fill in basic job details

	- Search Name
	- Export Path
	- Region

	[Supported Regions](https://learn.microsoft.com/en-us/powershell/module/exchange/set-compliancesecurityfilter?view=exchange-ps#-region)

	| Name |            Region           |
	|------|-----------------------------|
	| APC  | Asia-Pacific                |
	| AUS  | Australia                   |
	| CAN  | Canada                      |
	| EUR  | Europe, Middle East, Africa |
	| FRA  | France                      |
	| GBR  | United Kingdon              |
	| IND  | India                       |
	| JPN  | Japan                       |
	| LAM  | Latin America               |
	| NAM  | North America               |
	<p>

	<img src="imgs/img1.png" style="border: 1px solid white">

6. Determine additional options

	- Include Unindexed Items
	- Deduplicate Items
	<p>

	<img src="imgs/img2.png" style="border: 1px solid white">

7. Review summary and create export job

	<img src="imgs/img3.png" style="border: 1px solid white">

8. Download results and review final export summary

	<img src="imgs/img4.png" style="border: 1px solid white">

### Disclaimer

This is a proof of concept script meant for testing purposes only. Use at your own risk.
