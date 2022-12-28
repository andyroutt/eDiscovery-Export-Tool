# eDiscovery-Export-Tool

Powershell script to create and download eDiscovery searches from Purview Compliance Portal. Supports assigning export jobs by region and downloading via the command line.

### Prerequisites

- [Microsoft eDiscovery Export Tool](https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application)
- [ExchangeOnlineManagement Module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/)
- [Figlet Module](https://www.powershellgallery.com/packages/Figlet/)
- Existing eDiscovery Search

### Usage

1. Create a new eDiscovery search within Purview Compliance Center 
	&rarr; [Search for content in a eDiscovery (Standard) case](https://learn.microsoft.com/en-us/microsoft-365/compliance/ediscovery-search-for-content?source=recommendations&view=o365-worldwide)
	<p>

2. Launch eDiscovery-Export-Tool

```
PS> eDiscovery-Export-Tool.ps1
```

3. Enter Search Name, Export Path, and Region

<img src="imgs/image-1.jpg" style="border: 1px solid white">

4. Accept prompt to download

<img src="imgs/image-2.jpg" style="border: 1px solid white">

### Disclaimer

These are proof of concept scripts. Use at your own risk.
