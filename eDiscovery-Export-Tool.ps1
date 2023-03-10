<#
.Synopsis
    eDiscovery-Export-Tool - Powershell script to create and download eDiscovery searches.
.NOTES
    Version:          1.1
    Author:           Andy Routt
    Creation Date:    12/27/2022
    License:          MIT
#>

# Enable/Disable Debug Logging
# $GLOBAL:DebugPreference="Continue"            # Enable Debug

# Load Modules
Import-Module ExchangeOnlineManagement
Import-Module Figlet

# Show Menu
Clear-Host
write-figlet "eDiscovery Export Tool" -Foreground Green

# Silence Module Installation
if ($GLOBAL:DebugPreference -eq "Continue"){
    Write-host "Debugging Enabled" -ForegroundColor Red
}

# Set Export Parameters
$Format = "FxStream"
$SharePointArchiveFormat = "SingleZip"
$ExchangeArchiveFormat = "PerUserPst"

# Create Region Table
$regiontable = New-Object System.Data.Datatable
[void]$regiontable.Columns.Add("Name")
[void]$regiontable.Columns.Add("Description")
[void]$regiontable.Rows.Add("APC","Asia-Pacific")
[void]$regiontable.Rows.Add("AUS","Australia")
[void]$regiontable.Rows.Add("CAN","Canada")
[void]$regiontable.Rows.Add("EUR","Europe, Middle East, Africa")
[void]$regiontable.Rows.Add("FRA","France")
[void]$regiontable.Rows.Add("GBR","United Kingdon")
[void]$regiontable.Rows.Add("IND","India")
[void]$regiontable.Rows.Add("JPN","Japan")
[void]$regiontable.Rows.Add("LAM","Latin America")
[void]$regiontable.Rows.Add("NAM","North America")

# Enable Basic Authentication
$reg = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic
If ($reg -ne 1){
    write-host "Enabling Basic Auth" -ForegroundColor Green
    Start-Process powershell -Verb runAs -ArgumentList "-NoProfile -Command `"Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic -Type Dword -Value 1`""
    Start-Sleep -s 5
    if ((Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic) -lt "1"){
        write-host ""
        Write-Host "ERROR! Enabling Basic Auth Failed!" -ForegroundColor Red
        write-host ""
        break
    }
} else {
    Write-host "Basic Auth Enabled" -ForegroundColor Green
}

# Locate Export Tool
$UnifiedExportTool = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)
if (Test-Path $UnifiedExportTool){
    Write-host "Unified Export Tool Found" -ForegroundColor Green
} else {
    write-host ""
    write-host "ERROR! Unified Export Tool Not Found!" -ForegroundColor Red
    write-host ""
    break
}

# Load Shlwapi for StrFormatByteSize function
$Shlwapi = Add-Type -MemberDefinition '
    [DllImport("Shlwapi.dll", CharSet=CharSet.Auto)]public static extern int StrFormatByteSize(long fileSize, System.Text.StringBuilder pwszBuff, int cchBuff);
' -Name "ShlwapiFunctions" -namespace ShlwapiFunctions -PassThru

#############
# Functions
#############

# Show Friendly Region Name
Function Get-RegionName($Region){
    Switch ($Region)
    {
        "APC" {$regiontext = "Asia-Pacific"; continue}
        "AUS" {$regiontext = "Australia"; continue}
        "CAN" {$regiontext = "Canada"; continue}
        "EUR" {$regiontext = "Europe, Middle East, Africa"; continue}
        "FRA" {$regiontext = "France"; continue}
        "GBR" {$regiontext = "United Kingdon"; continue}
        "IND" {$regiontext = "India"; continue}
        "JPN" {$regiontext = "Japan"; continue}
        "LAM" {$regiontext = "Latin America"; continue}
        "NAM" {$regiontext = "North America"; continue}
        default {$regiontext = ""}
    }
    return $regiontext
}

# Format String as ByteSize
Function Format-ByteSize([Long]$Size){
    $Bytes = New-Object Text.StringBuilder 20
    $Return = $Shlwapi::StrFormatByteSize($Size, $Bytes, $Bytes.Capacity)
    If ($Return) {$Bytes.ToString()}
}

# Show Friendly Format Name
Function Get-MailFormat($Format){
    Switch ($Format)
    {
        "FxStream"  {$itemformat = "PST"}
        "Mime"      {$itemformat = "EML"}
        "Msg"       {$itemformat = "MSG"}
    }
    return $itemformat
}

# Show Friendly Dedupe Name
Function Get-Dedupe($Dedupe){
    Switch ($Dedupe)
    {
        "True"      {$dedupestatus = "Yes"}
        "False"     {$dedupestatus = "No"}
    }
    return $dedupestatus
}

# Connect to Compliance Center
Function Connect-ComplianceCenter(){
    if (!((get-psSession).ComputerName -like "*ps.compliance.protection.outlook.com")){
        write-host "Establishing Connection to Compliance Center" -ForegroundColor Grey
        Connect-IPPSSession -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        if (!((get-psSession).ComputerName -like "*ps.compliance.protection.outlook.com")){
            Write-Host "ERROR! Unable to Establish Compliance Center Connection!" -ForegroundColor Red
            break
        } else {
            Write-host "Compliance Center Connectioin Established" -ForegroundColor Green
        }
    } else {
        Write-host "Compliance Center Connectioin Established" -ForegroundColor Green
    }
}

# Create Export Job
Function Create-ExportJob($SearchName, $Region, $Format, $SharePointArchiveFormat, $ExchangeArchiveFormat, $Scope, $Dedupe){

    # Create New Export
    if ($search = Get-ComplianceSearch -identity $SearchName -erroraction SilentlyContinue){
        New-ComplianceSearchAction `
        -SearchName $SearchName `
        -Region $Region `
        -Format $Format `
        -SharePointArchiveFormat $SharePointArchiveFormat `
        -ExchangeArchiveFormat $ExchangeArchiveFormat `
        -Scope $Scope `
        -EnableDedupe $Dedupe `
        -Export
        | Out-Null

        # Wait for Export Job to Complete
        $JobName = $SearchName+"_Export"
        $job = Get-ComplianceSearchAction -Identity $JobName -includeCredential -erroraction SilentlyContinue
        if ($job."Status" -ne "Completed"){
            write-host ""
            write-host "Creating Export Job ." -NoNewline -ForegroundColor Green
            Start-Sleep -s 1
            while ($job."Status" -ne "Completed"){
                write-host "." -NoNewline -ForegroundColor Green
                Start-Sleep -s 5
                $job = Get-ComplianceSearchAction -Identity $JobName -includeCredential -erroraction SilentlyContinue
            }
            write-host " Done!" -ForegroundColor Green
            write-host ""
        }

    } else {
        write-host ""
        write-host "ERROR! Search Not Found!" -ForegroundColor Red
        write-host ""
        break
    }
}

# Show Export Job
Function Show-NewExportJob($SearchName, $Region, $Format, $SharePointArchiveFormat, $ExchangeArchiveFormat, $Scope, $Dedupe){

    $regionname = Get-RegionName $Region
    $itemformat = Get-MailFormat $Format
    $dedupestatus = Get-Dedupe $Dedupe

    write-host ""
    write-host "Search Name: $SearchName" -ForegroundColor Yellow
    write-host "Region Name: $regionname" -ForegroundColor Yellow
    write-host ""
    write-host "Mail Format: $itemformat" -ForegroundColor Yellow
    write-host "Exchange Output: $ExchangeArchiveFormat" -ForegroundColor Yellow
    write-host "SharePoint Output: $SharePointArchiveFormat" -ForegroundColor Yellow
    write-host "Scope: $Scope" -ForegroundColor Yellow
    write-host "Dedupe: $dedupestatus" -ForegroundColor Yellow
    write-host ""
}

# Show Export Job
Function Show-FinishedExportJob($SearchName){

    $JobName = $SearchName+"_Export"
    if ($export = Get-ComplianceSearchAction -Identity $JobName -includeCredential -erroraction SilentlyContinue){

        # Show Export Detail
        if ($GLOBAL:DebugPreference -eq "Continue"){
            write-host ""
            $export | fl | out-string
        } else {
            $y = $export.Results.split(";")
            $bloburl = $y[0].trimStart("Container url: ")
            $sastoken = $y[1].trimStart(" SAS token: ")
            $s_sources = $y[15].trimStart(" Started sources: ")
            $f_sources = $y[17].trimStart(" Failed sources: ")
            $e_size = $y[18].trimStart(" Total estimated bytes: ")
            $e_items = $y[19].trimStart(" Total estimated items: ")

            $casename = $export."CaseName"
            $exportname = $export."Name"
            $size = Format-ByteSize $e_size

            write-host ""
            write-host "Case Name: $casename" -ForegroundColor Blue
            write-host "Export Job: $exportname" -ForegroundColor Blue
            write-host ""
            write-host "Total Sources: $s_sources" -ForegroundColor Blue
            write-host "Failed Sources: $f_sources" -ForegroundColor Blue
            write-host ""
            write-host "Total Items: $e_items" -ForegroundColor Blue
            write-host "Total Size: $size" -ForegroundColor Blue
            write-host ""
            Write-Debug "Blob URL: $bloburl"
            Write-Debug "SAS Key: $sastoken"
            Write-Debug ""

        }

    } else {
        write-host "No Export Job Found ..." -ForegroundColor Red
        write-host ""
        break
    }
}

# Retrieve Export Data
Function Get-ExportData($ExportPath, $SearchName, $UnifiedExportTool){

    $JobName = $SearchName+"_Export"
    if (Get-ComplianceSearchAction -identity $JobName -erroraction SilentlyContinue){  
    } else {
        write-host "No Export Job Found ..." -ForegroundColor Red
        write-host ""
        break
    }

    # Retrieve Search
    $index = Get-ComplianceSearchAction -Identity $JobName -includeCredential

    # Retrieve Blob URL and SAS token
    $y = $index.Results.split(";")
    $bloburl = $y[0].trimStart("Container url: ")
    $sastoken = $y[1].trimStart(" SAS token: ")

    # Configure Logs
    $t_log = $ExportPath+"\"+$JobName+"\Log.txt"
    $e_log = $ExportPath+"\"+$JobName+"\Errorlog.txt"

    # UnifiedExportTool Arguments
    $arguments = "-name `"$JobName`" -source `"$bloburl`" -key `"$sastoken`" -dest `"$ExportPath`" -trace `"$t_log`""

    # Download Export Data
    $downLoadProcess = Start-Process -FilePath $UnifiedExportTool -ArgumentList $arguments -Windowstyle Normal -RedirectStandardError $e_log -PassThru
    $proc = Get-Process | where -Property name -EQ "microsoft.office.client.discovery.unifiedexporttool"
    write-host ""
    write-host "Downloading ." -NoNewline -ForegroundColor Green
    
    # Wait for Download to Finish
    Start-Sleep -s 1
    if ($proc){
        while(Get-Process -Name "microsoft.office.client.discovery.unifiedexporttool" -ErrorAction SilentlyContinue){
            write-host "." -NoNewline -ForegroundColor Grey
            Start-Sleep -s 5
        }
    }
    write-host " Done!" -ForegroundColor Green
    write-host ""
}

########
# Main
########

# Establish Compliance Center Connection
Connect-ComplianceCenter

# Prompt for input
$SearchName = Read-Host -Prompt "Enter Name of Search"
$ExportPath = Read-Host -Prompt "Export Path"
$Region = Read-Host -Prompt "Region"

# Check Validate Region Code
$regionname = Get-RegionName $region
while (!($regionname)){
    write-host "ERROR! Incorrect Region Code!" -ForegroundColor Red
    $regiontable | ft
    $Region = Read-Host -Prompt "Try Again"
    $regionname = Get-RegionName $region
}

write-host ""

# Set Scope
$title = ""
$message = "Include Unindexed Items?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Include Unindexed Items"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do Not Include Unindexed Items"
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 1)

switch ($result)
{
    0 { $scope = "BothIndexedAndUnindexedItems" }
    1 { $scope = "IndexedItemsOnly" }
    2 { return; break }
}

# Set Dedupe
$title = ""
$message = "Deduplicate Items?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Deduplicate Items"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do Not Deduplicate Items"
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 1)

switch ($result)
{
    0 { $Dedupe = $true }
    1 { $Dedupe = $false }
    2 { return; break }
}

# Create Export Job
Show-NewExportJob $SearchName $Region $Format $SharePointArchiveFormat $ExchangeArchiveFormat $Scope $Dedupe

# Prompt for Creating Job
$title = ""
$message = "Create Export Job?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Create Export Job"
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 0)

switch ($result)
{
    0 { Create-ExportJob $SearchName $Region $Format $SharePointArchiveFormat $ExchangeArchiveFormat $Scope $Dedupe }
    1 { return; break }
}

# Prompt for Download
$title = ""
$message = "Download Data?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Download Data"
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 0)

switch ($result)
{
    0 { Get-ExportData $ExportPath $SearchName $UnifiedExportTool }
    1 { return; break }
}

# Show Export Job Detail
Show-FinishedExportJob $SearchName
