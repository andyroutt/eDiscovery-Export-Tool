<#
.Synopsis
    eDiscovery-Export-Tool - Powershell script to create and download eDiscovery searches.
.NOTES
    Version:          1.0
    Author:           Andy Routt
    Creation Date:    12/27/2022
    License:          MIT
#>

# Enable/Disable Debug Logging
# $GLOBAL:DebugPreference="Continue"            # Enable Debug

# Silence Module Installation
if ($GLOBAL:DebugPreference -eq "Continue"){
    Write-host "Debugging Enabled" -ForegroundColor Red
    $ProgressPreference = "Continue"            # Show Module Warnings
} else {
    $ProgressPreference = "SilentlyContinue"    # Hide Module Warnings
    Clear-Host
}

# Set Export Parameters
$Format = "FxStream"
$SharePointArchiveFormat = "SingleZip"
$ExchangeArchiveFormat = "PerUserPst"
$Dedupe = $false

# Create Region Table
$table = New-Object System.Data.Datatable
[void]$table.Columns.Add("Name")
[void]$table.Columns.Add("Description")
[void]$table.Rows.Add("APC","Asia-Pacific")
[void]$table.Rows.Add("AUS","Australia")
[void]$table.Rows.Add("CAN","Canada")
[void]$table.Rows.Add("EUR","Europe, Middle East, Africa")
[void]$table.Rows.Add("FRA","France")
[void]$table.Rows.Add("GBR","United Kingdon")
[void]$table.Rows.Add("IND","India")
[void]$table.Rows.Add("JPN","Japan")
[void]$table.Rows.Add("LAM","Latin America")
[void]$table.Rows.Add("NAM","North America")

# Install ExchangeOnlineManagement
if(!(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement)){
    Write-Debug "Installing ExchangeOnlineManagement Module"
    Write-Debug ""
    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop -Scope CurrentUser 
    if (!(Get-Module ExchangeOnlineManagement -ListAvailable)){
        Write-Host "ERROR! ExchangeOnlineManagement Module Installation Failed!" -ForegroundColor Red
        break
    }
} else {
    write-debug "ExchangeOnlineManagement Module Loaded"
    Write-Debug ""
}

# Enable Basic Authentication
$reg = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic
If ($reg -ne 1){
    Write-Debug "Enabling Basic Auth"
    Write-Debug ""
    Start-Process powershell -Verb runAs -ArgumentList "-NoProfile -Command `"Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic -Type Dword -Value 1`""
    if ((Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic) -lt "1"){
        Write-Host "ERROR! Enabling Basic Auth Failed!" -ForegroundColor Red
        break
    }
} else {
    Write-Debug "Basic Auth Enabled"
    Write-Debug ""
}

# Locate Export Tool
$UnifiedExportTool = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)
if (Test-Path $UnifiedExportTool){
    Write-Debug "Unified Export Tool Found"
    Write-Debug ""
} else {
    write-host "ERROR! Unified Export Tool Not Found!" -ForegroundColor Red
    break
}

# Load Shlwapi for StrFormatByteSize function
$Shlwapi = Add-Type -MemberDefinition '
    [DllImport("Shlwapi.dll", CharSet=CharSet.Auto)]public static extern int StrFormatByteSize(long fileSize, System.Text.StringBuilder pwszBuff, int cchBuff);
' -Name "ShlwapiFunctions" -namespace ShlwapiFunctions -PassThru

# Install Figlet Module
if(!(Get-Module Figlet)){
    Write-Debug "Installing Figlet Module"
    Write-Debug ""
    Install-Module Figlet -AllowClobber -Force -ErrorAction Stop -Scope CurrentUser
    if (!(Get-Module Figlet -ListAvailable)){
        Write-Host "ERROR! Figlet Module Installation Failed!" -ForegroundColor Red
        break
    }
} else {
    write-debug "Figlet Module Loaded"
    Write-Debug ""
}

#############
# Functions
#############

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

# Function - Format String as ByteSize
Function Format-ByteSize([Long]$Size){
    $Bytes = New-Object Text.StringBuilder 20
    $Return = $Shlwapi::StrFormatByteSize($Size, $Bytes, $Bytes.Capacity)
    If ($Return) {$Bytes.ToString()}
}

# Function - Connect to CC
Function Connect-ComplianceCenter(){

    # Connect to Compliance Center
    if (!((get-psSession).ComputerName -like "*ps.compliance.protection.outlook.com")){
        write-Debug "Connecting to Compliance Center"
        Write-Debug ""
        Import-Module ExchangeOnlineManagement
        Connect-IPPSSession -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        if (!((get-psSession).ComputerName -like "*ps.compliance.protection.outlook.com")){
            Write-Host "ERROR! Unable to Establish Compliance Center Connection!" -ForegroundColor Red
            break
        }
    } else {
        Write-Debug "Compliance Center Connectioin Established"
        Write-Debug ""
    }
}

# Function - Create Export Job
Function Create-ExportJob($SearchName, $Region, $Format, $SharePointArchiveFormat, $ExchangeArchiveFormat, $Dedupe){

    if ($search = Get-ComplianceSearch -identity $SearchName -erroraction SilentlyContinue){

        # Wait for Search to Complete
        if ($search."Status" -ne "Completed"){
            write-host "Waiting for Search to Finish ..." -NoNewline 
            Start-Sleep -s 1
            while ($search."Status" -ne "Completed"){
                write-host "." -NoNewline
                Start-Sleep -s 5
                $search = Get-ComplianceSearch -identity $SearchName -erroraction SilentlyContinue
            }
            write-host " Done!"
        }

        # Create New Export
        Write-Debug "Creating Export Job"
        Write-Debug ""
        New-ComplianceSearchAction `
            -SearchName $SearchName `
            -Region $Region `
            -Format $Format `
            -SharePointArchiveFormat $SharePointArchiveFormat `
            -ExchangeArchiveFormat $ExchangeArchiveFormat `
            -EnableDedupe $Dedupe `
            -Export
            | Out-Null
    } else {
        write-host "No Export Job Found ..." -ForegroundColor Red
        write-host ""
        break
    }    
}

# Function - Create Export Job
Function Show-ExportJob($SearchName){

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
            $scope = $y[3].trimStart(" Scope: ")
            $ei_format = $y[8].trimStart(" Exchange item format: ")
            $ea_format = $y[9].trimStart(" Exchange archive format: ")
            $sa_format = $y[10].trimStart(" SharePoint archive format ")
            $edd = $y[12].trimStart(" Enable dedupe: ")
            $region = $y[14].trimStart(" Region: ")
            $s_sources = $y[15].trimStart(" Started sources: ")
            $f_sources = $y[17].trimStart(" Failed sources: ")
            $e_size = $y[18].trimStart(" Total estimated bytes: ")
            $e_items = $y[19].trimStart(" Total estimated items: ")

            $sa_format_fixed = $sa_format.trimStart(": ")
            $casename = $export."CaseName"
            $exportname = $export."Name"
            $createdby = $export."CreatedBy"
            $size = Format-ByteSize($e_size)

            Switch ($ei_format)
            {
                "FxStream"  {$itemformat = "PST"}
                "Mime"      {$itemformat = "EML"}
                "Msg"       {$itemformat = "MSG"} 
            }

            Switch ($edd)
            {
                "True"      {$dedupestatus = "Yes"}
                "False"     {$dedupestatus = "No"}
            }

            write-host ""
            write-host "Case Name: $casename"
            write-host "Export Name: $exportname"
            write-host "Created By: $createdby"
            write-host ""
            write-host "Total Sources: $s_sources"
            write-host "Failed Sources: $f_sources"
            write-host ""
            write-host "Scope: $scope"
            write-host "Region: $region"
            write-host "Total Items: $e_items"
            write-host "Total Size: $size"
            write-host ""
            write-host "Dedupe: $dedupestatus"
            write-host "Mail Item Format: $itemformat"
            write-host "Mail Archive: $ea_format"
            write-host "SharePoint Archive: $sa_format_fixed"
            Write-host ""

            Write-Debug "Blob URL: $bloburl"
            Write-Debug "SAS Key: $sastoken"
        }

    } else {
        write-host "No Export Job Found ..." -ForegroundColor Red
        write-host ""
        break
    }
}

# Function - Create Export Search
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
    write-host "Downloading ..." -NoNewline
    Start-Sleep -s 1
    if ($proc){
        while(Get-Process -Name "microsoft.office.client.discovery.unifiedexporttool" -ErrorAction SilentlyContinue){
            write-host "." -NoNewline
            Start-Sleep -s 5
        }
    }
    write-host " Done!"
    write-host
}

#############
# User Input
#############

# Show Menu
Import-Module Figlet
write-figlet "eDiscovery Export Tool" -Foreground Green

# Prompt for input
$SearchName = Read-Host -Prompt "Enter Name of Search"
$ExportPath = Read-Host -Prompt "Export Path"
$Region = Read-Host -Prompt "Region"
$RegionName = Get-RegionName($Region)

write-host "Region Name: $RegionName"

# Check Validate Region Code
if (!($RegionName)){
    write-host ""
    write-host "Incorrect Region" -ForegroundColor Red
    $table | ft
    $Region = Read-Host -Prompt "Try Again"
}

# Establish Compliance Center Connection
Connect-ComplianceCenter

# Create Export Job
Create-ExportJob $SearchName $Region $Format $SharePointArchiveFormat $ExchangeArchiveFormat $Dedupe

# Show Export Job Detail
Show-ExportJob $SearchName

# Prompt for Download
$prompt = Read-Host -Prompt "Download (Y/N): "
if ($prompt -eq "Y"){
    Get-ExportData $ExportPath $SearchName $UnifiedExportTool
} else {
    write-host "Abort"
    break
}
