<#
.SYNOPSIS
    Jamf Managed Software Update Plan Analysis with Extension Attribute Support and User Details
    
.DESCRIPTION
    This PowerShell script analyzes Jamf Pro Managed Software Update plans and optionally creates/updates 
    Extension Attributes for advanced Smart Group management. It combines computer inventory data with 
    update plan details to provide comprehensive reporting and automated status tracking.
    
    The script handles scenarios where Managed Software Updates is disabled and provides fallback 
    computer inventory reporting. It always includes computer details (name, serial, model, OS version, 
    last contact, and user information) in analysis and CSV exports.

.PARAMETER ConfigFile
    Path to JSON configuration file containing Jamf Pro connection details
    Default: .\jamf_config.json

.PARAMETER ExportToCSV
    Export detailed results and summary statistics to timestamped CSV files

.PARAMETER WriteToEA
    Write update plan status and details to Extension Attributes for Smart Group creation

.PARAMETER CreateEA
    Create Extension Attributes if they don't exist (requires WriteToEA parameter)

.PARAMETER OutputPath
    Directory path for CSV exports
    Default: Current directory

.PARAMETER Unattended
    Run silently without interactive output (suitable for scheduled tasks)

.PARAMETER Help
    Display detailed help information and examples

.NOTES
    Version: 1
    Author: Carl Flanagan
    Date: 25 June 2025

    
    This script handles scenarios where Managed Software Updates is disabled,
    providing alternative computer inventory reporting and ensuring Extension Attributes  
    always reflect the current state (including "No Plan" when appropriate).
    
    The script always includes computer details and user information in analysis and CSV exports.
    
    Note: Extension Attributes are only supported for computers, not mobile devices due to API limitations.
#>

param(
    [string]$ConfigFile = ".\jamf_config.json",
    [switch]$ExportToCSV,
    [switch]$WriteToEA,
    [switch]$CreateEA,
    [string]$OutputPath = ".",
    [switch]$Help,
    [switch]$Unattended
)

# Global variables
$script:AccessToken = $null
$script:TokenExpiry = $null
$script:config = $null
$script:LogFile = $null
$script:StartTime = Get-Date
$script:SessionStats = @{
    DevicesProcessed = 0
    ComputersUpdated = 0
    MobileDevicesProcessed = 0
    ExtensionAttributesCreated = 0
    ExtensionAttributeUpdateErrors = 0
    CSVFilesCreated = 0
    StatusCounts = @{}
}

function Write-LogMessage {
    param([string]$Message, [string]$Level = "INFO", [ConsoleColor]$ForegroundColor = "White")
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to log file if available
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $LogEntry -ErrorAction SilentlyContinue
    }
    
    # Write to console unless unattended
    if (-not $Unattended) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
}

function Initialize-LogFile {
    param([string]$OutputPath)
    
    try {
        # Create logs subdirectory if it doesn't exist
        $LogsPath = Join-Path $OutputPath "logs"
        if (-not (Test-Path $LogsPath)) {
            New-Item -ItemType Directory -Path $LogsPath -Force | Out-Null
        }
        
        $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $script:LogFile = Join-Path $LogsPath "JamfMSU-Analysis-$Timestamp.log"
        
        # Create log file with header
        $LogHeader = @"
================================================================================
Jamf Managed Software Update Analysis Log
================================================================================
Start Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Computer: $env:COMPUTERNAME
User Context: $env:USERNAME
Parameters: ConfigFile=$ConfigFile, ExportToCSV=$ExportToCSV, WriteToEA=$WriteToEA, CreateEA=$CreateEA, OutputPath=$OutputPath, Unattended=$Unattended
================================================================================

"@
        Set-Content -Path $script:LogFile -Value $LogHeader -ErrorAction Stop
        
        Write-LogMessage "Log file initialized: $script:LogFile" "INFO" "Green"
        return $true
    }
    catch {
        Write-Host "Warning: Could not create log file: $($_.Exception.Message)" -ForegroundColor Yellow
        $script:LogFile = $null
        return $false
    }
}

function Write-SessionSummary {
    $EndTime = Get-Date
    $Duration = $EndTime - $script:StartTime
    
    $Summary = @"

================================================================================
SESSION SUMMARY
================================================================================
Start Time: $($script:StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
End Time: $($(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Duration: $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s

DEVICES PROCESSED:
- Total Devices: $($script:SessionStats.DevicesProcessed)
- Computers with EAs Updated: $($script:SessionStats.ComputersUpdated)
- Mobile Devices Processed: $($script:SessionStats.MobileDevicesProcessed)

EXTENSION ATTRIBUTES:
- Extension Attributes Created: $($script:SessionStats.ExtensionAttributesCreated)
- EA Update Errors: $($script:SessionStats.ExtensionAttributeUpdateErrors)

FILES CREATED:
- CSV Files Generated: $($script:SessionStats.CSVFilesCreated)

STATUS DISTRIBUTION:
"@

    foreach ($Status in $script:SessionStats.StatusCounts.Keys | Sort-Object) {
        $Count = $script:SessionStats.StatusCounts[$Status]
        $Summary += "`n- $Status`: $Count devices"
    }

    $Summary += @"

LOG FILE LOCATION: $script:LogFile
================================================================================
"@
    
    Write-LogMessage $Summary "SUMMARY" "Cyan"
    
    # Also write summary to console even in unattended mode
    if ($Unattended -and $script:LogFile) {
        Write-Host "`nSession completed. Log file: $script:LogFile" -ForegroundColor Green
        Write-Host "Duration: $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s | Devices: $($script:SessionStats.DevicesProcessed) | EA Updates: $($script:SessionStats.ComputersUpdated)" -ForegroundColor Gray
    }
}

if ($Help) {
    Write-Host "Jamf Managed Software Update Plan Analysis Script" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\script.ps1 [parameters]" -ForegroundColor White
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Yellow
    Write-Host "  -ConfigFile <path>                Path to JSON configuration file" -ForegroundColor White
    Write-Host "  -ExportToCSV                      Export results to CSV files with user details" -ForegroundColor White
    Write-Host "  -WriteToEA                        Write plan details to Extension Attributes" -ForegroundColor White
    Write-Host "  -CreateEA                         Create EAs if they don't exist" -ForegroundColor White
    Write-Host "  -OutputPath <path>                Path for CSV exports" -ForegroundColor White
    Write-Host "  -Unattended                       Run silently" -ForegroundColor White
    Write-Host ""
    Write-Host "Extension Attributes Created (Computers Only):" -ForegroundColor Magenta
    Write-Host "  - MSU_Plan_Status                 (PlanCompleted/PlanFailed/PlanInProgress/No Plan)" -ForegroundColor White
    Write-Host "  - MSU_Plan_Action                 (DOWNLOAD_INSTALL_SCHEDULE/etc)" -ForegroundColor White
    Write-Host "  - MSU_Plan_Version_Type           (LATEST_MAJOR/LATEST_MINOR/etc)" -ForegroundColor White
    Write-Host "  - MSU_Plan_Error_Reasons          (Error details or 'No Errors')" -ForegroundColor White
    Write-Host "  - MSU_Plan_Force_Install_Date     (Scheduled date or 'Not Set')" -ForegroundColor White
    Write-Host ""
    Write-Host "Note:" -ForegroundColor Yellow
    Write-Host "  Mobile devices don't support Extension Attributes but are included in CSV exports." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\script.ps1 -ExportToCSV" -ForegroundColor Gray
    Write-Host "  .\script.ps1 -WriteToEA -CreateEA" -ForegroundColor Gray
    exit 0
}

function Test-ConfigFile {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-LogMessage "Configuration file not found: $Path" "ERROR" "Red"
        Write-LogMessage "Creating sample configuration file..." "INFO" "Yellow"
        
        $sampleConfig = @{
            jamf_url = "https://your-jamf-instance.jamfcloud.com"
            client_id = "YOUR_CLIENT_ID_HERE"
            client_secret = "YOUR_CLIENT_SECRET_HERE"
        } | ConvertTo-Json -Depth 2
        
        $sampleConfig | Out-File -FilePath $Path -Encoding UTF8
        Write-LogMessage "Sample configuration created. Edit with your Jamf Pro OAuth details." "INFO" "Green"
        Write-LogMessage "" "INFO"
        Write-LogMessage "Required API Permissions:" "INFO" "Cyan"
        Write-LogMessage "  - Read Managed Software Updates" "INFO" "White"
        Write-LogMessage "  - Read Computers" "INFO" "White"
        Write-LogMessage "  - Read Computer Extension Attributes" "INFO" "White"
        Write-LogMessage "  - Create Computer Extension Attributes" "INFO" "White"
        Write-LogMessage "  - Update Computer Extension Attributes" "INFO" "White"
        Write-LogMessage "  - Update Computers" "INFO" "White"
        Write-LogMessage "  - Read Mobile Devices" "INFO" "White"
        Write-LogMessage "" "INFO"
        Write-LogMessage "Setup: Jamf Pro > Settings > User Accounts & Groups > API Roles and Clients" "INFO" "Gray"
        return $false
    }
    return $true
}

function Get-JamfAccessToken {
    param([string]$BaseURL, [string]$ClientID, [string]$ClientSecret)
    
    try {
        $Headers = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
        $FormData = @{
            'client_id' = $ClientID
            'grant_type' = 'client_credentials'
            'client_secret' = $ClientSecret
        }
        
        $TokenURL = "$BaseURL/api/oauth/token"
        Write-LogMessage "Requesting access token..." "INFO" "Yellow"
        
        $Response = Invoke-RestMethod -Uri $TokenURL -Method Post -Headers $Headers -Body $FormData
        
        $script:AccessToken = $Response.access_token
        $script:TokenExpiry = (Get-Date).AddSeconds($Response.expires_in - 60)
        
        Write-LogMessage "Successfully obtained access token" "INFO" "Green"
        return $true
    }
    catch {
        Write-LogMessage "Error obtaining access token: $($_.Exception.Message)" "ERROR" "Red"
        return $false
    }
}

function Test-JamfTokenValidity {
    if (-not $script:AccessToken -or -not $script:TokenExpiry) { return $false }
    return (Get-Date) -lt $script:TokenExpiry
}

function Invoke-JamfAPIRequest {
    param([string]$Endpoint, [hashtable]$Parameters = @{})
    
    if (-not (Test-JamfTokenValidity)) {
        if (-not (Get-JamfAccessToken -BaseURL $script:config.jamf_url -ClientID $script:config.client_id -ClientSecret $script:config.client_secret)) {
            throw "Failed to obtain valid access token"
        }
    }
    
    $Headers = @{
        'Authorization' = "Bearer $script:AccessToken"
        'Accept' = 'application/json'
    }
    
    $URI = "$($script:config.jamf_url)$Endpoint"
    
    try {
        return Invoke-RestMethod -Uri $URI -Method Get -Headers $Headers -Body $Parameters
    }
    catch {
        # Check for specific errors related to Managed Software Updates
        if ($_.Exception.Response.StatusCode -eq 503 -and $Endpoint -like "*managed-software-updates*") {
            Write-LogMessage "" "INFO"
            Write-LogMessage "MANAGED SOFTWARE UPDATES NOT AVAILABLE" "ERROR" "Red"
            Write-LogMessage "========================================" "ERROR" "Red"
            Write-LogMessage "" "INFO"
            Write-LogMessage "The Managed Software Updates feature appears to be disabled in Jamf Pro." "INFO" "Yellow"
            Write-LogMessage "" "INFO"
            Write-LogMessage "Possible reasons:" "INFO" "Cyan"
            Write-LogMessage "  - The feature is not enabled in your Jamf Pro instance" "INFO" "White"
            Write-LogMessage "  - The feature may be temporarily unavailable" "INFO" "White"
            Write-LogMessage "  - Your Jamf Pro licence may not include this feature" "INFO" "White"
            Write-LogMessage "" "INFO"
            Write-LogMessage "To resolve this:" "INFO" "Cyan"
            Write-LogMessage "  1. Contact your Jamf Pro administrator" "INFO" "White"
            Write-LogMessage "  2. Verify Managed Software Updates is enabled in:" "INFO" "White"
            Write-LogMessage "     Computers -> Software Updates -> Enable/Disable Software Updates" "INFO" "Gray"
            Write-LogMessage "  3. Check your Jamf Pro licence includes this feature" "INFO" "White"
            Write-LogMessage "" "INFO"
            throw "Managed Software Updates feature is not available (HTTP 503)"
        }
        
        Write-LogMessage "API request failed: $($_.Exception.Message)" "ERROR" "Red"
        if ($_.Exception.Response) {
            $StatusCode = $_.Exception.Response.StatusCode.Value__
            Write-LogMessage "   HTTP Status: $StatusCode" "ERROR" "Red"
        }
        throw
    }
}

function Get-JamfManagedSoftwareUpdatePlans {
    param([int]$PageSize = 100)
    
    Write-LogMessage "Fetching managed software update plans..." "INFO" "Cyan"
    
    try {
        $AllPlans = @()
        $Page = 0
        
        do {
            $Parameters = @{
                'page' = $Page
                'page-size' = $PageSize
                'sort' = 'planUuid:asc'
            }
            
            $Response = Invoke-JamfAPIRequest -Endpoint '/api/v1/managed-software-updates/plans' -Parameters $Parameters
            $Plans = $Response.results
            
            if ($Plans.Count -gt 0) {
                $AllPlans += $Plans
                Write-LogMessage "   Retrieved $($Plans.Count) plans from page $($Page + 1)" "INFO" "Gray"
                $Page++
            }
        } while ($Plans.Count -eq $PageSize)
        
        Write-LogMessage "Total plans retrieved: $($AllPlans.Count)" "INFO" "Green"
        return $AllPlans
    }
    catch {
        if ($_.Exception.Message -like "*Managed Software Updates feature is not available*") {
            # Re-throw the formatted error from Invoke-JamfAPIRequest
            throw
        } else {
            Write-LogMessage "Failed to retrieve update plans: $($_.Exception.Message)" "ERROR" "Red"
            throw
        }
    }
}

function Get-AllComputers {
    try {
        Write-LogMessage "Retrieving all computers with detailed information..." "INFO" "Cyan"
        
        $AllComputers = @()
        $Page = 0
        $PageSize = 100
        
        do {
            $Parameters = @{
                'page' = $Page
                'page-size' = $PageSize
                'section' = 'GENERAL,OPERATING_SYSTEM,HARDWARE,USER_AND_LOCATION'
                'sort' = 'general.name:asc'
            }
            
            $Response = Invoke-JamfAPIRequest -Endpoint '/api/v1/computers-inventory' -Parameters $Parameters
            $Computers = $Response.results
            
            if ($Computers.Count -gt 0) {
                $AllComputers += $Computers
                Write-LogMessage "Retrieved $($Computers.Count) computers from page $($Page + 1)" "INFO" "Gray"
                $Page++
            }
        } while ($Computers.Count -eq $PageSize)
        
        Write-LogMessage "Total computers retrieved: $($AllComputers.Count)" "INFO" "Green"
        return $AllComputers
    }
    catch {
        Write-LogMessage "Failed to retrieve computers: $($_.Exception.Message)" "ERROR" "Red"
        throw
    }
}

function Get-AllMobileDevices {
    try {
        Write-LogMessage "Retrieving all mobile devices with detailed information..." "INFO" "Cyan"
        
        $AllMobileDevices = @()
        $Page = 0
        $PageSize = 100
        
        do {
            $Parameters = @{
                'page' = $Page
                'page-size' = $PageSize
                'section' = 'GENERAL,HARDWARE,USER_AND_LOCATION'
                'sort' = 'displayName:asc'
            }
            
            $Response = Invoke-JamfAPIRequest -Endpoint '/api/v2/mobile-devices/detail' -Parameters $Parameters
            $MobileDevices = $Response.results
            
            if ($MobileDevices.Count -gt 0) {
                $AllMobileDevices += $MobileDevices
                Write-LogMessage "Retrieved $($MobileDevices.Count) mobile devices from page $($Page + 1)" "INFO" "Gray"
                $Page++
            }
        } while ($MobileDevices.Count -eq $PageSize)
        
        Write-LogMessage "Total mobile devices retrieved: $($AllMobileDevices.Count)" "INFO" "Green"
        return $AllMobileDevices
    }
    catch {
        Write-LogMessage "Failed to retrieve mobile devices: $($_.Exception.Message)" "ERROR" "Red"
        throw
    }
}

function Get-AllDevices {
    try {
        Write-LogMessage "Retrieving all devices (computers and mobile devices)..." "INFO" "Cyan"
        
        # Get computers and mobile devices
        $AllComputers = Get-AllComputers
        $AllMobileDevices = Get-AllMobileDevices
        
        # Create unified device lookup by ID
        $DeviceLookup = @{}
        
        # Add computers to lookup
        foreach ($computer in $AllComputers) {
            $DeviceLookup[$computer.id] = @{
                Type = "Computer"
                Data = $computer
                Name = if ($computer.general -and $computer.general.name) { $computer.general.name } else { 'Unknown' }
                SerialNumber = if ($computer.hardware -and $computer.hardware.serialNumber) { $computer.hardware.serialNumber } else { 'Unknown' }
                Model = if ($computer.hardware -and $computer.hardware.model) { $computer.hardware.model } else { 'Unknown' }
                OSVersion = if ($computer.operatingSystem -and $computer.operatingSystem.version) { $computer.operatingSystem.version } else { 'Unknown' }
                LastContact = if ($computer.general -and $computer.general.lastContactTime) { $computer.general.lastContactTime } else { 'Unknown' }
                Username = if ($computer.userAndLocation -and $computer.userAndLocation.username) { $computer.userAndLocation.username } else { 'Unknown' }
                RealName = if ($computer.userAndLocation -and $computer.userAndLocation.realname) { $computer.userAndLocation.realname } else { 'Unknown' }
                Email = if ($computer.userAndLocation -and $computer.userAndLocation.email) { $computer.userAndLocation.email } else { 'Unknown' }
                Position = if ($computer.userAndLocation -and $computer.userAndLocation.position) { $computer.userAndLocation.position } else { 'Unknown' }
            }
        }
        
        # Add mobile devices to lookup
        foreach ($mobileDevice in $AllMobileDevices) {
            $DeviceLookup[$mobileDevice.mobileDeviceId] = @{
                Type = "Mobile Device"
                Data = $mobileDevice
                Name = if ($mobileDevice.general -and $mobileDevice.general.displayName) { $mobileDevice.general.displayName } else { 'Unknown' }
                SerialNumber = if ($mobileDevice.hardware -and $mobileDevice.hardware.serialNumber) { $mobileDevice.hardware.serialNumber } else { 'Unknown' }
                Model = if ($mobileDevice.hardware -and $mobileDevice.hardware.model) { $mobileDevice.hardware.model } else { 'Unknown' }
                OSVersion = if ($mobileDevice.general -and $mobileDevice.general.osVersion) { $mobileDevice.general.osVersion } else { 'Unknown' }
                LastContact = if ($mobileDevice.general -and $mobileDevice.general.lastInventoryUpdateDate) { $mobileDevice.general.lastInventoryUpdateDate } else { 'Unknown' }
                Username = if ($mobileDevice.userAndLocation -and $mobileDevice.userAndLocation.username) { $mobileDevice.userAndLocation.username } else { 'Unknown' }
                RealName = if ($mobileDevice.userAndLocation -and $mobileDevice.userAndLocation.realName) { $mobileDevice.userAndLocation.realName } else { 'Unknown' }
                Email = if ($mobileDevice.userAndLocation -and $mobileDevice.userAndLocation.email) { $mobileDevice.userAndLocation.email } else { 'Unknown' }
                Position = if ($mobileDevice.userAndLocation -and $mobileDevice.userAndLocation.position) { $mobileDevice.userAndLocation.position } else { 'Unknown' }
            }
        }
        
        Write-LogMessage "Total devices in lookup: $($DeviceLookup.Count) (Computers: $($AllComputers.Count), Mobile: $($AllMobileDevices.Count))" "INFO" "Green"
        
        return @{
            DeviceLookup = $DeviceLookup
            Computers = $AllComputers
            MobileDevices = $AllMobileDevices
        }
    }
    catch {
        Write-LogMessage "Failed to retrieve devices: $($_.Exception.Message)" "ERROR" "Red"
        throw
    }
}

function Get-JamfExtensionAttribute {
    param([string]$AttributeName)
    
    try {
        $Response = Invoke-JamfAPIRequest -Endpoint '/api/v1/computer-extension-attributes'
        $ExistingEA = $Response.results | Where-Object { $_.name -eq $AttributeName }
        
        if ($ExistingEA) {
            Write-LogMessage "Found Extension Attribute: $AttributeName (ID: $($ExistingEA.id))" "INFO" "Green"
            return $ExistingEA
        } else {
            Write-LogMessage "Extension Attribute '$AttributeName' not found" "INFO" "Yellow"
            return $null
        }
    }
    catch {
        Write-LogMessage "Error checking Extension Attribute: $($_.Exception.Message)" "ERROR" "Red"
        return $null
    }
}

function Get-AllRequiredExtensionAttributes {
    param([switch]$CreateIfMissing)
    
    $RequiredEAs = @(
        @{ Name = "MSU_Plan_Status"; Description = "Managed Software Update Plan Status (PlanCompleted/PlanFailed/PlanInProgress/No Plan)" },
        @{ Name = "MSU_Plan_Action"; Description = "Managed Software Update Plan Action (DOWNLOAD_INSTALL_SCHEDULE/etc)" },
        @{ Name = "MSU_Plan_Version_Type"; Description = "Managed Software Update Plan Version Type (LATEST_MAJOR/LATEST_MINOR/etc)" },
        @{ Name = "MSU_Plan_Error_Reasons"; Description = "Managed Software Update Plan Error Reasons" },
        @{ Name = "MSU_Plan_Force_Install_Date"; Description = "Managed Software Update Plan Force Install Date" }
    )
    
    $ExtensionAttributes = @{}
    
    foreach ($EADef in $RequiredEAs) {
        $EA = Get-JamfExtensionAttribute -AttributeName $EADef.Name
        
        if (-not $EA -and $CreateIfMissing) {
            $EA = New-JamfExtensionAttribute -AttributeName $EADef.Name -Description $EADef.Description
        } elseif (-not $EA) {
            throw "Extension Attribute '$($EADef.Name)' not found and -CreateExtensionAttribute not specified"
        }
        
        $ExtensionAttributes[$EADef.Name] = $EA
    }
    
    return $ExtensionAttributes
}

function New-JamfExtensionAttribute {
    param([string]$AttributeName, [string]$Description)
    
    try {
        Write-LogMessage "Creating Extension Attribute: $AttributeName..." "INFO" "Cyan"
        
        if (-not (Test-JamfTokenValidity)) {
            if (-not (Get-JamfAccessToken -BaseURL $script:config.jamf_url -ClientID $script:config.client_id -ClientSecret $script:config.client_secret)) {
                throw "Failed to obtain valid access token"
            }
        }
        
        $Headers = @{
            'Authorization' = "Bearer $script:AccessToken"
            'Accept' = 'application/json'
            'Content-Type' = 'application/json'
        }
        
        $Body = @{
            name = $AttributeName
            description = $Description
            dataType = "STRING"
            inputType = "TEXT"
            inventoryDisplayType = "OPERATING_SYSTEM"
            enabled = $true
        } | ConvertTo-Json -Depth 3
        
        $URI = "$($script:config.jamf_url)/api/v1/computer-extension-attributes"
        $Response = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -Body $Body
        
        Write-LogMessage "Successfully created Extension Attribute: $AttributeName (ID: $($Response.id))" "INFO" "Green"
        $script:SessionStats.ExtensionAttributesCreated++
        return $Response
    }
    catch {
        Write-LogMessage "Error creating Extension Attribute: $($_.Exception.Message)" "ERROR" "Red"
        throw
    }
}

function Set-ComputerExtensionAttributeValues {
    param(
        [string]$ComputerId, 
        [hashtable]$ExtensionAttributes,
        [hashtable]$Values
    )
    
    try {
        if (-not (Test-JamfTokenValidity)) {
            if (-not (Get-JamfAccessToken -BaseURL $script:config.jamf_url -ClientID $script:config.client_id -ClientSecret $script:config.client_secret)) {
                throw "Failed to obtain valid access token"
            }
        }
        
        $Headers = @{
            'Authorization' = "Bearer $script:AccessToken"
            'Accept' = 'application/json'
            'Content-Type' = 'application/json'
        }
        
        # Build extension attributes array
        $ExtensionAttributesArray = @()
        foreach ($EAName in $Values.Keys) {
            if ($ExtensionAttributes.ContainsKey($EAName)) {
                $ExtensionAttributesArray += @{
                    definitionId = $ExtensionAttributes[$EAName].id
                    values = @($Values[$EAName])
                }
            }
        }
        
        $Body = @{
            operatingSystem = @{
                extensionAttributes = $ExtensionAttributesArray
            }
        } | ConvertTo-Json -Depth 4
        
        $URI = "$($script:config.jamf_url)/api/v1/computers-inventory-detail/$ComputerId"
        Invoke-RestMethod -Uri $URI -Method Patch -Headers $Headers -Body $Body | Out-Null
        
        return $true
    }
    catch {
        Write-LogMessage "Error updating Extension Attributes for Computer ID $ComputerId : $($_.Exception.Message)" "ERROR" "Red"
        $script:SessionStats.ExtensionAttributeUpdateErrors++
        return $false
    }
}

function Update-ComputerExtensionAttributes {
    param([array]$UpdatePlans, [hashtable]$ExtensionAttributes)
    
    Write-LogMessage "Updating Extension Attributes for computers..." "INFO" "Cyan"
    
    $AllDevices = Get-AllDevices
    $DeviceLookup = $AllDevices.DeviceLookup
    
    # Filter to only computers since mobile devices don't support Extension Attributes
    $ComputerDevices = @{}
    foreach ($deviceId in $DeviceLookup.Keys) {
        $device = $DeviceLookup[$deviceId]
        if ($device.Type -eq "Computer") {
            $ComputerDevices[$deviceId] = $device
        }
    }
    
    Write-LogMessage "Processing $($ComputerDevices.Count) computers (excluding $($DeviceLookup.Count - $ComputerDevices.Count) mobile devices)" "INFO" "Yellow"
    
    # Create lookup of update plans by device ID
    $UpdatePlanLookup = @{}
    foreach ($plan in $UpdatePlans) {
        $deviceId = $plan.device.deviceId
        if ($deviceId) {
            $UpdatePlanLookup[$deviceId] = $plan
        }
    }
    
    $UpdatedCount = 0
    $ErrorCount = 0
    $TotalComputers = $ComputerDevices.Count
    $script:SessionStats.DevicesProcessed = $DeviceLookup.Count
    
    # Statistics tracking
    $StatusCounts = @{}
    
    foreach ($deviceId in $ComputerDevices.Keys) {
        $device = $ComputerDevices[$deviceId]
        $deviceName = $device.Name
        
        Write-Progress -Activity "Updating Extension Attributes" -Status "Processing $deviceName" -PercentComplete (($UpdatedCount / $TotalComputers) * 100)
        
        # Default values for devices without update plans
        $Values = @{
            "MSU_Plan_Status" = "No Plan"
            "MSU_Plan_Action" = "No Plan"
            "MSU_Plan_Version_Type" = "No Plan" 
            "MSU_Plan_Error_Reasons" = "No Plan"
            "MSU_Plan_Force_Install_Date" = "No Plan"
        }
        
        if ($UpdatePlanLookup.ContainsKey($deviceId)) {
            $updatePlan = $UpdatePlanLookup[$deviceId]
            
            # Use actual raw plan values instead of standardised ones
            $Values["MSU_Plan_Status"] = if ($updatePlan.status.state) { $updatePlan.status.state } else { "Unknown" }
            $Values["MSU_Plan_Action"] = if ($updatePlan.updateAction) { $updatePlan.updateAction } else { "Unknown" }
            $Values["MSU_Plan_Version_Type"] = if ($updatePlan.versionType) { $updatePlan.versionType } else { "Unknown" }
            $Values["MSU_Plan_Error_Reasons"] = if ($updatePlan.status.errorReasons -and $updatePlan.status.errorReasons.Count -gt 0) { 
                ($updatePlan.status.errorReasons -join ", ") 
            } else { 
                "No Errors" 
            }
            $Values["MSU_Plan_Force_Install_Date"] = if ($updatePlan.forceInstallLocalDateTime) { 
                $updatePlan.forceInstallLocalDateTime 
            } else { 
                "Not Set" 
            }
        }
        
        # Track statistics
        $statusValue = $Values["MSU_Plan_Status"]
        if ($StatusCounts.ContainsKey($statusValue)) {
            $StatusCounts[$statusValue]++
        } else {
            $StatusCounts[$statusValue] = 1
        }
        
        # Update Extension Attributes for this computer
        if (Set-ComputerExtensionAttributeValues -ComputerId $deviceId -ExtensionAttributes $ExtensionAttributes -Values $Values) {
            $UpdatedCount++
            $script:SessionStats.ComputersUpdated++
        } else {
            $ErrorCount++
        }
        
        Start-Sleep -Milliseconds 50
    }
    
    # Store status counts for summary (includes all devices for overall statistics)
    foreach ($deviceId in $DeviceLookup.Keys) {
        $device = $DeviceLookup[$deviceId]
        if ($device.Type -eq "Mobile Device") {
            $script:SessionStats.MobileDevicesProcessed++
        }
    }
    
    # Store status counts for summary
    $script:SessionStats.StatusCounts = $StatusCounts
    
    Write-Progress -Activity "Updating Extension Attributes" -Completed
    
    Write-LogMessage "" "INFO"
    Write-LogMessage "Extension Attribute Update Summary:" "INFO" "Green"
    Write-LogMessage "===================================" "INFO" "Green"
    Write-LogMessage "Total computers processed: $TotalComputers" "INFO" "White"
    Write-LogMessage "Successfully updated: $UpdatedCount" "INFO" "Green"
    Write-LogMessage "Errors: $ErrorCount" "INFO" "Red"
    Write-LogMessage "Mobile devices skipped: $($script:SessionStats.MobileDevicesProcessed)" "INFO" "Gray"
    
    # Show status distribution
    Write-LogMessage "" "INFO"
    Write-LogMessage "Status Value Distribution (Computers Only):" "INFO" "Yellow"
    foreach ($status in $StatusCounts.Keys | Sort-Object) {
        $count = $StatusCounts[$status]
        $percentage = [math]::Round(($count / $TotalComputers) * 100, 1)
        Write-LogMessage "  $status`: $count computers ($percentage%)" "INFO" "White"
    }
    
    # Only show Smart Group suggestions in console, not in log file
    if (-not $Unattended) {
        Write-LogMessage "" "INFO"
        Write-Host "Smart Group Suggestions:" -ForegroundColor Magenta
        Write-Host "========================" -ForegroundColor Magenta
        Write-Host "RAW STATUS-BASED SMART GROUPS:" -ForegroundColor Cyan
        Write-Host "  Extension Attribute 'MSU_Plan_Status' is 'PlanFailed' -> Failed updates" -ForegroundColor Red
        Write-Host "  Extension Attribute 'MSU_Plan_Status' is 'PlanCompleted' -> Completed updates" -ForegroundColor Green
        Write-Host "  Extension Attribute 'MSU_Plan_Status' is 'PlanInProgress' -> Currently updating" -ForegroundColor Yellow
        Write-Host "  Extension Attribute 'MSU_Plan_Status' is 'No Plan' -> Devices without update plans" -ForegroundColor Gray
        Write-Host "" -ForegroundColor White
        Write-Host "ACTION-BASED SMART GROUPS:" -ForegroundColor Cyan
        Write-Host "  Extension Attribute 'MSU_Plan_Action' is 'DOWNLOAD_INSTALL_SCHEDULE' -> Scheduled updates" -ForegroundColor White
        Write-Host "  Extension Attribute 'MSU_Plan_Version_Type' is 'LATEST_MAJOR' -> Major version updates" -ForegroundColor White
        Write-Host "  Extension Attribute 'MSU_Plan_Error_Reasons' contains 'SPECIFIC_VERSION_UNAVAILABLE' -> Version issues" -ForegroundColor Red
    }
}

function ConvertTo-UpdatePlanTable {
    param([array]$Plans)
    
    Write-LogMessage "Processing plan data with enhanced user details..." "INFO" "Cyan"
    
    # Always fetch all device details (computers and mobile devices)
    Write-LogMessage "Fetching all device details including user information..." "INFO" "Yellow"
    $AllDevices = Get-AllDevices
    $DeviceLookup = $AllDevices.DeviceLookup
    
    # Debug: Show sample device IDs and device IDs from plans
    if ($DeviceLookup.Count -gt 0) {
        $SampleDeviceIds = $DeviceLookup.Keys | Select-Object -First 3
        Write-LogMessage "Sample device IDs: $($SampleDeviceIds -join ', ')" "INFO" "Gray"
    }
    if ($Plans.Count -gt 0) {
        Write-LogMessage "Sample plan device IDs: $($Plans[0..2] | ForEach-Object { $_.device.deviceId })" "INFO" "Gray"
    }
    
    Write-LogMessage "Device lookup created with $($DeviceLookup.Count) devices" "INFO" "Gray"
    
    $TableData = @()
    $MatchedCount = 0
    
    foreach ($Plan in $Plans) {
        $Device = $Plan.device
        $Status = $Plan.status
        $DeviceId = $Device.deviceId
        
        # Debug: Show the device ID we're trying to match
        Write-LogMessage "Processing plan for device ID: '$DeviceId' (type: $($DeviceId.GetType().Name))" "INFO" "Gray"
        
        # Initialize with device detail fields including user information
        $Row = [PSCustomObject]@{
            'Plan UUID' = $Plan.planUuid
            'Device ID' = $DeviceId
            'Device Type' = 'Unknown'
            'Computer Name' = 'Unknown'
            'Serial Number' = 'Unknown'
            'Model' = 'Unknown'
            'OS Version' = 'Unknown'
            'Last Contact' = 'Unknown'
            'Username' = 'Unknown'
            'Real Name' = 'Unknown'
            'Email' = 'Unknown'
            'Position' = 'Unknown'
            'Update Action' = $Plan.updateAction
            'Version Type' = $Plan.versionType
            'Max Deferrals' = $Plan.maxDeferrals
            'Status' = $Status.state
            'Error Reasons' = ($Status.errorReasons -join ', ')
            'Force Install Date' = if ($Plan.forceInstallLocalDateTime) { $Plan.forceInstallLocalDateTime } else { 'Not Set' }
        }
        
        # Try multiple matching approaches
        $deviceInfo = $null
        
        # Method 1: Direct lookup
        if ($DeviceLookup.ContainsKey($DeviceId)) {
            $deviceInfo = $DeviceLookup[$DeviceId]
            Write-LogMessage "   Direct match found for $DeviceId" "INFO" "Green"
        }
        # Method 2: String conversion lookup
        elseif ($DeviceLookup.ContainsKey($DeviceId.ToString())) {
            $deviceInfo = $DeviceLookup[$DeviceId.ToString()]
            Write-LogMessage "   String match found for $DeviceId" "INFO" "Green"
        }
        # Method 3: Try finding by iterating (in case of type mismatch)
        else {
            foreach ($DevId in $DeviceLookup.Keys) {
                if ($DevId -eq $DeviceId -or $DevId.ToString() -eq $DeviceId.ToString()) {
                    $deviceInfo = $DeviceLookup[$DevId]
                    Write-LogMessage "   Iteration match found: '$DevId' matches '$DeviceId'" "INFO" "Green"
                    break
                }
            }
        }
        
        if ($deviceInfo) {
            $MatchedCount++
            $Row.'Device Type' = $deviceInfo.Type
            $Row.'Computer Name' = $deviceInfo.Name
            $Row.'Serial Number' = $deviceInfo.SerialNumber
            $Row.'Model' = $deviceInfo.Model
            $Row.'OS Version' = $deviceInfo.OSVersion
            $Row.'Last Contact' = $deviceInfo.LastContact
            $Row.'Username' = $deviceInfo.Username
            $Row.'Real Name' = $deviceInfo.RealName
            $Row.'Email' = $deviceInfo.Email
            $Row.'Position' = $deviceInfo.Position
            
            Write-LogMessage "   SUCCESS: $DeviceId -> $($deviceInfo.Type) | $($deviceInfo.Name) | $($deviceInfo.SerialNumber) | $($deviceInfo.Username) | $($deviceInfo.Email)" "INFO" "Cyan"
        } else {
            Write-LogMessage "   NO MATCH: Could not find device for device ID '$DeviceId'" "INFO" "Red"
        }
        
        $TableData += $Row
    }
    
    Write-LogMessage "Created table with $($TableData.Count) rows" "INFO" "Green"
    Write-LogMessage "Successfully matched $MatchedCount out of $($Plans.Count) devices" "INFO" "Yellow"
    if ($TableData.Count -gt 0) {
        $PropertyNames = $TableData[0] | Get-Member -MemberType NoteProperty | ForEach-Object { $_.Name } | Sort-Object
        Write-LogMessage "CSV will contain columns: $($PropertyNames -join ', ')" "INFO" "Cyan"
    }
    
    return $TableData
}

function Get-UpdatePlanSummary {
    param([array]$Plans)
    
    $StatusCounts = $Plans | Group-Object { $_.status.state } | ForEach-Object {
        [PSCustomObject]@{
            Status = $_.Name
            Count = $_.Count
        }
    }
    
    $ErrorReasons = $Plans | Where-Object { $_.status.errorReasons.Count -gt 0 } | 
        ForEach-Object { $_.status.errorReasons } | 
        Group-Object | ForEach-Object {
            [PSCustomObject]@{
                'Error Reason' = $_.Name
                Count = $_.Count
            }
        }
    
    $CompletedCount = ($Plans | Where-Object { $_.status.state -eq 'PlanCompleted' }).Count
    $SuccessRate = if ($Plans.Count -gt 0) { [math]::Round(($CompletedCount / $Plans.Count) * 100, 2) } else { 0 }
    
    return @{
        StatusDistribution = $StatusCounts
        ErrorReasons = $ErrorReasons
        SuccessRate = $SuccessRate
        TotalPlans = $Plans.Count
    }
}

function Invoke-JamfUpdatePlanAnalysis {
    param(
        [switch]$ExportToCSV,
        [switch]$WriteToEA,
        [switch]$CreateEA,
        [string]$OutputPath = "."
    )
    
    try {
        # Get access token
        if (-not (Get-JamfAccessToken -BaseURL $script:config.jamf_url -ClientID $script:config.client_id -ClientSecret $script:config.client_secret)) {
            throw "Failed to authenticate with Jamf API"
        }
        
        # Handle Extension Attribute setup if requested
        $ExtensionAttributes = $null
        if ($WriteToEA) {
            Write-LogMessage "Setting up Extension Attributes..." "INFO" "Magenta"
            $ExtensionAttributes = Get-AllRequiredExtensionAttributes -CreateIfMissing:$CreateEA
            
            Write-LogMessage "Extension Attributes ready:" "INFO" "Green"
            foreach ($EAName in $ExtensionAttributes.Keys) {
                $EA = $ExtensionAttributes[$EAName]
                Write-LogMessage "  $EAName (ID: $($EA.id))" "INFO" "Cyan"
            }
        }
        
        # Fetch all update plans
        $Plans = @()
        try {
            $Plans = Get-JamfManagedSoftwareUpdatePlans
        }
        catch {
            if ($_.Exception.Message -like "*Managed Software Updates feature is not available*") {
                Write-LogMessage "" "INFO"
                Write-LogMessage "Alternative Actions:" "INFO" "Magenta"
                Write-LogMessage "===================" "INFO" "Magenta"
                Write-LogMessage "Since Managed Software Updates is not available, you could:" "INFO" "White"
                Write-LogMessage "" "INFO"
                Write-LogMessage "  - Run a computer inventory report to see current OS versions:" "INFO" "Cyan"
                Write-LogMessage "    Settings -> Computer Management -> Reports" "INFO" "Gray"
                Write-LogMessage "" "INFO"
                Write-LogMessage "  - Create Smart Groups based on OS version to track updates:" "INFO" "Cyan"
                Write-LogMessage "    Computers -> Smart Computer Groups" "INFO" "Gray"
                Write-LogMessage "" "INFO"
                Write-LogMessage "  - Use Self Service to deploy available updates:" "INFO" "Cyan"
                Write-LogMessage "    Computers -> Policies -> New Policy" "INFO" "Gray"
                Write-LogMessage "" "INFO"
                Write-LogMessage "  - Export computer inventory with OS versions and user details:" "INFO" "Cyan"
                
                # Provide alternative device inventory information
                if ($ExportToCSV) {
                    Write-LogMessage "    Generating device inventory CSV with user details..." "INFO" "Yellow"
                    $AllDevices = Get-AllDevices
                    
                    $DeviceTable = @()
                    foreach ($DeviceId in $AllDevices.DeviceLookup.Keys) {
                        $Device = $AllDevices.DeviceLookup[$DeviceId]
                        $DeviceTable += [PSCustomObject]@{
                            DeviceType = $Device.Type
                            DeviceName = $Device.Name
                            SerialNumber = $Device.SerialNumber
                            Model = $Device.Model
                            OSVersion = $Device.OSVersion
                            LastContact = $Device.LastContact
                            Username = $Device.Username
                            RealName = $Device.RealName
                            Email = $Device.Email
                            Position = $Device.Position
                            ManagementID = $DeviceId
                        }
                    }
                    
                    $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
                    $DevicePath = Join-Path $OutputPath "JamfDevices-Inventory-$Timestamp.csv"
                    $DeviceTable | Export-Csv -Path $DevicePath -NoTypeInformation
                    
                    Write-LogMessage "" "INFO"
                    Write-LogMessage "Device inventory with user details exported to: $DevicePath" "INFO" "Green"
                    $script:SessionStats.CSVFilesCreated = 1
                }
                
                # Still update EAs to reflect "No Plan" state if requested
                if ($WriteToEA) {
                    Write-LogMessage "" "INFO"
                    Write-LogMessage "Updating Extension Attributes to reflect 'No Plan' state..." "INFO" "Yellow"
                    Update-ComputerExtensionAttributes -UpdatePlans @() -ExtensionAttributes $ExtensionAttributes
                }
                
                return $null
            } else {
                throw
            }
        }
        
        if ($Plans.Count -eq 0) {
            Write-LogMessage "No managed software update plans found" "INFO" "Yellow"
            
            # Still update EAs to reflect "No Plan" state if requested
            if ($WriteToEA) {
                Write-LogMessage "Updating Extension Attributes to show 'No Plan' for all computers..." "INFO" "Yellow"
                Update-ComputerExtensionAttributes -UpdatePlans @() -ExtensionAttributes $ExtensionAttributes
            }
            
            # Still provide device inventory if requested
            if ($ExportToCSV) {
                Write-LogMessage "Generating device inventory CSV with user details..." "INFO" "Yellow"
                $AllDevices = Get-AllDevices
                
                $DeviceTable = @()
                foreach ($DeviceId in $AllDevices.DeviceLookup.Keys) {
                    $Device = $AllDevices.DeviceLookup[$DeviceId]
                    $DeviceTable += [PSCustomObject]@{
                        DeviceType = $Device.Type
                        DeviceName = $Device.Name
                        SerialNumber = $Device.SerialNumber
                        Model = $Device.Model
                        OSVersion = $Device.OSVersion
                        LastContact = $Device.LastContact
                        Username = $Device.Username
                        RealName = $Device.RealName
                        Email = $Device.Email
                        Position = $Device.Position
                        ManagementID = $DeviceId
                    }
                }
                
                $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
                $DevicePath = Join-Path $OutputPath "JamfDevices-Inventory-$Timestamp.csv"
                $DeviceTable | Export-Csv -Path $DevicePath -NoTypeInformation
                
                Write-LogMessage "Device inventory with user details exported to: $DevicePath" "INFO" "Green"
                $script:SessionStats.CSVFilesCreated = 1
            }
            
            return $null
        }
        
        # Convert to table format (always includes device details and user information now)
        $DetailedTable = ConvertTo-UpdatePlanTable -Plans $Plans
        
        # Generate summary
        $Summary = Get-UpdatePlanSummary -Plans $Plans
        
        # Display results
        Write-LogMessage "`nSUMMARY STATISTICS" "INFO" "Green"
        Write-LogMessage "Total Plans: $($Summary.TotalPlans)" "INFO" "White"
        Write-LogMessage "Success Rate: $($Summary.SuccessRate)%" "INFO" "White"
        
        if (-not $Unattended) {
            Write-LogMessage "`nSTATUS DISTRIBUTION" "INFO" "Yellow"
            $Summary.StatusDistribution | Format-Table -AutoSize
            
            if ($Summary.ErrorReasons.Count -gt 0) {
                Write-LogMessage "ERROR REASONS" "INFO" "Red"
                $Summary.ErrorReasons | Format-Table -AutoSize
            }
            
            Write-LogMessage "DETAILED PLAN INFORMATION WITH USER DETAILS" "INFO" "Cyan"
            $DetailedTable | Format-Table -AutoSize
        }
        
        # Update Extension Attributes if requested
        if ($WriteToEA) {
            Update-ComputerExtensionAttributes -UpdatePlans $Plans -ExtensionAttributes $ExtensionAttributes
        }
        
        # Export to CSV if requested
        if ($ExportToCSV) {
            $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $DetailedPath = Join-Path $OutputPath "JamfUpdatePlans-Detailed-$Timestamp.csv"
            $SummaryPath = Join-Path $OutputPath "JamfUpdatePlans-Summary-$Timestamp.csv"
            
            $DetailedTable | Export-Csv -Path $DetailedPath -NoTypeInformation
            $Summary.StatusDistribution | Export-Csv -Path $SummaryPath -NoTypeInformation
            
            Write-LogMessage "`nExported to:" "INFO" "Green"
            Write-LogMessage "   Detailed (with user info): $DetailedPath" "INFO" "Gray"
            Write-LogMessage "   Summary: $SummaryPath" "INFO" "Gray"
            
            $script:SessionStats.CSVFilesCreated = 2
        }
        
        return @{
            DetailedTable = $DetailedTable
            Summary = $Summary
            Plans = $Plans
        }
    }
    catch {
        Write-LogMessage "Error during analysis: $($_.Exception.Message)" "ERROR" "Red"
        throw
    }
}

# MAIN EXECUTION
Write-LogMessage "Jamf Managed Software Update Plan Analysis with Enhanced User Details" "INFO" "Magenta"

# Initialize logging first
if (-not (Initialize-LogFile -OutputPath $OutputPath)) {
    Write-LogMessage "Continuing without log file..." "WARNING" "Yellow"
}

Write-LogMessage "Script started with parameters: ConfigFile=$ConfigFile, ExportToCSV=$ExportToCSV, WriteToEA=$WriteToEA, CreateEA=$CreateEA, OutputPath=$OutputPath, Unattended=$Unattended" "INFO" "Gray"

# Check configuration
if (-not (Test-ConfigFile -Path $ConfigFile)) {
    exit 1
}

try {
    # Test if file exists and has content
    if (-not (Test-Path $ConfigFile)) {
        throw "Configuration file not found: $ConfigFile"
    }
    
    $ConfigContent = Get-Content -Path $ConfigFile -Raw
    if ([string]::IsNullOrWhiteSpace($ConfigContent)) {
        throw "Configuration file is empty"
    }
    
    Write-LogMessage "Raw config content preview: $($ConfigContent.Substring(0, [Math]::Min(100, $ConfigContent.Length)))" "INFO" "Gray"
    
    $script:config = $ConfigContent | ConvertFrom-Json
    Write-LogMessage "Configuration loaded successfully" "INFO" "Green"
} catch {
    Write-LogMessage "Failed to load configuration: $($_.Exception.Message)" "ERROR" "Red"
    Write-LogMessage "Config file path: $ConfigFile" "ERROR" "Red"
    Write-LogMessage "" "INFO"
    Write-LogMessage "Please check your jamf_config.json file. It should contain:" "INFO" "Yellow"
    Write-LogMessage '{' "INFO" "White"
    Write-LogMessage '  "jamf_url": "https://your-instance.jamfcloud.com",' "INFO" "White"
    Write-LogMessage '  "client_id": "your_client_id",' "INFO" "White"
    Write-LogMessage '  "client_secret": "your_client_secret"' "INFO" "White"
    Write-LogMessage '}' "INFO" "White"
    exit 1
}

# Validate required fields
$requiredFields = @('jamf_url', 'client_id', 'client_secret')
foreach ($field in $requiredFields) {
    if (-not $script:config.$field -or $script:config.$field -eq "YOUR_CLIENT_ID_HERE" -or $script:config.$field -eq "YOUR_CLIENT_SECRET_HERE") {
        Write-LogMessage "Missing or placeholder value for: $field" "ERROR" "Red"
        exit 1
    }
}

Write-LogMessage "Jamf URL: $($script:config.jamf_url)" "INFO" "Cyan"

# Run analysis
try {
    $Results = Invoke-JamfUpdatePlanAnalysis -ExportToCSV:$ExportToCSV -WriteToEA:$WriteToEA -CreateEA:$CreateEA -OutputPath:$OutputPath
    
    Write-LogMessage "`nAnalysis complete!" "INFO" "Green"
    
    Write-LogMessage "Computer and mobile device details (including user information) included in analysis." "INFO" "Gray"
    if ($ExportToCSV) {
        Write-LogMessage "Results exported to CSV files with enhanced user details." "INFO" "Gray"
    }
    if ($WriteToEA) {
        Write-LogMessage "Extension Attributes updated with current plan state (computers only)." "INFO" "Green"
        Write-LogMessage "Note: Mobile devices don't support Extension Attributes due to API limitations." "INFO" "Yellow"
    }
    
    # Write session summary
    Write-SessionSummary
    
} catch {
    Write-LogMessage "Analysis failed: $($_.Exception.Message)" "ERROR" "Red"
    Write-SessionSummary
    exit 1
}