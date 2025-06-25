# Jamf Managed Software Update (MSU) Analysis

A  PowerShell script for analysing Jamf Pro Managed Software Update plans and automating Extension Attribute management for enhanced MSU reporting.

## Overview

This PowerShell script combines computer and mobile device inventory data with Managed Software Update plan details to provide:

- **Comprehensive Device Analysis**: Includes computers and mobile devices with user details
- **Extension Attribute Management**: Creates and updates 5 Extension Attributes for Smart Groups
- **CSV Export**: Detailed reports with user information (Username, Real Name, Email, Position)
- **Automated Logging**: Timestamped log files in dedicated logs folder
- **Scheduled Task Support**: Unattended mode for automation
- **Fallback Support**: Handles scenarios where Managed Software Updates is disabled

## Prerequisites

- PowerShell 5.1 or later
- Network access to your Jamf Pro instance
- Jamf Pro API credentials with appropriate permissions

## Initial Setup

### Step 1: Create API Credentials in Jamf Pro

1. **Navigate to API Settings**:
   - Log into Jamf Pro
   - Go to: Settings -> User Accounts & Groups -> API Roles and Clients

2. **Create API Role**:
   - Click "New" under API Roles
   - Name: `Jamf MSU Analysis Script`
   - Assign these permissions:
     - Read Managed Software Updates
     - Read Computers
     - Read Mobile Devices
     - Read Computer Extension Attributes
     - Create Computer Extension Attributes
     - Update Computer Extension Attributes
     - Update Computers

3. **Create API Client**:
   - Click "New" under API Clients
   - Display Name: `MSU Analysis Script`
   - Assign the role you just created
   - **Important**: Copy the Client ID and Client Secret immediately

### Step 2: Configure the Script

1. **Download and Run Script Once**:
   ```powershell
   .\Invoke-JamfMSUAnalysis.ps1
   ```

2. **Edit Configuration File**:
   The script creates `jamf_config.json` automatically. Update with your details:
   ```json
   {
     "jamf_url": "https://your-instance.jamfcloud.com",
     "client_id": "your_oauth_client_id",
     "client_secret": "your_oauth_client_secret"
   }
   ```

### Step 3: Create Extension Attributes (First Run Only)

```powershell
.\Invoke-JamfMSUAnalysis.ps1 -WriteToEA -CreateEA
```

This creates 5 Extension Attributes:

- **MSU_Plan_Status** - Update plan status (PlanCompleted/PlanFailed/PlanInProgress/No Plan)
- **MSU_Plan_Action** - Update action type (DOWNLOAD_INSTALL_SCHEDULE/etc)
- **MSU_Plan_Version_Type** - Version target (LATEST_MAJOR/LATEST_MINOR/etc)
- **MSU_Plan_Error_Reasons** - Error details or "No Errors"
- **MSU_Plan_Force_Install_Date** - Scheduled installation date

## Usage

### Basic Commands

**Analysis with CSV Export**:
```powershell
.\Invoke-JamfMSUAnalysis.ps1 -ExportToCSV
```

**Update Extension Attributes**:
```powershell
.\Invoke-JamfMSUAnalysis.ps1 -WriteToEA
```

**Full Analysis (CSV + Extension Attributes)**:
```powershell
.\Invoke-JamfMSUAnalysis.ps1 -ExportToCSV -WriteToEA
```

**Silent Mode (for Scheduled Tasks)**:
```powershell
.\Invoke-JamfMSUAnalysis.ps1 -WriteToEA -Unattended
```

### Command Parameters

| Parameter | Description |
|-----------|-------------|
| `-ConfigFile` | Path to JSON configuration file (default: `.\jamf_config.json`) |
| `-ExportToCSV` | Export detailed results to CSV files |
| `-WriteToEA` | Update Extension Attributes on computers |
| `-CreateEA` | Create Extension Attributes if they don't exist |
| `-OutputPath` | Directory for exports (default: current directory) |
| `-Unattended` | Run silently without console output |
| `-Help` | Display help information |

## Output Files

### Log Files
- **Location**: `logs/JamfMSU-Analysis-YYYYMMDD-HHMMSS.log`
- **Contains**: Timestamped activity log with session summary
- **Includes**: Start/end times, device counts, Extension Attribute updates, error details

### CSV Files (when using `-ExportToCSV`)

**Detailed Report**: `JamfUpdatePlans-Detailed-YYYYMMDD-HHMMSS.csv`
- Complete update plan details with device and user information
- Columns include: Device Type, Computer Name, Serial Number, Model, OS Version, Username, Real Name, Email, Position, Update Status, etc.

**Summary Report**: `JamfUpdatePlans-Summary-YYYYMMDD-HHMMSS.csv`
- Status distribution summary

**Device Inventory**: `JamfDevices-Inventory-YYYYMMDD-HHMMSS.csv` (when no update plans exist)
- Complete device inventory with user details

## Smart Group Integration

### Status-Based Smart Groups

Create Smart Groups using the Extension Attributes:

**Failed Updates**:
- Extension Attribute `MSU_Plan_Status` is `PlanFailed`

**Completed Updates**:
- Extension Attribute `MSU_Plan_Status` is `PlanCompleted`

**Updates in Progress**:
- Extension Attribute `MSU_Plan_Status` is `PlanInProgress`

**No Update Plans**:
- Extension Attribute `MSU_Plan_Status` is `No Plan`

### Action-Based Smart Groups

**Scheduled Updates**:
- Extension Attribute `MSU_Plan_Action` is `DOWNLOAD_INSTALL_SCHEDULE`

**Major Version Updates**:
- Extension Attribute `MSU_Plan_Version_Type` is `LATEST_MAJOR`

**Error Troubleshooting**:
- Extension Attribute `MSU_Plan_Error_Reasons` contains `SPECIFIC_VERSION_UNAVAILABLE`

## Scheduled Task Setup

### Windows Task Scheduler

**Daily Extension Attribute Updates**:
- **Schedule**: Daily at 6:00 AM
- **Program**: `powershell.exe`
- **Arguments**: `-ExecutionPolicy Bypass -File "C:\Scripts\Invoke-JamfMSUAnalysis.ps1" -WriteToEA -Unattended`
- **Start in**: `C:\Scripts`

**Weekly Full Reports**:
- **Schedule**: Weekly on Sunday at 2:00 AM
- **Program**: `powershell.exe`
- **Arguments**: `-ExecutionPolicy Bypass -File "C:\Scripts\Invoke-JamfMSUAnalysis.ps1" -ExportToCSV -WriteToEA -Unattended -OutputPath "C:\Reports\Jamf"`
- **Start in**: `C:\Scripts`

### Task Scheduler Settings
- Run whether user is logged on or not
- Run with highest privileges (if needed for network access)
- Hidden
- Stop task if runs longer than 2 hours

## Troubleshooting

### Common Issues

**Authentication Errors**:
- Verify Client ID and Client Secret in `jamf_config.json`
- Check API client is enabled in Jamf Pro
- Ensure API role has correct permissions

**Extension Attribute Errors**:
- Run with `-CreateEA` parameter first time
- Verify "Create Computer Extension Attributes" permission
- Check "Update Computer Extension Attributes" permission

**Network Issues**:
- Verify Jamf Pro URL is accessible
- Check firewall settings
- Test network connectivity to Jamf instance

**Managed Software Updates Not Available**:
- Feature may not be enabled in Jamf Pro
- Check: Computers -> Software Updates -> Enable/Disable Software Updates
- Verify licence includes this feature

### Log Analysis

Check log files in the `logs/` directory for detailed error information. Each log includes:
- Session start/end times and duration
- Device processing statistics
- Extension Attribute creation/update results
- Detailed error messages with troubleshooting context

## Important Notes

- **Extension Attributes**: Only supported for computers, not mobile devices (API limitation)
- **Performance**: Script processes mobile devices for CSV exports but skips them during Extension Attribute updates
- **Permissions**: Ensure API client has all required permissions before first run
- **Backup**: Consider backing up existing Extension Attributes before first run with `-CreateEA`

## Support

For questions or issues with this script, please check:

1. The detailed log files in the `logs/` directory
2. The help output: `.\Invoke-JamfMSUAnalysis.ps1 -Help`

## Author

**Carl Flanagan**  

## Licence

This project is licensed under the MIT Licence - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

