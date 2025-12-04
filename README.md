# MtaStsResults PowerShell Module

A comprehensive PowerShell module for downloading, parsing, and analyzing MTA-STS (Mail Transfer Agent Strict Transport Security) and DMARC (Domain-based Message Authentication, Reporting and Conformance) JSON reports from Microsoft Exchange mailboxes using the Microsoft Graph API.

## Overview

This module streamlines the process of collecting and analyzing DMARC/MTA-STS aggregate reports directly from Exchange mailboxes. It provides three primary functions that work together to:

1. **Download** DMARC/MTA-STS JSON report attachments from one or more mailboxes
2. **Parse** the JSON reports and extract policy compliance statistics
3. **Clean up** downloaded files after processing

The module uses Microsoft Graph API for all Exchange operations, eliminating the need for Exchange server access and supporting modern cloud-based deployments.

## Features

- **Efficient Attachment Retrieval**: Leverages Microsoft Graph API's inline attachment expansion to minimize API calls
- **Intelligent Fallback Logic**: Gracefully handles API response variations when inline content is unavailable
- **Structured Output**: Parses JSON reports into PowerShell objects for easy analysis and manipulation
- **Batch Processing**: Supports processing multiple mailboxes in a single operation
- **Flexible Configuration**: Customizable lookback periods, output paths, and selective operation skipping
- **Comprehensive Logging**: Detailed verbose output and progress indicators throughout execution
- **Error Resilience**: Continues processing despite individual file failures with detailed warning messages

## Files

- **MtaStsResults.psd1** — Module manifest defining module metadata, dependencies, and exported functions
- **MtaStsResults.psm1** — Module implementation containing the three exported functions:
  - `Invoke-DmarcAttachmentDownloader` — Downloads report attachments from mailboxes
  - `Invoke-JsonParse` — Parses and extracts statistics from JSON reports
  - `Invoke-CleanUp` — Removes downloaded files after processing
- **Get-MtaStsResults.ps1** — Legacy script format (deprecated, use module instead)

## Prerequisites

### PowerShell Version
- **PowerShell 7.0 or later** (Windows, macOS, or Linux)
- Windows PowerShell 5.1 may work with the Microsoft.Graph module but is not officially supported

### Dependencies
- **Microsoft.Graph** PowerShell module (version 1.0 or later)
  ```powershell
  Install-Module -Name Microsoft.Graph -Force
  ```

### Azure AD / Graph API Configuration
An Azure AD application registration is required with the following configuration:

1. **Certificate-based Authentication**: Create a self-signed or CA-issued certificate for authentication
   - Export the public certificate (.cer) and upload to Azure AD app registration
   - Store the private certificate in Windows Certificate Store (LocalMachine\My)

2. **API Permissions**: Grant the following delegated permissions:
   - `Mail.Read` — Read mail messages and attachments
   - `User.Read` — Read user profiles (for mailbox resolution)

3. **Required Information**:
   - Tenant ID (Azure AD tenant GUID)
   - Client ID (Application ID from app registration)
   - Certificate Thumbprint (hex string of the certificate's thumbprint)

## Installation

1. Clone or download the module to your PowerShell modules directory:
   ```powershell
   # User module path (current user only)
   $ModulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\MtaStsResults"
   
   # System module path (all users, requires admin)
   $ModulePath = "C:\Program Files\PowerShell\Modules\MtaStsResults"
   
   New-Item -ItemType Directory -Path $ModulePath -Force
   Copy-Item -Path ".\MtaStsResults.psd1" -Destination $ModulePath
   Copy-Item -Path ".\MtaStsResults.psm1" -Destination $ModulePath
   ```

2. Import the module:
   ```powershell
   Import-Module MtaStsResults
   ```

3. Verify the module is loaded:
   ```powershell
   Get-Module MtaStsResults
   ```

## Usage

### Basic Workflow

The typical workflow involves three steps that must be executed in order:

#### 1. Connect to Microsoft Graph
```powershell
# Connect with certificate-based authentication
Connect-MgGraph -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
                -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
                -Certificate (Get-Item Cert:\LocalMachine\My\THUMBPRINT)
```

#### 2. Download Attachments
```powershell
# Download from a single mailbox with default settings (1 day lookback)
Invoke-DmarcAttachmentDownloader -Mailbox "dmarc-reports@contoso.com"

# Download from multiple mailboxes with custom lookback period
Invoke-DmarcAttachmentDownloader -Mailbox @("dmarc@contoso.com", "dmarc@fabrikam.com") `
                                 -OutRoot "C:\Reports" `
                                 -DaysLookBack 7
```

#### 3. Parse Reports
```powershell
# Parse downloaded files from the default location
$results = Invoke-JsonParse

# Parse files from a specific directory
$results = Invoke-JsonParse -PathToScan "C:\Reports\dmarc@contoso.com\04122025\"

# Access parsed data programmatically
$results | Where-Object { $_.TotalFailedSessions -gt 0 } | Format-Table
```

#### 4. Clean Up Files (Optional)
```powershell
# Remove downloaded files after processing
Invoke-CleanUp -PathToScan "C:\Reports\dmarc@contoso.com\04122025\"
```

#### 5. Disconnect
```powershell
Disconnect-MgGraph
```

### Complete Example Script

```powershell
# Import module
Import-Module MtaStsResults

# Set variables
$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$ClientId = "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
$CertThumbprint = "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
$Mailboxes = @("dmarc@contoso.com", "dmarc@fabrikam.com")
$OutputPath = "C:\DMARC-Reports"

# Connect to Microsoft Graph
$Cert = Get-Item -Path "Cert:\LocalMachine\My\$CertThumbprint"
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $Cert -NoWelcome

try {
    # Download attachments from the last 30 days
    Write-Host "Downloading DMARC reports..." -ForegroundColor Cyan
    Invoke-DmarcAttachmentDownloader -Mailbox $Mailboxes `
                                     -OutRoot $OutputPath `
                                     -DaysLookBack 30

    # Parse JSON reports
    Write-Host "Parsing reports..." -ForegroundColor Cyan
    $parsedReports = Invoke-JsonParse -PathToScan $OutputPath

    # Display high-level summary
    Write-Host "Summary:" -ForegroundColor Green
    $parsedReports | Measure-Object -Property TotalFailedSessions -Sum

    # Export to CSV for further analysis
    $parsedReports | Export-Csv -Path "$OutputPath\summary.csv" -NoTypeInformation

    # Clean up downloaded files
    Invoke-CleanUp -PathToScan $OutputPath
} finally {
    Disconnect-MgGraph
}
```

## Function Reference

### Invoke-DmarcAttachmentDownloader

Downloads MTA-STS/DMARC JSON report attachments from Exchange mailboxes.

**Parameters:**
- `-Mailbox` (String[], Required) — Email address(es) of mailbox(es) to download from
- `-OutRoot` (String, Optional) — Root output directory; defaults to `.\dmarc-attachments`
- `-DaysLookBack` (Int, Optional) — Lookback period in days; defaults to `1`

**Output:**
Files are organized in the directory structure:
```
$OutRoot/
  └── $Mailbox/
      └── DDMMYYYY/
          ├── report1.json.gz
          ├── report2.json.gz
          └── ...
```

### Invoke-JsonParse

Parses GZIP-compressed JSON reports and extracts policy compliance statistics.

**Parameters:**
- `-PathToScan` (String, Optional) — Directory containing .json.gz files to parse

**Output:**
Returns a PSCustomObject array with properties:
- `File`, `OrgName`, `StartDatetime`, `EndDatetime`, `ContactInfo`, `ReportId`
- `PolicyType`, `PolicyDomain`, `PolicyString`, `MXHosts`
- `TotalSuccessfulSessions`, `TotalFailedSessions`

### Invoke-CleanUp

Removes downloaded report files from a directory.

**Parameters:**
- `-PathToScan` (String, Optional) — Directory to clean up

## Troubleshooting

### Certificate Not Found
```
Error: Certificate not found. Provide a valid certificate path.
```
**Solution**: Verify the certificate thumbprint and that it's installed in the LocalMachine\My store.
```powershell
Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*your-cert*" }
```

### No Attachments Found
Verify that:
- The mailbox has received DMARC/MTA-STS reports
- The lookback period is sufficient (`-DaysLookBack` parameter)
- The mailbox email address is correct and accessible

### JSON Parsing Fails
- Ensure files are valid GZIP-compressed JSON following DMARC aggregate report format
- Check file permissions
- Verify disk space for temporary decompression

### API Permission Errors
Ensure the Azure AD app registration has:
- `Mail.Read` delegated permission
- `User.Read` delegated permission
- Certificate properly configured for the app

## Best Practices

1. **Run with Verbose Output**: Use `-Verbose` flag to see detailed processing logs
   ```powershell
   Invoke-DmarcAttachmentDownloader -Mailbox $Mailboxes -Verbose
   ```

2. **Schedule Regular Downloads**: Create a Windows Task Scheduler job to run daily/weekly
3. **Archive Reports**: Keep parsed CSV exports for historical analysis
4. **Monitor Failures**: Check `TotalFailedSessions` to identify TLS/policy compliance issues
5. **Cleanup Strategy**: Remove old files to manage disk space, or archive to external storage

## Limitations

- Requires certificate-based authentication to Microsoft Graph API
- Processes one mailbox at a time (though multiple mailboxes can be specified)
- JSON files must follow the standard DMARC aggregate report format
- Filters for only .gz/.tgz attachments; other compressed formats are skipped

## License

See LICENSE file for details.

## Support

For issues, questions, or contributions, please refer to the GitHub repository.
