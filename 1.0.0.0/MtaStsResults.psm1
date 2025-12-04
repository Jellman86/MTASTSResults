#Requires -Version 7.0

<#
.SYNOPSIS
    MTA-STS/DMARC Report Downloader and Parser Module
    
.DESCRIPTION
    This module provides functionality to download MTA-STS (Mail Transfer Agent Strict Transport Security)
    and DMARC (Domain-based Message Authentication, Reporting and Conformance) JSON reports from Microsoft
    Exchange mailboxes using the Microsoft Graph API, parse the JSON content, extract policy compliance
    statistics, and optionally clean up downloaded files.
    
.NOTES
    - Requires PowerShell 7.0 or later
    - Requires Microsoft.Graph module
    - Requires proper Azure AD app registration with certificate-based authentication
#>

# Module-scoped variables to store parsed data across function calls
$script:parsedJsonContent = @()
$script:msgFolder = $null

<#
.SYNOPSIS
    Downloads MTA-STS/DMARC JSON attachments from specified mailboxes.
    
.DESCRIPTION
    Connects to Microsoft Exchange mailboxes and downloads DMARC/MTA-STS JSON report attachments
    (.gz and .tgz files) that were received within a specified lookback period. The function filters
    for messages with attachments received since a specified date and saves them to a local directory
    structure organized by mailbox and date.
    
    The function intelligently handles attachment retrieval, preferring inline attachment content
    when available (returned via $expand=attachments) and falling back to explicit attachment
    retrieval API calls when necessary. Content is decoded from Base64 and written to disk.
    
.PARAMETER Mailbox
    Specifies one or more mailbox email addresses to scan. Example: "dmarc-reports@domain.com"
    
.PARAMETER OutRoot
    Root output directory where downloaded files will be saved. Subdirectories are created per
    mailbox and date. Default is ".\dmarc-attachments" in the current working directory.
    
.PARAMETER DaysLookBack
    Number of days to look back when filtering messages. Default is 1. For example, DaysLookBack
    of 7 will retrieve all qualifying messages from the past 7 days.
    
.EXAMPLE
    Invoke-DmarcAttachmentDownloader -Mailbox "dmarc@contoso.com" -OutRoot "C:\Reports" -DaysLookBack 7
    
    Downloads all MTA-STS/DMARC attachments from dmarc@contoso.com received in the last 7 days
    and saves them under C:\Reports\dmarc@contoso.com\DDMMYYYY\ directories.
    
.NOTES
    - Requires active Microsoft Graph connection (Connect-MgGraph must be called first)
    - Files are organized under $OutRoot\$Mailbox\$DateFolder\
    - Skips non-.gz/.tgz attachments with a warning
    - Sanitizes filenames to be compatible with Windows filesystem
#>
function Invoke-DmarcAttachmentDownloader {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Mailbox,
        
        [Parameter(Mandatory = $false)]
        [string]$OutRoot = ".\dmarc-attachments",
        
        [Parameter(Mandatory = $false)]
        [int]$DaysLookBack = 1
    )

    foreach ($mbx in $Mailbox) {
        Write-Verbose "Processing mailbox: $mbx"
        
        $outDir = Join-Path $OutRoot $mbx
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
        Write-Verbose "Created output directory: $outDir"

        # Calculate the UTC timestamp for the lookback period
        $sinceUtc = (Get-Date).AddDays(-[int]$DaysLookBack).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        Write-Verbose "Filtering messages from: $sinceUtc"
        
        # Build filter for messages with attachments received in the lookback period
        $filter = "hasAttachments eq true and receivedDateTime ge $sinceUtc"

        # Retrieve messages with attachments expanded inline to reduce API calls
        Write-Verbose "Retrieving messages with filter: $filter"
        $messages = Get-MgUserMessage -UserId $mbx -Filter $filter -All -ExpandProperty Attachments
        
        if (-not $messages) {
            Write-Host "[$mbx] No messages found in lookback period" -ForegroundColor Yellow
            continue
        }
        
        Write-Host "[$mbx] Found $($messages.Count) message(s) with attachments" -ForegroundColor Cyan

        foreach ($m in $messages) {
            # Prefer attachments returned inline with the message to reduce API calls;
            # fallback to explicit attachment retrieval if not present
            $atts = if ($null -ne $m.Attachments -and $m.Attachments.Count -gt 0) {
                $m.Attachments
            } else {
                Get-MgUserMessageAttachment -UserId $mbx -MessageId $m.Id -All
            }

            foreach ($att in $atts) {
                # Skip non-GZIP attachments
                if ($att.Name -notmatch '(?i)\.(gz|tgz)$') {
                    Write-Host "[$mbx] Skipping attachment '$($att.Name)' (not .gz/.tgz)" -ForegroundColor DarkGray
                    continue
                }

                # Create date-based subdirectory
                $script:msgFolder = Join-Path $outDir (Get-Date -Format ddMMyyyy)
                New-Item -Path $script:msgFolder -ItemType Directory -Force | Out-Null

                # Sanitize filename to be Windows-compatible
                $fileName = ($att.Name -replace '[\\/:*?""<>|]', '_')
                $path = Join-Path $script:msgFolder $fileName

                # Attempt to retrieve attachment content from AdditionalProperties (inline expansion)
                if ($null -ne $att.AdditionalProperties -and $att.AdditionalProperties.ContainsKey('contentBytes') -and $att.AdditionalProperties.contentBytes) {
                    try {
                        $contentBytes = [System.Convert]::FromBase64String($att.AdditionalProperties.contentBytes)
                        [System.IO.File]::WriteAllBytes($path, $contentBytes)
                        Write-Host "[$mbx] Saved $path (via inline attachment)" -ForegroundColor Green
                        continue
                    } catch {
                        Write-Warning "Failed to decode inline attachment $($att.Name): $_"
                    }
                }

                # Fallback: Ensure we have the actual attachment resource (may be minimal when expanded)
                $fullAtt = $att
                if (($null -eq $fullAtt.ContentBytes -or $fullAtt.ContentBytes -eq '') -and $null -ne $fullAtt.Id) {
                    Write-Verbose "Retrieving full attachment object for: $($att.Name)"
                    $fullAtt = Get-MgUserMessageAttachment -UserId $mbx -MessageId $m.Id -AttachmentId $fullAtt.Id
                }

                # Decode ContentBytes (may be string or byte array depending on API response)
                try {
                    if ($fullAtt.ContentBytes -is [string]) {
                        $bytes = [Convert]::FromBase64String($fullAtt.ContentBytes)
                    } elseif ($fullAtt.ContentBytes -is [byte[]]) {
                        $bytes = $fullAtt.ContentBytes
                    } else {
                        $bytes = [Convert]::FromBase64String([string]$fullAtt.ContentBytes)
                    }

                    [IO.File]::WriteAllBytes($path, $bytes)
                    Write-Host "[$mbx] Saved $path" -ForegroundColor Green
                } catch {
                    Write-Error "Failed to process attachment $($att.Name) for message $($m.Id): $_"
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Parses DMARC/MTA-STS JSON reports and extracts policy compliance statistics.
    
.DESCRIPTION
    Scans a directory for GZIP-compressed JSON files (.json.gz or .tgz) and extracts
    policy compliance data. Decompresses GZIP content, parses JSON structure, and builds
    a collection of policy statistics including organization name, policy domain, and
    success/failure session counts.
    
    The function expects JSON files to follow the standard DMARC aggregate report format
    which includes:
    - Organization metadata (organization-name, contact-info, report-id)
    - Date range (start-datetime, end-datetime)
    - Policy information (policy-type, policy-domain, policy-string, mx-host)
    - Summary statistics (total-successful-session-count, total-failure-session-count)
    
.PARAMETER PathToScan
    Directory path containing GZIP-compressed JSON files to parse. If not specified,
    uses the current message folder from the last download operation.
    
.EXAMPLE
    Invoke-JsonParse -PathToScan "C:\Reports\dmarc@contoso.com\04122025\"
    
    Parses all .json.gz files in the specified directory and displays a formatted table
    of policy statistics.
    
.OUTPUTS
    Returns a PSCustomObject collection with the following properties:
    - File: Original filename
    - OrgName: Reporting organization name
    - StartDatetime: Report period start
    - EndDatetime: Report period end
    - ContactInfo: Organization contact email
    - ReportId: Unique report identifier
    - PolicyType: Type of policy (e.g., 'tlsrpt', 'dmarc')
    - PolicyDomain: The domain that the policy applies to
    - PolicyString: Raw policy string
    - MXHosts: Mail server hosts covered by the policy
    - TotalSuccessfulSessions: Count of successful SMTP connections
    - TotalFailedSessions: Count of failed SMTP connections
    
.NOTES
    - Handles both compressed (.gz) and plain JSON files
    - Returns the parsed content formatted as an auto-sized table
    - Failed files are logged as warnings but processing continues
    - Stores results in script-scoped variable $script:parsedJsonContent for use by other functions
#>
function Invoke-JsonParse {
    param (
        [Parameter(Mandatory = $false)]
        [string]$PathToScan = $script:msgFolder
    )

    Write-Verbose "Parsing JSON files from: $PathToScan"
    $script:parsedJsonContent = @()

    if (-not (Test-Path $PathToScan)) {
        Write-Warning "Path not found: $PathToScan"
        return $null
    }

    $jsonFiles = Get-ChildItem -Path $PathToScan -Recurse -Include *.json.gz -File
    Write-Host "Found $($jsonFiles.Count) JSON.GZ file(s) to parse" -ForegroundColor Cyan
    
    if ($jsonFiles.Count -eq 0) {
        Write-Warning "No .json.gz files found in $PathToScan"
        return $null
    }

    $jsonFiles | ForEach-Object {
        $file = $_
        Write-Verbose "Processing file: $($file.FullName)"
        
        try {
            # Decompress GZIP file and read JSON content
            if ($file.Extension -ieq '.gz') {
                $fs = [System.IO.File]::OpenRead($file.FullName)
                $gz = New-Object System.IO.Compression.GzipStream($fs, [System.IO.Compression.CompressionMode]::Decompress)
                $sr = New-Object System.IO.StreamReader($gz)
                $text = $sr.ReadToEnd()
                $sr.Close()
                $gz.Close()
                $fs.Close()
            } else {
                $text = Get-Content -Path $file.FullName -Raw
            }

            # Parse JSON content
            $json = $text | ConvertFrom-Json

            # Iterate through policies in the report
            if ($null -ne $json.policies -and $json.policies.Count -gt 0) {
                foreach ($entry in $json.policies) {
                    $policy = $entry.policy
                    $summary = $entry.summary

                    # Create custom object with extracted data
                    $script:parsedJsonContent += [PSCustomObject]@{
                        File                     = $file.Name
                        OrgName                  = $json.'organization-name'
                        StartDatetime            = $json.'date-range'.'start-datetime'
                        EndDatetime              = $json.'date-range'.'end-datetime'
                        ContactInfo              = $json.'contact-info'
                        ReportId                 = $json.'report-id'
                        PolicyType               = $policy.'policy-type'
                        PolicyDomain             = $policy.'policy-domain'
                        PolicyString             = $policy.'policy-string'
                        MXHosts                  = $policy.'mx-host'
                        TotalSuccessfulSessions  = $summary.'total-successful-session-count'
                        TotalFailedSessions      = $summary.'total-failure-session-count'
                    }
                }
                Write-Host "[$($file.Name)] Extracted $($json.policies.Count) policies" -ForegroundColor Green
            } else {
                Write-Warning "No policies found in $($file.Name)"
            }
        } catch {
            Write-Warning "Failed to parse $($file.FullName): $_"
        }
    }

    # Display formatted summary table
    Write-Host "`nParsed Policy Summary:" -ForegroundColor Cyan
    return $script:parsedJsonContent | Select-Object PolicyDomain, OrgName, TotalFailedSessions, TotalSuccessfulSessions | Format-Table -AutoSize
}

<#
.SYNOPSIS
    Deletes downloaded report files from the specified directory.
    
.DESCRIPTION
    Removes all files from a specified directory and its subdirectories. This is useful
    for cleaning up downloaded DMARC/MTA-STS reports after processing to recover disk space
    or maintain compliance with data retention policies.
    
.PARAMETER PathToScan
    Directory path containing files to delete. If not specified, uses the current message
    folder from the last download operation. The function removes all files recursively
    from this path.
    
.EXAMPLE
    Invoke-CleanUp -PathToScan "C:\Reports\dmarc@contoso.com\04122025\"
    
    Deletes all files from the specified directory.
    
.NOTES
    - Removes files recursively from the specified directory
    - Logs deleted files in green with success messages
    - Logs failures as warnings but continues processing
    - Does not delete the directory structure itself, only files
#>
function Invoke-CleanUp {
    param (
        [Parameter(Mandatory = $false)]
        [string]$PathToScan = $script:msgFolder
    )

    Write-Verbose "Cleaning up files from: $PathToScan"

    if (-not (Test-Path $PathToScan)) {
        Write-Warning "Path not found: $PathToScan"
        return
    }

    $files = Get-ChildItem -Path $PathToScan -Recurse -File
    Write-Host "Removing $($files.Count) file(s) from $PathToScan" -ForegroundColor Cyan
    
    if ($files.Count -eq 0) {
        Write-Host "No files found to clean up" -ForegroundColor Yellow
        return
    }

    $files | ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Force
            Write-Host "Deleted: $($_.FullName)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to delete $($_.FullName): $_"
        }
    }
}

Export-ModuleMember -Function Invoke-DmarcAttachmentDownloader, Invoke-JsonParse, Invoke-CleanUp
