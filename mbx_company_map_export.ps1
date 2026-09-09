<#
.SYNOPSIS
    Processes mailboxes and stores emails with attachments to SharePoint with folder structure

.DESCRIPTION
    This script scans specified mailboxes for emails within a time range, extracts case information,
    creates SharePoint folder structure, and stores emails and attachments.
    Uses multiple CSV files (by year) for case metadata lookup and URL tracking.
    Logs all operations to a timestamped log file in the script directory.

    Authentication uses certificate-based auth with certificate loaded from Windows Certificate Store
    by thumbprint (searches CurrentUser\My and LocalMachine\My stores).

.PARAMETER FromTimestamp
    Start date/time for email search (format: "yyyy-MM-ddTHH:mm:ssZ")

.PARAMETER ToTimestamp
    End date/time for email search (format: "yyyy-MM-ddTHH:mm:ssZ")

.PARAMETER Mailboxes
    Array of email addresses to process

.PARAMETER SharePointSiteUrl
    SharePoint site URL 

.PARAMETER SharePointBasePath
    Base folder path in SharePoint (e.g., "ARCHIV_Dokumenty")

.PARAMETER CSVFolderPath
    Path to folder containing CSV files named 2021.csv, 2022.csv, 2023.csv, 2024.csv

.PARAMETER CSVCaseIdColumn
    Column name in CSV containing case IDs (default: "Číslo podání")

.PARAMETER ClientId
    Azure AD Application Client ID

.PARAMETER TenantId
    Azure AD Tenant ID

.PARAMETER CertificateThumbprint
    Certificate thumbprint to load from certificate store

.PARAMETER Proxy
    Proxy server URL (optional)

.PARAMETER UpdateCSVWithUrl
    Whether to update CSV files with SharePoint URLs (default: $true)

.EXAMPLE
    
#>


param(
[string]$FromTimestamp ="2024-11-13T14:07:44Z",
[string]$ToTimestamp = "2024-11-13T14:07:48Z",
[string]$workerId,
#$yearsToIgnore = @(2021,2022,2023,2024)
[string[]]$Mailboxes = @('mbx@domena.cz'),
[string]$SharePointSiteUrl = "",
[string]$SharePointBasePath = "ARCHIV_Dokumenty",
[string]$CSVFolderPath = "C:\Users\xyz\cases-csv",
[string]$CSVCaseIdColumn = "Číslo podání",
[string]$ClientId = '',
[string]$TenantId = '',
[string]$CertificateThumbprint = '',
[string]$Proxy = '',
[bool]$UpdateCSVWithUrl = $true
#[int]$processFirst = 24
)

# Import required modules
Import-Module AadAuthenticationFactory -ErrorAction Stop

# Initialize log file path
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFileName = "Migration-$($workerId)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:LogFilePath = Join-Path $scriptPath $logFileName

# Function to write log messages
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Write to console with color based on level
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor White }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
    }

    # Write to log file
    try {
        Add-Content -Path $script:LogFilePath -Value $logMessage -ErrorAction Stop
    } catch {
        Write-Host "Warning: Could not write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Function to get certificate from store by thumbprint
function Get-CertificateByThumbprint {
    param(
        [string]$Thumbprint
    )

    try {
        Write-Log "Loading certificate from store with thumbprint: $Thumbprint"

        # Search in CurrentUser\My store first
        $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint } | Select-Object -First 1

        if ($cert) {
            Write-Log "Certificate found in CurrentUser\My store"
            return $cert
        }

        # If not found, search in LocalMachine\My store
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint } | Select-Object -First 1

        if ($cert) {
            Write-Log "Certificate found in LocalMachine\My store"
            return $cert
        }

        throw "Certificate with thumbprint '$Thumbprint' not found in CurrentUser\My or LocalMachine\My stores"
    }
    catch {
        Write-Log "Error loading certificate: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to get access token using AadAuthenticationFactory with certificate
function Get-GraphAccessToken {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Proxy = $null
    )

    try {
        Write-Log "Authenticating to Microsoft Graph API using certificate..."

        # Configure proxy if provided
        if ($Proxy) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($Proxy)
            [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            Write-Log "Proxy configured: $Proxy"
        }

        # Create authentication factory with certificate
        New-AadAuthenticationFactory -DefaultScopes "https://graph.microsoft.com/.default" `
            -ClientId $ClientId `
            -TenantId $TenantId `
            -X509Certificate $Certificate

        # Get token as hashtable for headers
        $headers = Get-AadToken -Verbose -AsHashTable

        Write-Log "Successfully authenticated to Microsoft Graph"
        return $headers
    }
    catch {
        Write-Log "Failed to authenticate: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to invoke Graph API
function Invoke-GraphApi {
    param(
        [string]$Uri,
        [System.Object]$Headers,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json"
    )

    try {
        $params = @{
            Uri = $Uri
            Headers = $Headers
            Method = $Method
        }

        if ($Body -and $Method -ne "GET") {
            if ($ContentType -eq "application/json") {
                $params.Body = ($Body | ConvertTo-Json -Depth 10)
            } else {
                $params.Body = $Body
            }
        }

        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        Write-Log "Graph API call failed: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to parse email subject for case information
function Parse-EmailSubject {
    param([string]$Subject)

    # Pattern: number-caseId-year (optionally preceded by RE:, FW:, etc.)
    if ($Subject -match '(?:RE:|FW:)?\s*(\d+)-(\d+)-(\d{4})') {
        return @{
            Number = $matches[1]
            CaseId = $matches[2]
            Year = $matches[3]
            IsValid = $true
        }
    }

    return @{
        IsValid = $false
    }
}

# Function to load all CSV files by year into a dictionary
function Load-CSVsByYear {
    param(
        [string]$CSVFolderPath,
        [string[]]$Years = @("2021", "2022", "2023", "2024","2025")
    )

    $csvDictionary = @{}

    try {
        Write-Log "Loading CSV files from folder: $CSVFolderPath"

        foreach ($year in $Years) {
            $csvPath = Join-Path $CSVFolderPath "$year.csv"

            if (Test-Path $csvPath) {
                Write-Log "Loading CSV for year $year from: $csvPath"
               $csvDictionary[$year] = Import-Csv -Path $csvPath -Encoding Default -Delimiter ";"
                #$csvDictionary[$year] =  Import-Csv -Path $csvPath -Encoding ([System.Text.Encoding]::GetEncoding(1250)) -Delimiter ";"
                Write-Log "Loaded $($csvDictionary[$year].Count) records for year $year"
            } else {
                Write-Log "CSV file not found for year $year at: $csvPath" -Level WARNING
                $csvDictionary[$year] = @()
            }
        }

        return $csvDictionary
    }
    catch {
        Write-Log "Error loading CSV files: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to search case ID in CSV dictionary by year and return all row data as metadata
function Find-CaseInCSV {
    param(
        [hashtable]$CSVDictionary,
        [string]$Year,
        [string]$CaseId,
        [string]$ColumnName = 'Číslo podání'
    )

    try {
        Write-Log "Searching for case ID '$CaseId' in year $Year CSV, column '$ColumnName'"

        if (-not $CSVDictionary.ContainsKey($Year)) {
            Write-Log "Year $Year not found in CSV dictionary" -Level WARNING
            return $null
        }

        $csvData = $CSVDictionary["$Year"]
        $case = $csvData | Where-Object { $_.$ColumnName -eq $CaseId }

        if ($case) {
            Write-Log "Found case ID '$CaseId' in year $Year CSV"

            # Convert the row to a hashtable with all properties
            $metadata = @{}
            $case.PSObject.Properties | ForEach-Object {
                $metadata[$_.Name] = $_.Value
            }

            return $metadata
        } else {
            Write-Log "Case ID '$CaseId' not found in year $Year CSV" -Level WARNING
            return $null
        }
    }
    catch {
        Write-Log "Error searching CSV: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-MailArchiveFolder {
    param(
        $siteId,
        $headers,
        $archiveFolderName = "Archiv_Dokumenty"
    )

    try {
        $url = "https://graph.microsoft.com/v1.0/sites/$($siteId)/drive/root/children"
        $folders = Invoke-RestMethod -Method GET -Uri $url -Headers $headers
    }
    catch {
        Write-Host "$($_.exception.message)"
    }

    $folder = $folders.value | Where-Object { $_.name -eq $archiveFolderName }
    return $folder
}

# Function to create SharePoint folder or get existing one
function New-SharePointFolder {
    param(
        [hashtable]$Headers,
        [string]$SiteId,
        [string]$DriveId,
        [string]$ParentPath,
        [string]$FolderName
    )

    try {
        # First, try to get the folder if it exists
        $fullPath = if ($ParentPath) { "$ParentPath/$FolderName" } else { $FolderName }
        $encodedPath = [System.Web.HttpUtility]::UrlEncode($fullPath)
        $checkUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$encodedPath"

        try {
            $existingFolder = Invoke-RestMethod -Method GET -Uri $checkUri -Headers $headers
            Write-Log "Folder '$FolderName' already exists at '$ParentPath'"
            return $existingFolder
        }
        catch {
            # Folder doesn't exist, create it
            Write-Log "Folder '$FolderName' does not exist, creating it..."
        }

        # Creating new folder
        $encodedParentPath = [System.Web.HttpUtility]::UrlEncode($ParentPath)
        $uri = "https://graph.microsoft.com/v1.0/sites/$($siteId)/drive/root:/$encodedParentPath`:/children"
        $body = @{
            name = $FolderName
            folder = @{}
            "@microsoft.graph.conflictBehavior" = "rename"
        } | ConvertTo-Json

        try {
            $folder = Invoke-RestMethod -Method POST -Uri $uri -Body $body -ContentType "application/json" -Headers $Headers
            Write-Log "Successfully created folder '$FolderName'"
            return $folder
        }
        catch {
            Write-Log "Couldn't create folder $($FolderName) - $($_.exception.message)" -Level ERROR
            throw
        }
    }
    catch {
        Write-Log "Error with folder operation: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to upload file to SharePoint
function Upload-FileToSharePoint {
    param(
        [hashtable]$Headers,
        [string]$SiteId,
        [string]$DriveId,
        [string]$FolderPath,
        [string]$FileName,
        [byte[]]$FileContent
    )

    try {
        $encodedFolderPath = [System.Web.HttpUtility]::UrlEncode($FolderPath)
        $encodedFileName = [System.Web.HttpUtility]::UrlEncode($FileName)
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$encodedFolderPath/$encodedFileName`:/content"

        # Create a copy of headers and update content type
        $uploadHeaders = $Headers.Clone()
        $uploadHeaders['Content-Type'] = 'application/octet-stream'

        $response = Invoke-RestMethod -Uri $uri -Headers $uploadHeaders -Method PUT -Body $FileContent
        Write-Log "Successfully uploaded file '$FileName'"
    }
    catch {
        Write-Log "Error uploading file '$FileName': $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to get email messages from mailbox
function Get-MailboxMessages {
    param(
        [hashtable]$Headers,
        [string]$Mailbox,
        [datetime]$FromDate,
        [datetime]$ToDate
    )

    try {
        $fromDateStr = $FromDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $toDateStr = $ToDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

        Write-Log "Scanning mailbox '$Mailbox' from $fromDateStr to $toDateStr"

        # Build the select and top parameters
        $select = "id,subject,receivedDateTime,from,toRecipients,hasAttachments,conversationId"
        $top = "999"

        # Build URI with properly encoded filter
        $baseUri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages"
        $filterValue = "receivedDateTime ge $fromDateStr and receivedDateTime le $toDateStr"

        $queryParams = @(
            "`$filter=$([System.Uri]::EscapeDataString($filterValue))",
            "`$select=$select",
            "`$top=$top"
        )

        $uri = "$baseUri`?$($queryParams -join '&')"

        $allMessages = @()

        do {
            Write-Log "Query URI: $uri"
            $response = Invoke-GraphApi -Uri $uri -Headers $Headers
            $allMessages += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)

        Write-Log "Found $($allMessages.Count) messages in mailbox '$Mailbox'"
        return $allMessages
    }
    catch {
        Write-Log "Error retrieving messages from mailbox: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to get email content as HTML
function Get-EmailContentAsHTML {
    param(
        [hashtable]$Headers,
        [string]$Mailbox,
        [string]$MessageId
    )

    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId`?`$select=subject,receivedDateTime,conversationId,body,from,hasAttachments"
        $message = Invoke-RestMethod -Method GET -Uri $uri -Headers $Headers

        # Extract HTML body
        if ($message.body.contentType -eq "html") {
            $htmlContent = $message.body.content
        } else {
            # Convert text to HTML if needed
            $htmlContent = "<html><body><pre>" + [System.Web.HttpUtility]::HtmlEncode($message.body.content) + "</pre></body></html>"
        }

        return @{
            ContentBytes = [System.Text.Encoding]::UTF8.GetBytes($htmlContent)
            ConversationId = $message.conversationId
        }
    }
    catch {
        Write-Log "Error retrieving email HTML content: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to get email attachments
function Get-EmailAttachments {
    param(
        [hashtable]$Headers,
        [string]$Mailbox,
        [string]$MessageId
    )

    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId/attachments"
        $response = Invoke-GraphApi -Uri $uri -Headers $Headers
        return $response.value
    }
    catch {
        Write-Log "Error retrieving attachments: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# Function to get all attachments from all messages in a conversation
function Get-AllConversationAttachments {
    param(
        [hashtable]$Headers,
        [string]$Mailbox,
        [string]$ConversationId
    )

    try {
        # URL encode the entire filter string
        $filter = "conversationId eq '$ConversationId'"
        $encodedFilter = [System.Web.HttpUtility]::UrlEncode($filter)

        # Get ALL messages in this conversation thread
        $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages?`$filter=$encodedFilter&`$select=id,subject,receivedDateTime,hasAttachments"

        Write-Log "Retrieving all messages in conversation: $ConversationId"
        $messages = Invoke-RestMethod -Uri $uri -Headers $Headers -Method Get

        Write-Log "Found $($messages.value.Count) messages in conversation thread"

        $allAttachments = @()

        foreach ($msg in $messages.value) {
            if ($msg.hasAttachments) {
                Write-Log "Message from $($msg.receivedDateTime) has attachments" -Level INFO

                # Get attachments for this specific message
                $attUri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$($msg.id)/attachments"
                $attachments = Invoke-RestMethod -Uri $attUri -Headers $Headers -Method Get

                foreach ($att in $attachments.value) {
                    # Only process file attachments (not item attachments)
                    if ($att.'@odata.type' -eq '#microsoft.graph.fileAttachment') {
                        # Compute SHA256 hash for uniqueness detection
                        $sha256 = New-Object System.Security.Cryptography.SHA256Managed
                        $hashBytes = $sha256.ComputeHash([System.Convert]::FromBase64String($att.contentBytes))
                        $contentHash = ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ''

                        $allAttachments += [PSCustomObject]@{
                            MessageId = $msg.id
                            MessageSubject = $msg.subject
                            MessageReceived = $msg.receivedDateTime
                            AttachmentId = $att.id
                            FileName = $att.name
                            Size = $att.size
                            ContentType = $att.contentType
                            ContentBytes = $att.contentBytes
                            ContentHash = $contentHash
                        }

                        Write-Log "  Found attachment: $($att.name) ($($att.size) bytes)" -Level INFO
                    }
                }
            }
        }

        Write-Log "Total attachments found in conversation: $($allAttachments.Count)"
        return $allAttachments

    } catch {
        Write-Log "Error retrieving conversation attachments: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-UniqueAttachments {
    param(
        [array]$Attachments
    )

    $uniqueAttachments = @()
    $seenHashes = @{}

    foreach ($att in $Attachments) {
        $key = $att.ContentHash

        if (-not $seenHashes.ContainsKey($key)) {
            # First time seeing this file content
            $uniqueAttachments += $att
            $seenHashes[$key] = $true
            Write-Log "  Unique attachment: $($att.FileName) (Hash: $($key.Substring(0, 16))...)"
        } else {
            Write-Log "  Duplicate attachment skipped: $($att.FileName) from message dated $($att.MessageReceived)" -Level WARNING
        }
    }

    Write-Log "Filtered from $($Attachments.Count) to $($uniqueAttachments.Count) unique attachments"
    return $uniqueAttachments
}

Add-Type -AssemblyName System.IO.Compression

function Get-DocxText {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)

    try {
        $entry = $archive.GetEntry('word/document.xml')
        if (-not $entry) { return '' }

        $reader = [System.IO.StreamReader]::new($entry.Open())
        try {
            [xml]$xml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace(
            'w',
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        )

        ($xml.SelectNodes('//w:t', $ns) |
            ForEach-Object { $_.'#text' }) -join ' '
    }
    finally {
        $archive.Dispose()
    }
}


function Get-CompanyInfo {
    <#
    .SYNOPSIS
        Extracts company ICO/IC numbers and company names from email subject lines.

    .DESCRIPTION
        Parses various formats of Czech company identifiers (ICO, IC, IČ) and their associated company names.

    .PARAMETER Subject
        The email subject line to parse.

    .EXAMPLE
        Get-CompanyInfo "Re: CUT MYDLAR s.r.o. ič 01399284"

    .EXAMPLE
        Get-CompanyInfo "TS-MB stav s.r.o. iČo: 08424322 - Projektové financování - prodej"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Subject
    )

    process {
        $result = @{
            ICO = $null
            CompanyName = $null
            OriginalSubject = $Subject
        }

        # Regex pattern to match ICO/IC/IČ/IČO followed by number (with optional colon, space, or dot)
        # Matches: "ič 01399284", "iČo: 08424322", "IČ 68782004", "ico: 26907445"
        $icoPattern = '(?:i[cč]o?|IC)[:\s.]+(\d{8})'

        # Clean up the subject first
        $cleanSubject = $Subject -replace '^(?:Re:|Fwd?:|FW:|Subject:)\s*', ''
        $cleanSubject = $cleanSubject.Trim()

        if ($Subject -match $icoPattern) {
            $result.ICO = $Matches[1]

            # Try to extract company name - it's usually before the ICO
            # Pattern: captures text before the ICO pattern, typically company name with s.r.o., a.s., etc.
            $namePattern = '(.+?)\s+(?:i[cč]o?|IC)[:\s.]'

            if ($cleanSubject -match $namePattern) {
                $companyName = $Matches[1].Trim()
                # Remove trailing commas or dashes
                $companyName = $companyName -replace '[,\-]+$', ''
                $result.CompanyName = $companyName.Trim()
            }
        }
        else {
            # No ICO found, try to extract company name using common patterns
            # Look for company legal forms: s.r.o., a.s., v.o.s., k.s., družstvo, etc.
            $companyPattern = '([A-ZÁ-Ž].+?\s+(?:s\.r\.o\.|a\.s\.|v\.o\.s\.|k\.s\.|družstvo|spolek)\.?)(?:\s+\d{8})?'

            if ($cleanSubject -match $companyPattern) {
                $companyName = $Matches[1].Trim()
                $result.CompanyName = $companyName

                # Check if there's a standalone 8-digit number that might be ICO
                if ($cleanSubject -match '\b(\d{8})\b') {
                    $result.ICO = $Matches[1]
                }
            }
            # Fallback: if there's an 8-digit number, extract it and try to get text before it
            elseif ($cleanSubject -match '(.+?)\s+(\d{8})') {
                $result.CompanyName = $Matches[1].Trim()
                $result.ICO = $Matches[2]
            }
        }

        return [PSCustomObject]$result
    }
}

# Main script execution
try {
    Write-Log "Starting email processing script..."

    # Validate parameters
    Write-Log "Validating script parameters..."
    if (-not (Test-Path $CSVFolderPath)) {
        throw "CSV folder path not found: $CSVFolderPath"
    }

    # Load certificate from certificate store by thumbprint
    $cert = Get-CertificateByThumbprint -Thumbprint $CertificateThumbprint

    # Parse date range
    $fromDate = [datetime]::Parse($FromTimestamp)
    $toDate = [datetime]::Parse($ToTimestamp)
    Write-Log "Processing emails from $fromDate to $toDate"

    # Load all CSV files by year
    Write-Log "Loading CSV files from: $CSVFolderPath"
    $csvDictionary = Load-CSVsByYear -CSVFolderPath $CSVFolderPath
    Write-Log "CSV files loaded successfully"

    if ($Proxy) {
        [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($Proxy)
        [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        Write-Log "Proxy configured: $Proxy"
    }

    New-AadAuthenticationFactory -DefaultScopes "https://graph.microsoft.com/.default" `
        -ClientId $ClientId `
        -TenantId $TenantId `
        -X509Certificate $Cert

    $headers = Get-AadToken -Verbose -AsHashTable

    Write-Log "Getting SharePoint site information..."
    $siteUrl = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteUrl"
    $site = Invoke-GraphApi -Uri $siteUrl -Headers $headers
    $siteId = $site.id
    Write-Log "SharePoint Site ID: $siteId"

    # Get drive information
    $driveUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drive"
    $drive = Invoke-GraphApi -Uri $driveUrl -Headers $headers
    $driveId = $drive.id
    Write-Log "SharePoint Drive ID: $driveId"

    $notProcessedMails = @()
    $svjICOchecks = @()
    $compliant = 0
    $noncompliant = 0
    $compliance = New-Object System.Collections.arraylist


    # Process each mailbox
    foreach ($mailbox in $Mailboxes) {
        Write-Log "Processing mailbox: $mailbox"
        $notProcessedMailsEntry = [PSCustomObject]@{
            mailbox = $mailbox
            mails = @()
        }

        # Get messages from mailbox
        $messages = Get-MailboxMessages -Headers $headers -Mailbox $mailbox -FromDate $fromDate -ToDate $toDate
        $headers = Get-AadToken -Verbose -AsHashTable

        $totalMsgs = $messages.Count
        $msgNum = 1

        Write-Log "Found: $($totalMsgs) messages"
        foreach ($message in $messages){ #| Select-Object -First $processFirst) {
            
            write-host
            
            
            Write-Log "Processing email: $($message.subject)"
            
            write-host "Processing message $msgNum of $totalMsgs" -ForegroundColor Cyan
            #start-sleep -Seconds 3

            $svj = $false
            $conc = $false
                    
            
            
            # Retry logic: attempt to process each email up to 5 times
            $maxRetries = 5
            $retryCount = 0
            $emailProcessed = $false

            while (-not $emailProcessed -and $retryCount -lt $maxRetries) {
                try {
                    if ($retryCount -gt 0) {
                        Write-Log "Retry attempt $retryCount of $maxRetries for email: $($message.subject)" -Level WARNING
                    }

                    $headers = Get-AadToken -Verbose -AsHashTable
                    # Parse subject for case information
                    
                    
                    Write-Host $message.subject
                    #break

                    $caseInfo = Parse-EmailSubject -Subject $message.subject
                    write-host $caseInfo.IsValid


                    <#
                    if (-not $caseInfo.IsValid) {
                        $notValidEmail = [PSCustomObject]@{
                            reason = "Email subject does not match expected format, skipping"
                            mail = $message
                        }
                        $notProcessedMailsEntry.mails += $notValidEmail
                        Write-Log "Email subject does not match expected format, skipping..." -Level WARNING
                        $emailProcessed = $true  # Mark as processed to exit retry loop
                        break
                    }
                    #>

                    if (-not $caseInfo.IsValid) {
                        $notValidEmail = [PSCustomObject]@{
                            reason = "Email subject does not match expected format, for further check"
                            mail = $message
                        }
                        
                        $compl = [PSCustomObject]@{
                            compl = $false
                            subject = $message.Subject
                        }
                        
                        #$notProcessedMailsEntry.mails += $notValidEmail
                        Write-Log "Email subject does not match expected format, for further check" -Level WARNING
                        #$emailProcessed = $true  # Mark as processed to exit retry loop
                        $noncompliant ++
                        [void]$compliance.Add($compl)

                        

                    } else {
                    
                        $notValidEmail = [PSCustomObject]@{
                            reason = "Email subject matches expected format, already processed"
                            mail = $message
                        }

                        $compl = [PSCustomObject]@{
                            compl = $true
                            subject = $message.Subject
                        }

                        $notProcessedMailsEntry.mails += $notValidEmail
                        Write-Log "Email subject matches expected format, skipping..." -Level WARNING
                        $emailProcessed = $true  # Mark as processed to exit retry loop
                        $compliant ++
                        [void]$compliance.Add($compl)

                        break
                    
                    
                    }

                    write-host "Further check" -ForegroundColor Red
                    
                    #Write-Log "Extracted case info - Number: $($caseInfo.Number), CaseId: $($caseInfo.CaseId), Year: $($caseInfo.Year)"

                    <#
                    if($caseInfo.Year -in $yearsToIgnore)
                    {
                        Write-Log "Ignoring :$($caseInfo.Number) from year: $($caseInfo.Year) - since its within ignored years" -Level WARNING
                        continue
                    }
                    #>

                    
                    <#unknown for stochastic search
                    # Search for case in CSV dictionary by year and get all metadata
                    $csvMetadata = Find-CaseInCSV -CSVDictionary $csvDictionary -Year $caseInfo.Year -CaseId $caseInfo.CaseId -ColumnName $CSVCaseIdColumn
                    #>
                    
                    
                    <#
                    if (-not $csvMetadata) {
                        Write-Log "Case not found in CSV for year $($caseInfo.Year), skipping..." -Level WARNING
                        $emailProcessed = $true  # Mark as processed to exit retry loop
                        break
                    }
                    #>


                    <#
                    # Create folder structure: {Year}-{CaseId}/
                    $yearCaseFolder = "$($caseInfo.Year)-$($caseInfo.CaseId)"
                    $basePath = $SharePointBasePath.TrimStart('/')
                    $caseFolderPath = "$basePath/$yearCaseFolder"

                    # Create year-caseId folder (or get if exists)
                    Write-Log "Creating/getting case folder: $yearCaseFolder"
                    $caseFolder = New-SharePointFolder -Headers $headers -SiteId $siteId `
                        -DriveId $driveId -ParentPath $basePath -FolderName $yearCaseFolder

                    # Create Mails subfolder (or get if exists)
                    Write-Log "Creating/getting Mails subfolder"
                    $mailsFolder = New-SharePointFolder -Headers $headers -SiteId $siteId `
                        -DriveId $driveId -ParentPath $caseFolderPath -FolderName "Mails"

                    # Create Attachments subfolder (or get if exists)
                    Write-Log "Creating/getting Attachments subfolder"
                    $attachmentsFolder = New-SharePointFolder -Headers $headers -SiteId $siteId `
                        -DriveId $driveId -ParentPath $caseFolderPath -FolderName "Attachments"

                    # Save metadata
                    $metadataJson = $csvMetadata | ConvertTo-Json -Depth 10
                    $metadataBytes = [System.Text.Encoding]::UTF8.GetBytes($metadataJson)
                    $metadataFileName = "$($caseInfo.CaseId)-metadata.json"

                    Write-Log "Uploading $($metadataFileName)"
                    Upload-FileToSharePoint -Headers $headers -SiteId $siteId `
                        -DriveId $driveId -FolderPath $caseFolderPath `
                        -FileName $metadataFileName -FileContent $metadataBytes

                    #>
                    
                    
                    # Get and save email content as HTML
                    Write-Log "Getting email content as HTML"
                    $emailData = Get-EmailContentAsHTML -Headers $headers -Mailbox $mailbox -MessageId $message.id
                    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                    $emailFileName = "$($caseInfo.Number)-$($caseInfo.CaseId)-$($caseInfo.Year)_$($message.receivedDateTime.Replace(':','-')).html"

                    $emailProcessed = $true  # Mark as processed to exit retry loop
                    
                    #$emailData | Out-Host
                    
                    if ( ($message.subject -notmatch '(Byto\w+)') -and ($message.subject -notmatch '(SVJ \w+|Společenství \w+)') ) {
                        
                        Write-Host "Extended detection" -ForegroundColor Green
                        $svj = $true
                        $svjICOcheck = [PSCustomObject]@{
                            Subject = $message.subject
                            ICO = ""
                            foundInSubject = $false
                            foundInMessage = $false
                            foundInAttachment = $false
                            attachmentName = ""
                            messageconversationId = $message.conversationId
                            messageid = $message.id
                            receivedDateTime = $message.receivedDateTime
                            fromAdd = ($message.from.emailAddress | foreach {$_.address}) -join ";"
                            fromName = ($message.from.emailAddress | foreach {$_.name}) -join ";"
                            toAdd = ($message.torecipients.emailAddress | foreach {$_.address}) -join ";"
                            toName = ($message.torecipients.emailAddress | foreach {$_.name}) -join ";"
                            cleanICO = ""
                            companyName = ""

                        }

                        #$mtchs = [regex]::Matches($message.subject,"(I)(C|Č)\D+\d{8}")

                        $catch = ""
                        if ($message.subject -match 'I\u010C.*[ ]*((?:[0-9][ ]*){8})(?![0-9])') {
                            $catch = 'IČO ' + ($Matches[1] -replace '[^0-9]', '')
                        }
                        
                        if ($catch -eq ""){
                            if ($message.subject -match '(?<!\d)((?:[0-9][ ]*){8})(?![0-9])(?!\d)') {
                            $catch = 'IČO ' + ($Matches[1] -replace '[^0-9]', '')
                            }
                        }
                            
                        if ($catch -ne ""){
                                
                            $svjICOcheck.ICO = $catch
                            $svjICOcheck.foundInSubject = $true
                            write-host "Found in subject" -ForegroundColor Magenta
                            #Start-Sleep -Seconds 2
                            
                        } else {
                            
                            $emailData['ContentBytes'] | Set-Content test.html -Encoding Byte
                            #Start-Sleep -Seconds 2

                            $messaget = [System.Text.Encoding]::UTF8.GetString($emailData['ContentBytes'])

                            $catch = ""
                            if ($messaget -match 'I\u010C.*[ ]*((?:[0-9][ ]*){8})(?![0-9])') {
                                $catch = 'IČO ' + ($Matches[1] -replace '[^0-9]', '')
                            }
                        
                            if ($catch -eq ""){
                                if ($messaget -match '(?<!\d)((?:[0-9][ ]*){8})(?![0-9])(?!\d)') {
                                $catch = 'IČO ' + ($Matches[1] -replace '[^0-9]', '')
                                }
                            }

                            #$mtchs = [regex]::Matches($message,"(I)(C|Č)\D+\d{8}")
                            #$mtchs[0] | Out-Host
                            #Start-Sleep -Seconds 2
                            
                            if ($catch -ne ""){
                                
                                $svjICOcheck.ICO = $catch
                                $svjICOcheck.foundInMessage = $true
                                write-host "Found in message" -ForegroundColor Magenta
                                #Start-Sleep -Seconds 2
                            
                            
                            } else {
                               $conc = $true
                            }
                        


                        }

                        

                    }
                    #break
                    
                    <#
                    Write-Log "Uploading email as HTML: $emailFileName"
                    
                    
                    Upload-FileToSharePoint -Headers $headers -SiteId $siteId `
                        -DriveId $driveId -FolderPath "$caseFolderPath/Mails" `
                        -FileName $emailFileName -FileContent $emailData.ContentBytes
                    
                    #>
                    if ($conc){
                        # Process ALL attachments from the entire conversation thread
                        Write-Log "Processing attachments from entire conversation thread..."
                        $allAttachments = Get-AllConversationAttachments -Headers $headers -Mailbox $mailbox -ConversationId $emailData.ConversationId

                        if ($allAttachments.Count -gt 0) {
                            Write-Log "Found $($allAttachments.Count) total attachments across all messages in conversation"

                            # Get only unique attachments based on content hash
                            $uniqueAttachments = Get-UniqueAttachments -Attachments $allAttachments
                            Write-Log "Uploading $($uniqueAttachments.Count) unique attachments"

                            
                            $fldName = Get-Date -Format 'yyyyMMddHHmmss'
                            $fld = New-Item -Path ".\testAttachs\$fldName" -ItemType Directory
                            
                            
                            foreach ($attachment in $uniqueAttachments) {
                                
                                if ($attachment.FileName -match 'docx'){
                                    Write-Log "Saving attachment: $($attachment.FileName) from message dated $($attachment.MessageReceived)"
                                    $attachmentBytes = [System.Convert]::FromBase64String($attachment.ContentBytes)
                                    $fileName = ($attachment.FileName.Replace("�",""))
                                    $fileName = $fileName.Replace("Ã­","A")
                                    $fileName = $fileName.Replace("Ã","A")
                                     $fileName = $fileName.Replace(":","-")
                            
                            
                                    $attachmentBytes | Set-Content "$($fld.FullName)\$fileName" -Encoding Byte
                                    Remove-Variable attachmentBytes -Force
                                }
                            
                                <#
                                Upload-FileToSharePoint -Headers $headers -SiteId $siteId `
                                    -DriveId $driveId -FolderPath "$caseFolderPath/Attachments" `
                                    -FileName $fileName -FileContent $attachmentBytes

                                #>

                            
                            }


                            #analyze DOCX

                            $root   = $fld.FullName
                            $search = 'IČO'

                            $attachFiles = Get-ChildItem $root -Filter '*.docx' -File -Recurse
                                ForEach ($attachFile in $attachFiles) {
                                    $text = Get-DocxText $attachFile.FullName

                                    <#
                                    if ($text.IndexOf($search, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                                        [PSCustomObject]@{
                                            Path  = $_.FullName
                                            Match = $search
                                        }
                                    }
                                    #>

                                    $catch = ""
                                    if ($text -match 'I\u010C.*[ ]*((?:[0-9][ ]*){8})(?![0-9])') {
                                        $catch = 'IČO ' + ($Matches[1] -replace '[^0-9]', '')
                                    }
                                    
                                    
                                    #$mtchs = [regex]::Matches($text,"(I)(C|Č).+\d{8}")
                                    #$mtchs[0] | Out-Host
                                    write-host $catch -ForegroundColor Magenta
                                    #Start-Sleep -Seconds 2
                            
                                    if ($catch -ne ""){
                                
                                        $svjICOcheck.ICO = $catch
                                        $svjICOcheck.foundInAttachment = $true
                                        $svjICOcheck.attachmentName = $attachFile.Name
                                        write-host "Found in attachment" -ForegroundColor Magenta
                                        #Start-Sleep -Seconds 2
                                        break
                            
                            
                                    }
                                }

                        
                            Write-Host "Attachments written" -ForegroundColor Cyan
                            Start-Sleep -Milliseconds 500

                            Get-ChildItem $root -File -Recurse | Remove-Item -Force

                            Remove-Variable attachFiles -Force
                            
                            [gc]::Collect()


                        } else {
                            Write-Log "No attachments found in this conversation thread"
                        }
                    }


                    Write-Log "Successfully processed email for case $($caseInfo.CaseId)"
                    $emailProcessed = $true  # Mark as successfully processed

                    if ($svj){
                        $svjICOcheck | Out-Host
                    }


                } catch {
                    $retryCount++
                    Write-Log "Error processing email (Attempt $retryCount of $maxRetries): $($_.Exception.Message)" -Level ERROR
                    Write-Log "Email subject: $($message.subject)" -Level ERROR
                    Write-Log "Email received: $($message.receivedDateTime)" -Level ERROR
                    Write-Log $_.ScriptStackTrace -Level ERROR

                    if ($retryCount -lt $maxRetries) {
                        Write-Log "Waiting 20 seconds before retry..." -Level WARNING
                        Start-Sleep -Seconds 20
                    } else {
                        Write-Log "Max retries reached for email: $($message.subject). Moving to next email." -Level ERROR
                        Write-Log "Last Known stamp: $($message.receivedDateTime)" -Level ERROR
                    }
                }
            }

            


            if ($svj){
            
            <#
                if ($svjICOcheck.ICO -match '.+'){
                    $cmp = Get-CompanyInfo -Subject $svjICOcheck.ICO

                    if ($cmp.CompanyName -match '.+'){
                        $svjICOcheck.cleanICO = $cmp.ico
                        $svjICOcheck.companyName = $cmp.companyname

                    }
                }

            #>

            $svjICOchecks += $svjICOcheck
            $svjICOchecks | Export-Csv svjICOchecks.csv -NoTypeInformation -Force -Encoding Unicode
            Remove-Variable svjICOcheck -Force

            }
            
            $msgNum++

            if (($msgNum % 100) -eq 0 ) {[gc]::Collect()}
        }
        
        $notProcessedMails += $notProcessedMailsEntry
        

        
    }

    Write-Log "Email processing completed successfully!"
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level ERROR
    Write-Log "Last Known stamp: $($message.receivedDateTime)"
    $lastKnownStamp = $message.receivedDateTime
    Write-Log $_.ScriptStackTrace -Level ERROR
    exit 1
}
