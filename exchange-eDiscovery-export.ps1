<#
.SYNOPSIS
    Exports calendar, contacts, and tasks from Exchange Online mailboxes using Microsoft Graph API.

.DESCRIPTION
    This script exports calendar items, contacts, and/or tasks from Exchange Online mailboxes
    using Microsoft Graph API directly without requiring Exchange Online PowerShell modules.
    It leverages eDiscovery capabilities to perform the exports through REST API calls.

.PARAMETER ClearTokenCache
    Clears all Microsoft Graph authentication tokens and cached states.

.PARAMETER VerboseTokenClearing
    Shows detailed information when clearing authentication tokens.

.PARAMETER EmailAddress
    The email address of the mailbox to export from.
    Default: "NestorW@M365x61250205.OnMicrosoft.com"

.PARAMETER AppID
    The application (client) ID of the app registration in Azure AD.
    Default: "5baa1427-1e90-4501-831d-a8e67465f0d9"

.PARAMETER TenantId
    The Azure AD tenant ID.
    Default: "85612ccb-4c28-4a34-88df-a538cc139a51"

.PARAMETER AuthMode
    Authentication mode to use when connecting to Microsoft Graph.
    Valid values: "Cert", "Secret"
    Default: "cert"

.PARAMETER certlocation
    Location of the certificate to use for authentication.
    Valid values: "Cert:\LocalMachine\My\", "Cert:\My\"
    Default: "Cert:\LocalMachine\My\"

.PARAMETER CertificateThumbprint
    The thumbprint of the certificate to use for authentication.
    Default: "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"

.PARAMETER ClientSecret
    The client secret to use when AuthMode is "Secret".

.PARAMETER Version
    The Microsoft Graph API version to use.
    Valid values: "prod" (v1.0), "beta"
    Default: "beta"

.PARAMETER ContentType
    The type of content to export.
    Valid values: "all", "calendar", "contacts", "tasks"
    Default: "all"

.PARAMETER Operation
    The operation to perform.
    Valid values: "menu", "create", "export", "download"
    Default: "menu"

.PARAMETER CaseId
    The ID of an existing eDiscovery case to use.

.PARAMETER SearchId
    The ID of an existing eDiscovery search to use.

.PARAMETER ExportId
    The ID of an existing export to download.

.PARAMETER DownloadPath
    The path where exported files should be saved.
    Default: "C:\temp"

.PARAMETER DebugOutput
    Enables verbose debug output during script execution.

.EXAMPLE
    .\exchange-cal-export.ps1 -EmailAddress user@domain.com
    Exports all content types for the specified user.

.EXAMPLE
    .\exchange-cal-export.ps1 -EmailAddress user@domain.com -ContentType calendar
    Exports only calendar items for the specified user.

.EXAMPLE
    .\exchange-cal-export.ps1 -ClearTokenCache
    Clears all authentication tokens and exits.

.EXAMPLE
    .\exchange-cal-export.ps1 -Version prod
    Uses v1.0 endpoints for all API calls.

.NOTES
    - Requires Microsoft Graph PowerShell module installed (Install-Module Microsoft.Graph)
    - Requires MSAL.PS module for Microsoft's official download method
    - App registration must have eDiscovery.ReadWrite.All permission with admin consent granted
    - App registration must have MicrosoftPurviewEDiscovery API permissions for downloads

Author: Mike Lee | Dempsey Dunkin | Ranjit Sharma
Date: 9/15/2025


.REQUIREMENTS

 - Microsoft Graph PowerShell module installed (Install-Module Microsoft.Graph)
 - MSAL.PS module (for official Microsoft download method)
 - App registration with eDiscovery.ReadWrite.All permission with admin consent granted
 - App registration with MicrosoftPurviewEDiscovery API permissions for downloads
https://techcommunity.microsoft.com/blog/microsoft-security-blog/getting-started-with-the-ediscovery-apis/4407597
https://learn.microsoft.com/en-us/graph/security-ediscovery-appauthsetup

#>

param (
    [Parameter(Mandatory = $false)]
    [switch]$ClearTokenCache,
    
    [Parameter(Mandatory = $false)]
    [switch]$VerboseTokenClearing,
    
    [Parameter(Mandatory = $false)]
    #[string]$EmailAddress = "admin@M365x61250205.onmicrosoft.com",
    [string]$EmailAddress = "NestorW@M365x61250205.OnMicrosoft.com",
    
    [Parameter(Mandatory = $false)]
    [string]$AppID = "5baa1427-1e90-4501-831d-a8e67465f0d9",
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "85612ccb-4c28-4a34-88df-a538cc139a51",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Cert", "Secret")]
    [string]$AuthMode = "cert",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Cert:\LocalMachine\My\", "Cert:\My\")]
    [string]$certlocation = "Cert:\LocalMachine\My\",
   
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = "",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("prod", "beta")]
    [string]$Version = "beta",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("all", "calendar", "contacts", "tasks")]
    [string]$ContentType = "all",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("menu", "create", "export", "download")]
    [string]$Operation = "menu",
    
    [Parameter(Mandatory = $false)]
    [string]$CaseId,
    
    [Parameter(Mandatory = $false)]
    [string]$SearchId,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportId,
    
    [Parameter(Mandatory = $false)]
    [string]$DownloadPath = "C:\temp",
    
    [Parameter(Mandatory = $false)]
    [switch]$DebugOutput
)

# Ensure $global:Operation is always set for use in functions
$global:Operation = $Operation

# Global variable to track if we're in initialization mode
$global:IsScriptInitializing = $true

# Start logging
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = "$env:TEMP\Exchange_Calendar_Export_$currentDateTime.log"
Write-Output "Starting Exchange Calendar Export Script at $(Get-Date)" | Out-File -FilePath $logPath

# Function to write log entries
Function Write-LogEntry {
    param(
        [string] $LogName,
        [Parameter(Mandatory = $true)]
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"
    )
    
    if ($null -ne $LogEntryText) {
        if ($null -ne $LogName) {
            "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -Append
        }
    }
}

# Function to clear Microsoft Graph authentication tokens
Function Clear-GraphTokenCache {
    param(
        [switch] $Verbose
    )
    
    Write-Host "Clearing all Microsoft Graph authentication tokens and cached states..." -ForegroundColor Yellow
    
    try {
        # Disconnect from Microsoft Graph if connected
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            if ($Verbose) { Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green }
        }
        catch {
            if ($Verbose) { Write-Host "No active Microsoft Graph connection to disconnect" -ForegroundColor Yellow }
        }
        
        # Force garbage collection to clean up any references
        [System.GC]::Collect()
        
        Write-Host "All Microsoft Graph authentication tokens have been cleared." -ForegroundColor Green
        Write-Host "You can now reconnect with updated permissions." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error clearing tokens: $_" -ForegroundColor Red
        return $false
    }
}

# Process token clearing request if specified
if ($ClearTokenCache) {
    Clear-GraphTokenCache -Verbose:$VerboseTokenClearing
    if (-not $VerboseTokenClearing) {
        Write-Host "To see detailed information about token clearing, use: .\exchange-cal-export.ps1 -ClearTokenCache -VerboseTokenClearing" -ForegroundColor Yellow
    }
    exit 0
}

# Import required modules
Write-Host "Importing Microsoft Graph modules..." -ForegroundColor Yellow
try {
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Security -Force -ErrorAction Stop
    Write-Host "Microsoft Graph modules imported successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error importing Microsoft Graph modules: $_" -ForegroundColor Red
    Write-Host "Please install Microsoft Graph PowerShell SDK: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Clear tokens at the beginning of each run to ensure fresh authentication
Write-Host "Clearing existing tokens to ensure fresh authentication..." -ForegroundColor Yellow
Clear-GraphTokenCache
Write-Host "Tokens cleared. Will authenticate with fresh credentials." -ForegroundColor Green

# Authentication parameters (kept for backward compatibility)
#if (-not $AppID) { $AppID = "5baa1427-1e90-4501-831d-a8e67465f0d9" }
#if (-not $TenantId) { $TenantId = "85612ccb-4c28-4a34-88df-a538cc139a51" }
#if ($AuthMode -eq "Cert" -and -not $CertificateThumbprint) { $CertificateThumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" }

# Authentication with Microsoft Graph
Write-Host "Authenticating with Microsoft Graph API using $AuthMode mode..." -ForegroundColor Yellow
try {
    # Connect to Microsoft Graph - App-Only Authentication
    if ($AuthMode -eq "Cert") {
        # Certificate-based authentication
        Write-Host "Using certificate-based authentication with thumbprint: $CertificateThumbprint" -ForegroundColor Yellow
        
        # Connect using the certificate (permissions are set in the app registration)
        Connect-MgGraph -ClientId $AppID -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    }
    else {
        # Client Secret authentication
        Write-Host "Using client secret authentication" -ForegroundColor Yellow
        
        if ([string]::IsNullOrEmpty($ClientSecret)) {
            throw "Client Secret is required when using Secret authentication mode"
        }
        
        # Create a PSCredential object for the client secret
        $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $clientSecretCredential = New-Object System.Management.Automation.PSCredential($AppID, $secureClientSecret)
        
        # Connect using the client secret credential
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential
    }
    
    # Display connection information
    $context = Get-MgContext
    Write-Host "Successfully connected to Microsoft Graph API" -ForegroundColor Green
    Write-Host "Connected as application: $($context.AppName)" -ForegroundColor Green
    Write-Host "Authorization Scheme: $($context.AuthType)" -ForegroundColor Green
    Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor Green
    
    # Mark initialization complete - safe to perform downloads now
    $global:IsScriptInitializing = $false
    
    # Verify scopes/permissions
    $scopes = $context.Scopes
    Write-Host "Permissions:" -ForegroundColor Yellow
    if ($scopes -and $scopes.Count -gt 0) {
        foreach ($scope in $scopes) {
            Write-Host "  - $scope" -ForegroundColor Green
        }
        
        # With application permissions, the individual permissions
        # won't be visible in the token, but they should be applied if granted in Azure
        Write-Host "Using .default scope which includes all granted application permissions" -ForegroundColor Green
    }
    else {
        Write-Host "  No permissions/scopes visible in token" -ForegroundColor Yellow
        Write-Host "  This is expected for applications using certificate or secret authentication" -ForegroundColor Yellow
        
        # Add diagnostic check for admin consent and permissions
        Write-Host "Checking if required permissions for eDiscovery are granted..." -ForegroundColor Yellow
        try {
            # Test access to a non-security endpoint that requires basic permissions
            $testBasicUrl = "https://graph.microsoft.com/v1.0/organization"
            Invoke-MgGraphRequest -Method GET -Uri $testBasicUrl -ErrorAction Stop
            Write-Host "  Basic Microsoft Graph permissions confirmed ?" -ForegroundColor Green
        }
        catch {
            Write-Host "  WARNING: Even basic Microsoft Graph permissions appear to have issues!" -ForegroundColor Red
            Write-Host "  Error: $_" -ForegroundColor Red
        }
        
        Write-Host "For eDiscovery, your app registration MUST have:" -ForegroundColor Yellow
        Write-Host "  - eDiscovery.ReadWrite.All (Application permission)" -ForegroundColor Yellow
        Write-Host "  - Admin consent granted" -ForegroundColor Yellow
        Write-Host "  - eDiscovery Roles properly configured" -ForegroundColor Yellow
        Write-Host "For more information on setting up app permissions, visit:" -ForegroundColor Yellow
        Write-Host "  - https://learn.microsoft.com/en-us/graph/security-ediscovery-appauthsetup" -ForegroundColor Yellow
    }
    
    Write-LogEntry -LogName $logPath -LogEntryText "Successfully authenticated with Microsoft Graph API using $AuthMode mode" -LogLevel "INFO"
}
catch {
    Write-LogEntry -LogName $logPath -LogEntryText "Authentication failed with Microsoft Graph API: $_" -LogLevel "ERROR"
    Write-Host "Authentication failed with Microsoft Graph API: $_" -ForegroundColor Red
    throw
}

# Function to get folder IDs from Exchange mailbox
function Get-FolderIDs {
    param(
        [Parameter(Mandatory = $false)]
        [string]$UserEmailAddress = $EmailAddress
    )
    
    try {
        Write-Host "Connecting to Exchange Online to retrieve folder IDs..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ErrorAction Stop
        
        $folderQueries = @()
        $folderStatistics = Get-MailboxFolderStatistics $UserEmailAddress
        foreach ($folderStatistic in $folderStatistics) {
            $folderId = $folderStatistic.FolderId
            $folderPath = $folderStatistic.FolderPath
            $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
            $nibbler = $encoding.GetBytes("0123456789ABCDEF")
            $folderIdBytes = [Convert]::FromBase64String($folderId)
            $indexIdBytes = New-Object byte[] 48
            $indexIdIdx = 0
            $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
            $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
            $folderStat = New-Object PSObject
            Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
            Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
            $folderQueries += $folderStat
        }
        
        Write-Host "-----Exchange Folders-----" -ForegroundColor Green
        $relevantFolders = $folderQueries | Where-Object { $_.folderpath -eq "/Calendar" -or $_.folderpath -eq "/Contacts" -or $_.folderpath -eq "/Tasks" }
        $relevantFolders | Format-Table -AutoSize
        
        # Create a hashtable to return the folder queries
        $result = @{
            CalendarQuery = ($folderQueries | Where-Object { $_.folderpath -eq "/Calendar" }).FolderQuery
            ContactsQuery = ($folderQueries | Where-Object { $_.folderpath -eq "/Contacts" }).FolderQuery
            TasksQuery    = ($folderQueries | Where-Object { $_.folderpath -eq "/Tasks" }).FolderQuery
        }
        
        return $result
    }
    catch {
        Write-Host "Error retrieving folder IDs: $_" -ForegroundColor Red
        Write-LogEntry -LogName $logPath -LogEntryText "Error retrieving folder IDs: $_" -LogLevel "ERROR"
        
        # Return empty values if unable to retrieve from Exchange
        # No default values since folder IDs are unique per user
        Write-Host "Unable to retrieve folder IDs from Exchange. Please check connection and permissions." -ForegroundColor Red
        return @{
            CalendarQuery = $null
            ContactsQuery = $null
            TasksQuery    = $null
        }
    }
    finally {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore errors on disconnect
        }
    }
}

# Initialize variables that will be populated when needed
$ContactsFolderQuery = $null
$TasksFolderQuery = $null
$CalFolderQuery = $null

# Create search name using the user's email (common variable used in multiple places)
$userSearchName = ($EmailAddress.split("@")[0]).replace(".", "_")
$searchName = ($userSearchName) + '_Mailbox'

Write-LogEntry -LogName $logPath -LogEntryText "Starting export process for user: $EmailAddress with search name: $searchName" -LogLevel "INFO"

# Function to make direct REST API call for all Graph operations
function Invoke-GraphAPIRequest {
    param (
        [string]$Method = "GET",
        [string]$Uri,
        [object]$Body = $null,
        [string]$ApiVersion = $null, # Will be determined by global Version parameter
        [switch]$ForceBeta # Force using beta endpoint for some operations
    )
    
    # If ApiVersion is not specified, use the global Version parameter
    if ([string]::IsNullOrEmpty($ApiVersion)) {
        $ApiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
    }
    
    # Honor the Version parameter setting, don't force beta for security APIs
    $useApiVersion = $ApiVersion
    if ($ForceBeta) {
        $useApiVersion = "beta"
        Write-Host "ForceBeta switch specified. Using beta endpoint for this call." -ForegroundColor Yellow
    }
    
    # Make sure the URI includes the API version if not already specified
    if (-not $Uri.Contains("graph.microsoft.com/")) {
        $Uri = "https://graph.microsoft.com/$useApiVersion$Uri"
        Write-Host "Using $useApiVersion endpoint for API call: $Uri" -ForegroundColor Cyan
    }
    
    try {
        Write-Host "Making API request to: $Uri" -ForegroundColor Cyan
        
        if ($Method -eq "GET") {
            $response = Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
        else {
            $response = Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body
        }
        
        return $response
    }
    catch {
        $errorDetails = "Error in Graph API call to $Uri`: $_"
        Write-LogEntry -LogName $logPath -LogEntryText $errorDetails -LogLevel "ERROR"
        
        # Try to get more detailed error information
        try {
            $errorObj = $null
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $errorObj = $_.ErrorDetails.Message | ConvertFrom-Json
            }
            
            if ($errorObj -and $errorObj.error) {
                Write-Host "Error details:" -ForegroundColor Red
                Write-Host "  Code: $($errorObj.error.code)" -ForegroundColor Red
                Write-Host "  Message: $($errorObj.error.message)" -ForegroundColor Red
                
                # Check for specific known errors
                if ($errorObj.error.code -eq "Authorization_RequestDenied" -or 
                    $errorObj.error.message -like "*permission*" -or 
                    $errorObj.error.message -like "*unauthorized*") {
                    
                    Write-Host "`nPermission issue detected!" -ForegroundColor Red
                    Write-Host "The app may be missing required permissions or admin consent." -ForegroundColor Red
                    
                    # If this is a security API call, suggest trying beta endpoint
                    if ($Uri.Contains("/security/") -and $Uri.Contains("/v1.0/")) {
                        Write-Host "`nTrying same call with beta endpoint..." -ForegroundColor Yellow
                        $betaUri = $Uri.Replace("/v1.0/", "/beta/")
                        
                        try {
                            if ($Method -eq "GET") {
                                $betaResponse = Invoke-MgGraphRequest -Method $Method -Uri $betaUri
                            }
                            else {
                                $betaResponse = Invoke-MgGraphRequest -Method $Method -Uri $betaUri -Body $Body
                            }
                            
                            Write-Host "Beta endpoint call succeeded!" -ForegroundColor Green
                            return $betaResponse
                        }
                        catch {
                            Write-Host "Beta endpoint call also failed: $_" -ForegroundColor Red
                        }
                    }
                }
                # Handle case where security API is only available in beta
                elseif ($errorObj.error.code -eq "ResourceNotFound" -and 
                    $Uri.Contains("/security/") -and 
                    $Uri.Contains("/v1.0/")) {
                    
                    Write-Host "Resource not found. Some security APIs are only available in beta." -ForegroundColor Yellow
                    Write-Host "Trying same call with beta endpoint..." -ForegroundColor Yellow
                    
                    $betaUri = $Uri.Replace("/v1.0/", "/beta/")
                    
                    try {
                        if ($Method -eq "GET") {
                            $betaResponse = Invoke-MgGraphRequest -Method $Method -Uri $betaUri
                        }
                        else {
                            $betaResponse = Invoke-MgGraphRequest -Method $Method -Uri $betaUri -Body $Body
                        }
                        
                        Write-Host "Beta endpoint call succeeded!" -ForegroundColor Green
                        return $betaResponse
                    }
                    catch {
                        Write-Host "Beta endpoint call also failed: $_" -ForegroundColor Red
                    }
                }
            }
            else {
                Write-Host "Error in Graph API call: $_" -ForegroundColor Red
                Write-Host "Status Code: $($_.Exception.Response.StatusCode.value__)" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Error in Graph API call: $_" -ForegroundColor Red
        }
        
        throw $_
    }
}

# Test the connection with a simple Graph API call before trying security APIs
Write-LogEntry -LogName $logPath -LogEntryText "Testing basic Microsoft Graph connectivity" -LogLevel "INFO"
try {
    # Use a more appropriate endpoint for app-only auth testing - organization details
    $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
    $testApiUrl = "https://graph.microsoft.com/$apiVersion/organization"
    $testResponse = Invoke-MgGraphRequest -Method GET -Uri $testApiUrl
    
    Write-Host "Using $apiVersion endpoint for Graph API calls based on Version parameter" -ForegroundColor Yellow
    
    if ($testResponse.value -and $testResponse.value.Count -gt 0) {
        Write-Host "Basic Graph connectivity confirmed! Organization info retrieved successfully." -ForegroundColor Green
        Write-Host "Organization: $($testResponse.value[0].displayName)" -ForegroundColor Green
    }
    else {
        Write-Host "Basic Graph connectivity confirmed, but no organization details returned." -ForegroundColor Green
    }
}
catch {
    Write-LogEntry -LogName $logPath -LogEntryText "Error testing basic Graph connectivity: $_" -LogLevel "WARNING"
    Write-Host "Error testing basic Graph connectivity: $_" -ForegroundColor Yellow
    Write-Host "This may indicate a problem with the app permissions or connectivity." -ForegroundColor Yellow
    Write-Host "Continuing anyway to see if security APIs work..." -ForegroundColor Yellow
}

# Check permissions by testing access to eDiscovery cases
Write-LogEntry -LogName $logPath -LogEntryText "Checking permissions by testing access to eDiscovery cases" -LogLevel "INFO"

try {
    # Use Invoke-MgGraphRequest to list cases (just one to check permissions)
    # Use the endpoint version based on the Version parameter
    $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
    Write-Host "Using $apiVersion endpoint for eDiscovery API based on Version=$Version parameter" -ForegroundColor Yellow
    $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases?`$select=displayName,id,status&`$top=1"
    
    Write-Host "Making request to: $apiUrl" -ForegroundColor Cyan
    try {
        $casesResponse = Invoke-MgGraphRequest -Method GET -Uri $apiUrl -ErrorAction Stop
        
        if ($null -ne $casesResponse.value) {
            Write-LogEntry -LogName $logPath -LogEntryText "Successfully verified permissions to access eDiscovery cases" -LogLevel "INFO"
            Write-Host "Successfully verified permissions to access eDiscovery cases" -ForegroundColor Green
        }
        else {
            Write-LogEntry -LogName $logPath -LogEntryText "Permission check returned no cases - may indicate permission issues" -LogLevel "WARNING"
            Write-Host "Permission check returned no cases - may indicate permission issues" -ForegroundColor Yellow
        }
    }
    catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            Write-Host "ERROR: Received 401 Unauthorized when accessing eDiscovery APIs." -ForegroundColor Red
            Write-Host "This is likely due to missing permissions or admin consent." -ForegroundColor Red
            Write-LogEntry -LogName $logPath -LogEntryText "Received 401 Unauthorized for eDiscovery API. Permission issue detected." -LogLevel "ERROR"
            throw "Authentication failed for eDiscovery APIs. Please ensure proper permissions are granted."
        }
        else {
            throw
        }
    }
}
catch {
    Write-LogEntry -LogName $logPath -LogEntryText "Error accessing eDiscovery cases. This may be a permissions issue: $_" -LogLevel "ERROR"
    Write-Host "Details: $_" -ForegroundColor Red
}

# Function to create a new eDiscovery case and search
Function New-CalendarExportCase {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentTypeFilter = "all"
    )
    
    # Create unique search name using the user's email and timestamp
    $userSearchName = ($UserEmail.split("@")[0]).replace(".", "_")
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $searchName = ($userSearchName) + '_Mailbox_' + $timestamp
    
    Write-Host "Creating new eDiscovery case for user: $UserEmail" -ForegroundColor Yellow
    Write-LogEntry -LogName $logPath -LogEntryText "Creating new eDiscovery case for user: $UserEmail" -LogLevel "INFO"
    
    # Get folder IDs from Exchange
    Write-Host "Retrieving folder IDs from Exchange..." -ForegroundColor Yellow
    $folderIDs = Get-FolderIDs -UserEmailAddress $UserEmail
    
    # Set folder query variables using values from Exchange
    $global:ContactsFolderQuery = $folderIDs.ContactsQuery
    $global:TasksFolderQuery = $folderIDs.TasksQuery
    $global:CalFolderQuery = $folderIDs.CalendarQuery
    
    # Check if we have valid folder queries
    if ([string]::IsNullOrEmpty($global:CalFolderQuery) -or [string]::IsNullOrEmpty($global:ContactsFolderQuery) -or [string]::IsNullOrEmpty($global:TasksFolderQuery)) {
        Write-LogEntry -LogName $logPath -LogEntryText "Failed to retrieve all required folder IDs from Exchange" -LogLevel "ERROR"
        Write-Host "Failed to retrieve all required folder IDs from Exchange." -ForegroundColor Red
        Write-Host "Cannot proceed with eDiscovery search without valid folder IDs." -ForegroundColor Red
        return $null
    }
    
    # Create a new eDiscovery case
    $NewCaseDetails = @{
        "displayName" = "$searchName"
        "description" = "$searchName Calendar, Contacts and Task Export"
        "externalId"  = "$searchName"
    }
    
    try {
        # Use Invoke-MgGraphRequest to create case using the configured API version
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases"
        $Case = Invoke-MgGraphRequest -Method "POST" -Uri $apiUrl -Body $NewCaseDetails
        
        $caseId = $Case.id
        Write-LogEntry -LogName $logPath -LogEntryText "Successfully created eDiscovery case with ID: $caseId" -LogLevel "INFO"
        Write-Host "Successfully created eDiscovery case: $searchName (ID: $caseId)" -ForegroundColor Green
        
        # Verify we have a valid case before continuing
        if ($null -eq $Case -or $null -eq $caseId) {
            Write-LogEntry -LogName $logPath -LogEntryText "No valid eDiscovery case ID available. Exiting function." -LogLevel "ERROR"
            Write-Host "No valid eDiscovery case ID available. Exiting function." -ForegroundColor Red
            return $null
        }
        
        # Create the content query
        $contentQuery = "$global:CalFolderQuery OR $global:TasksFolderQuery OR $global:ContactsFolderQuery"
        
        # Modify the query based on ContentType parameter
        if ($ContentTypeFilter -eq "calendar") {
            $contentQuery = "$global:CalFolderQuery"
            Write-LogEntry -LogName $logPath -LogEntryText "Content type set to Calendar only" -LogLevel "INFO"
        }
        elseif ($ContentTypeFilter -eq "contacts") {
            $contentQuery = "$global:ContactsFolderQuery"
            Write-LogEntry -LogName $logPath -LogEntryText "Content type set to Contacts only" -LogLevel "INFO"
        }
        elseif ($ContentTypeFilter -eq "tasks") {
            $contentQuery = "$global:TasksFolderQuery"
            Write-LogEntry -LogName $logPath -LogEntryText "Content type set to Tasks only" -LogLevel "INFO"
        }
        else {
            Write-LogEntry -LogName $logPath -LogEntryText "Content type set to All (Calendar, Contacts, and Tasks)" -LogLevel "INFO"
        }
        
        $NewSearchDetails = @{
            "displayName"      = "$searchName"
            "description"      = "Search for Mailbox Calendar, Contact and Tasks."
            "contentQuery"     = $contentQuery
            "dataSourceScopes" = "allTenantMailboxes"
            "sources"          = @(
                @{
                    "@odata.type" = "#microsoft.graph.security.userSource"
                    "email"       = $UserEmail
                }
            )
        }
        
        # Create the eDiscovery search using MgGraphRequest
        Write-LogEntry -LogName $logPath -LogEntryText "Creating new eDiscovery search: $searchName" -LogLevel "INFO"
        try {
            $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$caseId/searches"
            $Search = Invoke-MgGraphRequest -Method "POST" -Uri $apiUrl -Body $NewSearchDetails
            
            $searchId = $Search.id
            Write-LogEntry -LogName $logPath -LogEntryText "Successfully created search with ID: $searchId" -LogLevel "INFO"
            Write-Host "Successfully created search: $searchName (ID: $searchId)" -ForegroundColor Green
            
            # Estimate search statistics
            Write-LogEntry -LogName $logPath -LogEntryText "Estimating search statistics" -LogLevel "INFO"
            try {
                $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$caseId/searches/$searchId/estimateStatistics"
                Invoke-MgGraphRequest -Method "POST" -Uri $apiUrl -Body @{} | Out-Null
                Write-Host "Successfully started search statistics estimation" -ForegroundColor Green
            }
            catch {
                Write-LogEntry -LogName $logPath -LogEntryText "Warning: Failed to estimate search statistics: $_" -LogLevel "WARNING"
                Write-Host "Warning: Failed to estimate search statistics: $_" -ForegroundColor Yellow
            }
            
            # Return the case and search IDs
            return @{
                CaseId   = $caseId
                SearchId = $searchId
                Name     = $searchName
            }
        }
        catch {
            Write-LogEntry -LogName $logPath -LogEntryText "Failed to create search: $_" -LogLevel "ERROR"
            Write-Host "Failed to create search: $_" -ForegroundColor Red
            return @{
                CaseId   = $caseId
                SearchId = $null
                Name     = $searchName
                Error    = "Failed to create search: $_"
            }
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Failed to create eDiscovery case: $_" -LogLevel "ERROR"
        Write-Host "Failed to create eDiscovery case: $_" -ForegroundColor Red
        return $null
    }
}

# Function to export search results - Updated with proper Review Set workflow
Function Export-CalendarResults {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchName
    )
    
    Write-Host "Starting enhanced export workflow for search: $SearchName" -ForegroundColor Yellow
    Write-LogEntry -LogName $logPath -LogEntryText "Starting enhanced export for search: $SearchName (SearchId: $SearchId, CaseId: $CaseId)" -LogLevel "INFO"
    
    # Use the improved Microsoft Graph eDiscovery workflow with Review Sets
    $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
    Write-Host "Using API version: $apiVersion for eDiscovery operations" -ForegroundColor Cyan
    
    try {
        Write-Host "Step 1: Checking for existing Review Sets..." -ForegroundColor Yellow
        
        # First, check if any Review Sets already exist
        $reviewSetListUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/reviewSets"
        $reviewSetId = $null
        
        try {
            Write-Host "Getting existing review sets from: $reviewSetListUrl" -ForegroundColor Cyan
            $existingReviewSets = Invoke-MgGraphRequest -Method "GET" -Uri $reviewSetListUrl
            
            if ($existingReviewSets.value -and $existingReviewSets.value.Count -gt 0) {
                # Use the first existing Review Set
                $reviewSetId = $existingReviewSets.value[0].id
                Write-Host "✅ Found existing Review Set: $reviewSetId" -ForegroundColor Green
                Write-Host "   Display Name: $($existingReviewSets.value[0].displayName)" -ForegroundColor Cyan
                Write-LogEntry -LogName $logPath -LogEntryText "Using existing Review Set: $reviewSetId" -LogLevel "INFO"
            }
            else {
                Write-Host "No existing Review Sets found, creating new one..." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "⚠️ Failed to check existing Review Sets: $_" -ForegroundColor Yellow
            Write-Host "Will attempt to create a new Review Set..." -ForegroundColor Yellow
        }
        
        # If no existing Review Set found, create a new one
        if ([string]::IsNullOrEmpty($reviewSetId)) {
            $reviewSetParams = @{
                "displayName" = "$SearchName ReviewSet"
                "description" = "Review set for $SearchName export"
            }
            
            try {
                Write-Host "Creating new review set at: $reviewSetListUrl" -ForegroundColor Cyan
                $reviewSet = Invoke-MgGraphRequest -Method "POST" -Uri $reviewSetListUrl -Body $reviewSetParams
                
                $reviewSetId = $reviewSet.id
                Write-Host "✅ Review Set created successfully: $reviewSetId" -ForegroundColor Green
                Write-LogEntry -LogName $logPath -LogEntryText "Review Set created: $reviewSetId" -LogLevel "INFO"
                
            }
            catch {
                Write-Host "❌ Review Set creation failed: $_" -ForegroundColor Red
                Write-LogEntry -LogName $logPath -LogEntryText "Review Set creation failed: $_, trying direct export" -LogLevel "WARNING"
                
                # Fallback to direct export if review set fails
                return Export-SearchDirectly -CaseId $CaseId -SearchId $SearchId -SearchName $SearchName -ApiVersion $apiVersion
            }
        }
        
        # Now proceed with export logic using the Review Set (either existing or newly created)
        Write-Host "Step 2: Adding search content to Review Set..." -ForegroundColor Yellow
        
        # Step 2: Add content from search to review set
        $addToReviewSetParams = @{
            "search"                = @{
                "id" = $SearchId
            }
            # Fixed: additionalDataOptions should be a single value, not an array
            "additionalDataOptions" = "linkedFiles"
        }
        
        $addContentUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/reviewSets/$reviewSetId/addToReviewSet"
        
        try {
            Write-Host "Adding search content to review set..." -ForegroundColor Cyan
            $addToReviewSetOperation = Invoke-MgGraphRequest -Method "POST" -Uri $addContentUrl -Body $addToReviewSetParams
            
            # Debug: Show what the API actually returned
            Write-Host "API Response debug:" -ForegroundColor Cyan
            if ($addToReviewSetOperation) {
                Write-Host "Response type: $($addToReviewSetOperation.GetType().Name)" -ForegroundColor Gray
                Write-Host "Response properties: $($addToReviewSetOperation.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                Write-Host "Full response:" -ForegroundColor Gray
                Write-Host ($addToReviewSetOperation | ConvertTo-Json -Depth 3) -ForegroundColor Gray
            }
            else {
                Write-Host "Response is null or empty" -ForegroundColor Yellow
            }
            
            # Try different ways to get the operation ID
            $addOperationId = $null
            if ($addToReviewSetOperation.id) {
                $addOperationId = $addToReviewSetOperation.id
            }
            elseif ($addToReviewSetOperation.'@odata.context' -and $addToReviewSetOperation.'@odata.context' -match "operations\('([^']+)'\)") {
                $addOperationId = $matches[1]
                Write-Host "Extracted operation ID from @odata.context: $addOperationId" -ForegroundColor Yellow
            }
            elseif ($addToReviewSetOperation.operationId) {
                $addOperationId = $addToReviewSetOperation.operationId
            }
            
            if ([string]::IsNullOrEmpty($addOperationId)) {
                Write-Host "⚠️ No operation ID returned from addToReviewSet API" -ForegroundColor Yellow
                Write-Host "This may indicate the operation was synchronous or API design issue" -ForegroundColor Yellow
                
                # Skip waiting and proceed directly to export
                Write-Host "Proceeding to export step without waiting..." -ForegroundColor Yellow
                
            }
            else {
                Write-Host "✅ Content addition operation started: $addOperationId" -ForegroundColor Green
                
                # Wait for content addition to complete
                Write-Host "Waiting for content to be added to review set..." -ForegroundColor Yellow
                $addCompleted = Wait-ForOperation -CaseId $CaseId -OperationId $addOperationId -OperationType "AddToReviewSet" -MaxAttempts 30
                
                if (-not $addCompleted) {
                    throw "Content addition to review set did not complete in time"
                }
            }
            
        }
        catch {
            Write-Host "❌ Failed to add content to review set: $_" -ForegroundColor Red
            Write-LogEntry -LogName $logPath -LogEntryText "Failed to add content to review set: $_" -LogLevel "ERROR"
            
            # Try direct export as fallback
            return Export-SearchDirectly -CaseId $CaseId -SearchId $SearchId -SearchName $SearchName -ApiVersion $apiVersion
        }
        
        Write-Host "Step 3: Starting export from Review Set..." -ForegroundColor Yellow
        
        # Step 3: Export from Review Set - using correct API parameters
        $exportParams = @{
            "displayName" = "$SearchName Export"
            "description" = "Export from review set for $SearchName"
            # Note: exportOptions and outputName may not be valid parameters
            # Let's try with minimal required parameters first
        }
        
        $exportUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/reviewSets/$reviewSetId/export"
        
        try {
            Write-Host "Initiating export from review set..." -ForegroundColor Cyan
            Write-Host "Export URL: $exportUrl" -ForegroundColor Gray
        
            # Try with minimal parameters first
            $exportResult = Invoke-MgGraphRequest -Method "POST" -Uri $exportUrl -Body $exportParams
        
            $exportId = $exportResult.id
            Write-Host "✅ Export operation started successfully!" -ForegroundColor Green
            Write-Host "   Export ID: $exportId" -ForegroundColor Cyan
            Write-LogEntry -LogName $logPath -LogEntryText "Export started from Review Set: $exportId" -LogLevel "INFO"
        
            # Wait for export to complete
            Write-Host "Waiting for export to complete..." -ForegroundColor Yellow
            $exportCompleted = Wait-ForOperation -CaseId $CaseId -OperationId $exportId -OperationType "Export" -MaxAttempts 60
        
            if ($exportCompleted) {
                Write-Host "✅ Export completed successfully!" -ForegroundColor Green
                return @{
                    Success     = $true
                    Message     = "Export completed successfully from Review Set"
                    ExportId    = $exportId
                    ReviewSetId = $reviewSetId
                }
            }
            else {
                throw "Export operation did not complete in time"
            }
        }
        catch {
            Write-Host "❌ Review Set export failed: $_" -ForegroundColor Red
            Write-Host "⚠️  Microsoft Graph eDiscovery export APIs may not be fully functional yet" -ForegroundColor Yellow
            Write-LogEntry -LogName $logPath -LogEntryText "Review Set export failed: $_, trying direct fallback" -LogLevel "WARNING"
            
            # Try direct export as final fallback
            return Export-SearchDirectly -CaseId $CaseId -SearchId $SearchId -SearchName $SearchName -ApiVersion $apiVersion
        }
        
    }
    catch {
        Write-Host "❌ Critical error in enhanced export workflow: $_" -ForegroundColor Red
        Write-LogEntry -LogName $logPath -LogEntryText "Critical error in enhanced export workflow: $_" -LogLevel "ERROR"
        
        # Fallback to original direct export method
        return Export-SearchDirectly -CaseId $CaseId -SearchId $SearchId -SearchName $SearchName -ApiVersion $apiVersion
    }
}

# Helper function for direct search export (fallback method)
Function Export-SearchDirectly {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        
        [Parameter(Mandatory = $true)]
        [string]$ApiVersion
    )
    
    Write-Host "🔄 Using direct search export as fallback method..." -ForegroundColor Yellow
    Write-LogEntry -LogName $logPath -LogEntryText "Using direct search export fallback for: $SearchName" -LogLevel "INFO"
    
    # Based on the API errors we're seeing, the Microsoft Graph eDiscovery export endpoints
    # are not fully functional yet. Let's provide the best alternative approach.
    
    Write-Host "" -ForegroundColor White
    Write-Host "⚠️  Microsoft Graph API Limitation Detected" -ForegroundColor Yellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow
    Write-Host "The Microsoft Graph eDiscovery export APIs are returning errors." -ForegroundColor White
    Write-Host "This is common with beta endpoints that aren't fully implemented yet." -ForegroundColor White
    Write-Host "" -ForegroundColor White
    
    # First, validate that we have a proper search to work with
    try {
        Write-Host "✅ Validating search configuration..." -ForegroundColor Cyan
        $searchValidationUrl = "https://graph.microsoft.com/$ApiVersion/security/cases/ediscoveryCases/$CaseId/searches/$SearchId"
        $searchDetails = Invoke-MgGraphRequest -Method "GET" -Uri $searchValidationUrl
        
        Write-Host "✅ Search Details Confirmed:" -ForegroundColor Green
        Write-Host "   Search Name: $($searchDetails.displayName)" -ForegroundColor Cyan
        Write-Host "   Search ID: $($searchDetails.id)" -ForegroundColor Cyan
        Write-Host "   Content Query: $($searchDetails.contentQuery)" -ForegroundColor Cyan
        Write-Host "   Created: $($searchDetails.createdDateTime)" -ForegroundColor Cyan
        
        # Verify the search has the expected folder targeting
        if ($searchDetails.contentQuery -match "folderid:") {
            Write-Host "✅ Search contains proper folder targeting!" -ForegroundColor Green
            $folderMatches = [regex]::Matches($searchDetails.contentQuery, "folderid:([A-F0-9]+)")
            Write-Host "   Found $($folderMatches.Count) folder ID(s) in search query" -ForegroundColor Cyan
        }
        
        Write-Host "" -ForegroundColor White
        Write-Host "🎯 RECOMMENDED NEXT STEPS:" -ForegroundColor Green
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
        Write-Host "1. 🌐 Navigate to the Microsoft Purview compliance center" -ForegroundColor White
        Write-Host "   https://compliance.microsoft.com/contentsearch" -ForegroundColor Cyan
        Write-Host "" -ForegroundColor White
        Write-Host "2. 🔍 Find your search: '$($searchDetails.displayName)'" -ForegroundColor White
        Write-Host "   ✅ This search already has the correct folder IDs!" -ForegroundColor Green
        Write-Host "   📧 Calendar: $(if ($searchDetails.contentQuery -match 'folderid:.*010D') {'✅ Included'} else {'❌ Missing'})" -ForegroundColor $(if ($searchDetails.contentQuery -match 'folderid:.*010D') { 'Green' } else { 'Red' })
        Write-Host "   👥 Contacts: $(if ($searchDetails.contentQuery -match 'folderid:.*0112') {'✅ Included'} else {'❌ Missing'})" -ForegroundColor $(if ($searchDetails.contentQuery -match 'folderid:.*0112') { 'Green' } else { 'Red' })
        Write-Host "   📋 Tasks: $(if ($searchDetails.contentQuery -match 'folderid:.*010E') {'✅ Included'} else {'❌ Missing'})" -ForegroundColor $(if ($searchDetails.contentQuery -match 'folderid:.*010E') { 'Green' } else { 'Red' })
        Write-Host "" -ForegroundColor White
        Write-Host "3. 📤 Use the web interface to export this search" -ForegroundColor White
        Write-Host "   - Click 'Export results' from the Actions menu" -ForegroundColor White
        Write-Host "   - Choose your export options (PST, individual messages, etc.)" -ForegroundColor White
        Write-Host "   - The export will include calendar, contacts, and tasks data" -ForegroundColor White
        Write-Host "" -ForegroundColor White
        Write-Host "🎯 KEY POINT: No need to create a new search!" -ForegroundColor Cyan
        Write-Host "   The existing search has perfect folder targeting already." -ForegroundColor Cyan
        Write-Host "" -ForegroundColor White
        
        return @{
            Success       = $true
            Status        = "manual_export_recommended"
            Message       = "Search validated successfully - manual export recommended"
            Method        = "compliance_portal"
            SearchDetails = @{
                Id     = $searchDetails.id
                Name   = $searchDetails.displayName
                Query  = $searchDetails.contentQuery
                CaseId = $CaseId
            }
            Instructions  = "Use Microsoft Purview compliance portal to export the validated search"
        }
        
    }
    catch {
        Write-Host "❌ Could not validate search details: $_" -ForegroundColor Red
        Write-LogEntry -LogName $logPath -LogEntryText "Search validation failed: $_" -LogLevel "ERROR"
        
        Write-Host "" -ForegroundColor White
        Write-Host "Even without full validation, you can still proceed with manual export:" -ForegroundColor Yellow
        Write-Host "1. Open: https://compliance.microsoft.com" -ForegroundColor White
        Write-Host "2. Find your search case and look for search ID: $SearchId" -ForegroundColor White
        Write-Host "3. Export that search using the web interface" -ForegroundColor White
        
        return @{
            Success    = $false
            Status     = "validation_failed_but_search_exists"
            Message    = "Could not validate search but it should still be usable"
            Error      = $_
            CaseId     = $CaseId
            SearchId   = $SearchId
            SearchName = $SearchName
        }
    }
}

# Helper function to wait for operations to complete
Function Wait-ForOperation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $false)]
        [string]$OperationId,
        
        [Parameter(Mandatory = $true)]
        [string]$OperationType,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 30
    )
    
    # If no operation ID provided, assume synchronous operation completed
    if ([string]::IsNullOrEmpty($OperationId)) {
        Write-Host "⚠️ No operation ID provided for $OperationType - assuming synchronous operation completed" -ForegroundColor Yellow
        return $true
    }
    
    Write-Host "Waiting for $OperationType operation to complete..." -ForegroundColor Yellow
    $attempts = 0
    
    do {
        Start-Sleep -Seconds 10
        $attempts++
        
        try {
            $operationUrl = "https://graph.microsoft.com/beta/security/cases/ediscoveryCases/$CaseId/operations/$OperationId"
            $operation = Invoke-MgGraphRequest -Method "GET" -Uri $operationUrl
            
            $status = $operation.status
            $progress = $operation.percentProgress
            
            Write-Host "[$attempts/$MaxAttempts] Status: $status, Progress: $progress%" -ForegroundColor Cyan
            
            if ($status -eq "succeeded") {
                Write-Host "✅ Operation completed successfully!" -ForegroundColor Green
                return $true
            }
            elseif ($status -eq "failed") {
                Write-Host "❌ Operation failed!" -ForegroundColor Red
                return $false
            }
        }
        catch {
            Write-Host "⚠️ Could not check operation status: $_" -ForegroundColor Yellow
        }
        
    } while ($attempts -lt $MaxAttempts)
    
    Write-Host "⏰ Operation did not complete within $MaxAttempts attempts" -ForegroundColor Yellow
    return $false
}

# Function to check export status
Function Get-ExportStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportId
    )
    
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations/$ExportId"
        $exportDetails = Invoke-MgGraphRequest -Method GET -Uri $apiUrl
        
        # Debug output to help troubleshoot - but with limited depth to avoid overwhelming
        Write-Host "Raw export details:" -ForegroundColor Cyan
        Write-Host ($exportDetails | ConvertTo-Json -Depth 3) -ForegroundColor Cyan
        
        $result = @{
            Status        = "unknown"
            CreatedBy     = "Unknown"
            CreatedDate   = "Unknown date"
            DisplayName   = "Export_$ExportId"
            ExportDetails = $exportDetails
            DownloadUrls  = @()
        }
        
        # Get status directly from the status property
        if ($exportDetails.status) {
            $result.Status = $exportDetails.status
        }
        elseif ($exportDetails["status"]) {
            $result.Status = $exportDetails["status"]
        }
        
        # Get the display name directly from the displayName property
        if ($exportDetails.displayName) {
            $result.DisplayName = $exportDetails.displayName
        }
        elseif ($exportDetails["displayName"]) {
            $result.DisplayName = $exportDetails["displayName"]
        }
        
        # Get creator directly
        if ($exportDetails.createdBy.user.displayName) {
            $result.CreatedBy = $exportDetails.createdBy.user.displayName
        }
        elseif ($exportDetails["createdBy"] -and $exportDetails["createdBy"]["user"] -and $exportDetails["createdBy"]["user"]["displayName"]) {
            $result.CreatedBy = $exportDetails["createdBy"]["user"]["displayName"]
        }
        
        # Get created date directly
        if ($exportDetails.createdDateTime) {
            $dateStr = $exportDetails.createdDateTime
            if ($dateStr -is [string] -and $dateStr -match "T") {
                $result.CreatedDate = $dateStr.Substring(0, 19).Replace("T", " ")
            }
            else {
                $result.CreatedDate = $dateStr
            }
        }
        elseif ($exportDetails["createdDateTime"]) {
            $dateStr = $exportDetails["createdDateTime"]
            if ($dateStr -is [string] -and $dateStr -match "T") {
                $result.CreatedDate = $dateStr.Substring(0, 19).Replace("T", " ")
            }
            else {
                $result.CreatedDate = $dateStr
            }
        }
        
        # Extract download URLs from exportFileMetadata
        Write-Host "[DEBUG] Processing export file metadata..." -ForegroundColor Cyan
        
        # Direct access to exportFileMetadata array
        $exportFileMetadata = $exportDetails.exportFileMetadata
        if (-not $exportFileMetadata -and $exportDetails["exportFileMetadata"]) {
            $exportFileMetadata = $exportDetails["exportFileMetadata"]
        }
        
        if ($exportFileMetadata) {
            Write-Host "✅ Found export file metadata" -ForegroundColor Green
            Write-Host "   Type: $($exportFileMetadata.GetType().Name)" -ForegroundColor Gray
            
            # Convert single item to array if needed
            if ($exportFileMetadata -isnot [Array]) {
                $exportFileMetadata = @($exportFileMetadata)
            }
            
            foreach ($fileMetadata in $exportFileMetadata) {
                $fileName = $null
                $downloadUrl = $null
                $fileSize = 0
                
                # Try both object and hashtable access methods
                if ($fileMetadata.fileName) {
                    $fileName = $fileMetadata.fileName
                    $downloadUrl = $fileMetadata.downloadUrl
                    $fileSize = $fileMetadata.size
                }
                elseif ($fileMetadata["fileName"]) {
                    $fileName = $fileMetadata["fileName"]
                    $downloadUrl = $fileMetadata["downloadUrl"]
                    $fileSize = $fileMetadata["size"]
                }
                
                if ($fileName -and $downloadUrl) {
                    Write-Host "   Found file: $fileName" -ForegroundColor Green
                    $result.DownloadUrls += @{
                        FileName    = $fileName
                        DownloadUrl = $downloadUrl
                        Size        = $fileSize
                    }
                }
            }
            Write-Host "[DEBUG] Created $($result.DownloadUrls.Count) download URL entries" -ForegroundColor Cyan
        }
        else {
            Write-Host "[DEBUG] No export file metadata found" -ForegroundColor Yellow
            Write-Host "[DEBUG] Available properties:" -ForegroundColor Gray
            
            # Handle both object and hashtable formats
            if ($exportDetails.PSObject.Properties) {
                Write-Host "   $($exportDetails.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
            }
            elseif ($exportDetails.Keys) {
                Write-Host "   $($exportDetails.Keys -join ', ')" -ForegroundColor Gray
            }
        }
        
        Write-Host "Final Results:" -ForegroundColor Cyan
        Write-Host "  Status: $($result.Status)" -ForegroundColor Gray
        Write-Host "  Created By: $($result.CreatedBy)" -ForegroundColor Gray
        Write-Host "  Created Date: $($result.CreatedDate)" -ForegroundColor Gray
        Write-Host "  Download URLs found: $($result.DownloadUrls.Count)" -ForegroundColor Gray
        
        return $result
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error checking export status: $_" -LogLevel "ERROR"
        Write-Host "Error checking export status: $_" -ForegroundColor Red
        return @{
            Status      = "error"
            Error       = $_
            DisplayName = "Export_$ExportId"
        }
    }
}

# Helper function to check if content is HTML (indicating auth redirect)
function Test-IsHtmlContent {
    param ([string]$FilePath)
        
    try {
        # Only check files smaller than 1MB - larger files are unlikely to be HTML login pages
        $fileInfo = Get-Item -Path $FilePath
        if ($fileInfo.Length -gt 1MB) {
            return $false
        }
            
        # Read the first bytes to determine content type
        $fileStream = $null
        try {
            $fileStream = [System.IO.File]::OpenRead($FilePath)
            $buffer = New-Object byte[] 1024
            $bytesRead = $fileStream.Read($buffer, 0, 1024)
            
            # Convert bytes to string for text analysis
            $fileText = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
                
            # Check for HTML signatures
            if ($fileText -match '<!DOCTYPE html' -or 
                $fileText -match '<html' -or 
                $fileText -match '<HTML' -or
                $fileText -match '<title>' -or
                $fileText -match '<body' -or
                $fileText -match 'login' -or
                $fileText -match 'sign in') {
                    
                Write-Host "WARNING: Downloaded content appears to be HTML instead of binary file!" -ForegroundColor Red
                Write-Host "HTML snippet: $($fileText.Substring(0, [Math]::Min(200, $fileText.Length)))..." -ForegroundColor Red
                return $true
            }
                
            # If we get here, the content is likely not HTML
            return $false
        }
        finally {
            if ($fileStream) { $fileStream.Close() }
        }
    }
    catch {
        Write-Host "Error checking content type: $_" -ForegroundColor Yellow
        return $false
    }
}

# Function to download files from M365 Compliance Center with advanced authentication
Function Save-M365ComplianceFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DownloadUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 2,
        
        [Parameter(Mandatory = $false)]
        [string]$AuthToken = $null
    )
    
    # Debug: Log guard variable values
    Write-Host "[DEBUG] IsScriptInitializing: $($global:IsScriptInitializing), Operation: $($global:Operation), DownloadUrl: $DownloadUrl, OutputFilePath: $OutputFilePath" -ForegroundColor Cyan

    # Skip downloads only during actual script initialization, not when user explicitly requests downloads
    if ($global:IsScriptInitializing -or
        [string]::IsNullOrEmpty($DownloadUrl) -or [string]::IsNullOrEmpty($OutputFilePath)) {
        Write-Host "[GUARD] Skipping download - missing required parameters or during initialization" -ForegroundColor Yellow
        Write-Host "IsScriptInitializing: $($global:IsScriptInitializing)" -ForegroundColor Gray 
        Write-Host "DownloadUrl: $(if([string]::IsNullOrEmpty($DownloadUrl)){'<empty>'}else{'<set>'})" -ForegroundColor Gray
        Write-Host "OutputFilePath: $(if([string]::IsNullOrEmpty($OutputFilePath)){'<empty>'}else{'<set>'})" -ForegroundColor Gray
        return $false
    }
    
    Write-Host "Starting specialized M365 Compliance download with auth-aware flow..." -ForegroundColor Yellow
    
    # Make sure the directory exists
    $outputDir = Split-Path -Path $OutputFilePath -Parent
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Delete the file if it already exists to start fresh
    if (Test-Path -Path $OutputFilePath) {
        Remove-Item -Path $OutputFilePath -Force
    }
    
    # TRY METHOD 1: Microsoft's Official App-Only Download Method (per TechCommunity article)
    Write-Host " Using the MSAL with app-only download method..." -ForegroundColor Cyan
    try {
        # Check if MSAL.PS module is available for MicrosoftPurviewEDiscovery API authentication
        $msalModule = Get-Module -Name MSAL.PS -ListAvailable
        if (-not $msalModule) {
            Write-Host "Installing MSAL.PS module for official Microsoft download method..." -ForegroundColor Yellow
            try {
                Install-Module MSAL.PS -Scope CurrentUser -Force -ErrorAction Stop
                Write-Host "✅ MSAL.PS module installed successfully" -ForegroundColor Green
            }
            catch {
                Write-Host "❌ Failed to install MSAL.PS module: $_" -ForegroundColor Red
                throw "MSAL.PS module installation required for app-only downloads"
            }
        }
        
        Write-Host "Acquiring MicrosoftPurviewEDiscovery API token using official method..." -ForegroundColor Yellow
        
        # Get MicrosoftPurviewEDiscovery API token using MSAL.PS (per Microsoft's documentation)
        $exportToken = $null
        if ($AuthMode -eq "Cert") {
            # Certificate authentication
            $ClientCert = Get-ChildItem "$certlocation$CertificateThumbprint" -ErrorAction SilentlyContinue
            if (-not $ClientCert) {
                throw "Certificate not found with thumbprint: $CertificateThumbprint"
            }
            
            $connectionDetails = @{
                'TenantId'          = $TenantId
                'ClientId'          = $AppID
                'ClientCertificate' = $ClientCert
                'Scope'             = "b26e684c-5068-4120-a679-64a5d2c909d9/.default"  # MicrosoftPurviewEDiscovery API scope
            }
            
            $exportToken = Get-MsalToken @connectionDetails -ErrorAction Stop
        }
        else {
            # Client Secret authentication
            $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            
            $exportToken = Get-MsalToken -ClientId $AppID -Scopes "b26e684c-5068-4120-a679-64a5d2c909d9/.default" -TenantId $TenantId -RedirectUri "http://localhost" -ClientSecret $secureClientSecret -ErrorAction Stop
        }
        
        if ($exportToken -and $exportToken.AccessToken) {
            Write-Host "✅ Successfully acquired MicrosoftPurviewEDiscovery API token" -ForegroundColor Green
            
            # Use Microsoft's official download method
            Write-Host "Downloading file using Microsoft's official method..." -ForegroundColor Cyan
            Write-Host "  Source: $DownloadUrl" -ForegroundColor Gray
            Write-Host "  Target: $OutputFilePath" -ForegroundColor Gray
            
            # Official headers as documented by Microsoft
            $headers = @{
                "Authorization"       = "Bearer $($exportToken.AccessToken)"
                "X-AllowWithAADToken" = "true"
            }
            
            # Download using Microsoft's recommended approach
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $OutputFilePath -Headers $headers -ErrorAction Stop
            
            # Verify download
            if (Test-Path -Path $OutputFilePath) {
                $fileSize = (Get-Item -Path $OutputFilePath).Length
                if ($fileSize -gt 0) {
                    Write-Host "✅ Microsoft official method download SUCCESS!" -ForegroundColor Green
                    Write-Host "   File: $OutputFilePath" -ForegroundColor Cyan
                    Write-Host "   Size: $([math]::Round($fileSize / 1MB, 2)) MB" -ForegroundColor Cyan
                    
                    # Verify it's not an HTML error page
                    if (-not (Test-IsHtmlContent -FilePath $OutputFilePath)) {
                        return @{ 
                            Success  = $true 
                            FilePath = $OutputFilePath 
                            Size     = $fileSize
                            Method   = "Microsoft Official App-Only (TechCommunity)"
                        }
                    }
                    else {
                        Write-Host "❌ Downloaded content appears to be HTML (authentication error)" -ForegroundColor Red
                        Remove-Item -Path $OutputFilePath -Force -ErrorAction SilentlyContinue
                        throw "Downloaded HTML instead of expected file content"
                    }
                }
                else {
                    throw "Downloaded file is empty"
                }
            }
            else {
                throw "Download file was not created"
            }
        }
        else {
            throw "Failed to acquire MicrosoftPurviewEDiscovery API token"
        }
    }
    catch {
        Write-Host "❌ Microsoft official method failed: $_" -ForegroundColor Red
        Write-LogEntry -LogName $logPath -LogEntryText "Microsoft official download failed: $_" -LogLevel "WARNING"
        
        Write-Host "❌ Microsoft's official app-only method is the only supported download method." -ForegroundColor Red
        Write-Host "   Please ensure MSAL.PS module is installed and configured properly." -ForegroundColor Yellow
        Write-Host "   Install-Module MSAL.PS -Scope CurrentUser" -ForegroundColor Cyan
        
        return @{ 
            Success = $false 
            Error   = "Microsoft official app-only method failed and no fallback methods available"
            Method  = "None - App-only authentication required"
        }
    }
}

# Function to download files from eDiscovery proxy service specifically
# Based on the article: https://michev.info/blog/post/5806/using-the-graph-api-to-export-ediscovery-premium-datasets
Function Get-eDiscoveryProxyFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DownloadUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$AuthToken = $null
    )
    
    Write-Host "Starting specialized eDiscovery Proxy download..." -ForegroundColor Yellow
    
    # Make sure the directory exists
    $outputDir = Split-Path -Path $OutputFilePath -Parent
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Delete the file if it already exists to start fresh
    if (Test-Path -Path $OutputFilePath) {
        Remove-Item -Path $OutputFilePath -Force
    }
    
    # If no token provided, try to get one from the Graph context
    if (-not $AuthToken) {
        try {
            $graphContext = Get-MgContext
            if ($graphContext -and $graphContext.AccessToken) {
                $AuthToken = $graphContext.AccessToken
                Write-Host "Retrieved authentication token from Graph context" -ForegroundColor Green
            }
            else {
                Write-Host "No Graph context found. Please connect to Microsoft Graph first." -ForegroundColor Yellow
                return $false
            }
        }
        catch {
            Write-Host "Error retrieving Graph token: $_" -ForegroundColor Red
            return $false
        }
    }
    
    # Create headers with the authentication token and special header for eDiscovery proxy service
    $headers = @{
        'Content-Type'        = 'application/json'
        'Authorization'       = "Bearer $AuthToken"
        'X-AllowWithAADToken' = "true"
        'Accept'              = 'application/octet-stream'
        'User-Agent'          = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    }
    
    # Disable progress bar for better performance
    $ProgressPreference = 'SilentlyContinue'
    
    $downloadSuccess = $false
    
    try {
        Write-Host "Downloading file with eDiscovery proxy specialized approach..." -ForegroundColor Yellow
        
        # First attempt: Use Invoke-WebRequest with X-AllowWithAADToken
        try {
            Invoke-WebRequest -Uri $DownloadUrl -Headers $headers -OutFile $OutputFilePath -Method Get -UseBasicParsing
            
            if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                $fileSize = (Get-Item -Path $OutputFilePath).Length
                
                # Check if the file is HTML or actual binary content
                $fileStream = [System.IO.File]::OpenRead($OutputFilePath)
                $buffer = New-Object byte[] 1024
                $bytesRead = $fileStream.Read($buffer, 0, 1024)
                $fileStream.Close()
                
                $fileText = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
                
                if ($fileText -match '<!DOCTYPE html' -or $fileText -match '<html') {
                    Write-Host "Download returned HTML content instead of binary file. Trying second approach..." -ForegroundColor Yellow
                    
                    # If the file is HTML, try second approach (use WebClient)
                    try {
                        $headers.Remove('X-AllowWithAADToken')
                        $headers['Accept'] = 'application/octet-stream, application/json, */*'
                        
                        # Create a WebClient for binary download
                        $webClient = New-Object System.Net.WebClient
                        
                        # Add headers to WebClient
                        foreach ($key in $headers.Keys) {
                            $webClient.Headers.Add($key, $headers[$key])
                        }
                        
                        # Download directly to file
                        $webClient.DownloadFile($DownloadUrl, $OutputFilePath)
                        
                        if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                            $fileSize = (Get-Item -Path $OutputFilePath).Length
                            Write-Host "File downloaded successfully with WebClient - size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                            $downloadSuccess = $true
                        }
                        else {
                            throw "WebClient download produced empty file"
                        }
                    }
                    catch {
                        Write-Host "WebClient approach failed: $_" -ForegroundColor Yellow
                        
                        # Try curl as last resort
                        try {
                            Write-Host "Trying curl approach..." -ForegroundColor Yellow
                            
                            # Format headers for curl
                            $curlHeaders = @()
                            foreach ($key in $headers.Keys) {
                                $curlHeaders += "-H `"$key`: $($headers[$key])`""
                            }
                            
                            $curlCommand = "curl.exe -L -o `"$OutputFilePath`" $($curlHeaders -join ' ') `"$DownloadUrl`""
                            Write-Host "Executing: $curlCommand" -ForegroundColor Yellow
                            
                            Invoke-Expression $curlCommand
                            
                            if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                                $fileSize = (Get-Item -Path $OutputFilePath).Length
                                Write-Host "File downloaded successfully with curl - size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                                $downloadSuccess = $true
                            }
                            else {
                                throw "Curl download produced empty file"
                            }
                        }
                        catch {
                            Write-Host "Curl approach failed: $_" -ForegroundColor Red
                        }
                    }
                }
                else {
                    # Not HTML, so it's likely the correct binary content
                    Write-Host "File downloaded successfully with size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                    $downloadSuccess = $true
                }
            }
            else {
                throw "Download produced empty file"
            }
        }
        catch {
            Write-Host "First download attempt failed: $_" -ForegroundColor Yellow
            
            # Try alternative approach without X-AllowWithAADToken
            try {
                $headers.Remove('X-AllowWithAADToken')
                
                Invoke-WebRequest -Uri $DownloadUrl -Headers $headers -OutFile $OutputFilePath -Method Get -UseBasicParsing
                
                if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                    $fileSize = (Get-Item -Path $OutputFilePath).Length
                    Write-Host "File downloaded successfully (second attempt) with size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                    $downloadSuccess = $true
                }
                else {
                    throw "Second download attempt produced empty file"
                }
            }
            catch {
                Write-Host "Second download attempt failed: $_" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "Error in eDiscovery proxy specialized download: $_" -ForegroundColor Red
        $downloadSuccess = $false
    }
    finally {
        $ProgressPreference = 'Continue'  # Reset progress preference
    }
    
    return $downloadSuccess
}

# Function to handle Microsoft 365 Purview/Compliance URLs specifically
Function ConvertTo-M365PurviewDownload {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DownloadUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath
    )
    
    # Skip downloads only during actual script initialization, not when user explicitly requests downloads
    if ($global:IsScriptInitializing -or
        [string]::IsNullOrEmpty($DownloadUrl) -or [string]::IsNullOrEmpty($OutputFilePath)) {
        Write-Host "[GUARD] Skipping purview download - missing required parameters or during initialization" -ForegroundColor Yellow
        Write-Host "IsScriptInitializing: $($global:IsScriptInitializing)" -ForegroundColor Gray 
        Write-Host "DownloadUrl: $(if([string]::IsNullOrEmpty($DownloadUrl)){'<empty>'}else{'<set>'})" -ForegroundColor Gray
        Write-Host "OutputFilePath: $(if([string]::IsNullOrEmpty($OutputFilePath)){'<empty>'}else{'<set>'})" -ForegroundColor Gray
        return $false
    }
    
    $downloadSuccess = $false
    $retryCount = 0
    $MaxRetriesLocal = 3
    $RetryDelaySecondsLocal = 2
    
    while (-not $downloadSuccess -and $retryCount -lt $MaxRetriesLocal) {
        try {
            if ($retryCount -gt 0) {
                Write-Host "Retry attempt $retryCount of $MaxRetriesLocal..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelaySecondsLocal
            }
            
            Write-Host "Creating web session with authentication support..." -ForegroundColor Yellow
            
            # Create a session to maintain cookies across requests
            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            
            # Add cookies container to store authentication cookies
            $session.Cookies = New-Object System.Net.CookieContainer
            
            # Set modern user agent to avoid authentication issues
            $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62"
            
            # Set headers appropriate for binary download from Microsoft services
            $headers = @{
                "Accept"          = "application/octet-stream, application/json, */*"
                "User-Agent"      = $userAgent
                "Sec-Fetch-Site"  = "same-origin"
                "Sec-Fetch-Mode"  = "navigate"
                "Sec-Fetch-Dest"  = "document"
                "Accept-Encoding" = "gzip, deflate, br"
                "Accept-Language" = "en-US,en;q=0.9"
            }
            
            # If we have an auth token, use it (for Graph API authenticated URLs)
            if ($AuthToken) {
                Write-Host "Using provided authentication token for download..." -ForegroundColor Green
                $headers["Authorization"] = "Bearer $AuthToken"
            }
            else {
                # Try to get auth token from Graph context if available
                try {
                    $graphContext = Get-MgContext
                    if ($graphContext -and $graphContext.TokenCache -and $graphContext.TokenCache.AccessToken) {
                        $AuthToken = $graphContext.TokenCache.AccessToken
                        Write-Host "Retrieved auth token from Graph context" -ForegroundColor Green
                        $headers["Authorization"] = "Bearer $AuthToken"
                    }
                }
                catch {
                    Write-Host "Could not retrieve Graph auth token: $_" -ForegroundColor Yellow
                    Write-Host "Continuing with cookie-based authentication..." -ForegroundColor Yellow
                }
            }
            
            # Disable progress bar for better performance
            $ProgressPreference = 'SilentlyContinue'
            
            # First make a request to the URL to follow redirects and establish authentication
            Write-Host "Step 1: Initial request to follow redirects and establish auth..." -ForegroundColor Yellow
            try {
                $initialResponse = Invoke-WebRequest -Uri $DownloadUrl -WebSession $session -Headers $headers -Method Get -UseBasicParsing -MaximumRedirection 10
                Write-Host "Initial request successful with status: $($initialResponse.StatusCode)" -ForegroundColor Green
                
                # Check if we received HTML content instead of expected binary
                if ($initialResponse.Content.Length -lt 50000 -and ($initialResponse.Content -match '<!DOCTYPE html' -or $initialResponse.Content -match '<html')) {
                    Write-Host "WARNING: Initial response appears to be HTML instead of binary data!" -ForegroundColor Yellow
                    
                    # Check if it's a login page
                    if ($initialResponse.Content -match 'login' -or $initialResponse.Content -match 'sign in' -or $initialResponse.Content -match 'authentication') {
                        Write-Host "Detected login/authentication page! Session appears to be unauthenticated." -ForegroundColor Red
                        # We'll still continue to try other methods
                    }
                }
            }
            catch [System.Net.WebException] {
                $respStream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($respStream)
                $respBody = $reader.ReadToEnd()
                Write-Host "Initial request returned error: $respBody" -ForegroundColor Yellow
                
                # Still continue as we may have established cookies needed for download
                Write-Host "Continuing despite error as authentication may still be established..." -ForegroundColor Yellow
            }
            
            # Now try to download the file directly
            Write-Host "Step 2: Downloading file using established session..." -ForegroundColor Yellow
            try {
                # Modify headers for direct file download
                $headers["Accept"] = "application/octet-stream"
                
                # Stream the response directly to a file
                Invoke-WebRequest -Uri $DownloadUrl -WebSession $session -Headers $headers -OutFile $OutputFilePath -Method Get -UseBasicParsing -MaximumRedirection 10
                
                # Check if file was created with content
                if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                    $fileSize = (Get-Item -Path $OutputFilePath).Length
                    
                    # Check if the file contains HTML (login page) instead of binary data
                    if (Test-IsHtmlContent -FilePath $OutputFilePath) {
                        Write-Host "Downloaded file appears to be HTML (possibly login page) instead of expected binary file!" -ForegroundColor Red
                        Write-Host "Will try alternative download methods..." -ForegroundColor Yellow
                        throw "Downloaded content is HTML instead of expected binary file"
                    }
                    
                    Write-Host "File downloaded successfully with size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                    $downloadSuccess = $true
                }
                else {
                    Write-Host "File appears empty or wasn't created properly" -ForegroundColor Yellow
                    throw "Downloaded file is empty or wasn't created"
                }
            }
            catch {
                Write-Host "Direct download failed: $_" -ForegroundColor Yellow
                Write-Host "Trying alternative download method..." -ForegroundColor Yellow
                
                # Try with System.Net.WebClient for more direct binary handling
                try {
                    $webClient = New-Object System.Net.WebClient
                    
                    # Transfer cookies from the session to WebClient
                    $webClient.Headers.Add([System.Net.HttpRequestHeader]::Cookie, $session.Cookies.GetCookieHeader([System.Uri]$DownloadUrl))
                    
                    # Add the headers
                    foreach ($key in $headers.Keys) {
                        $webClient.Headers.Add($key, $headers[$key])
                    }
                    
                    # Download the file directly to disk
                    $webClient.DownloadFile($DownloadUrl, $OutputFilePath)
                    
                    # Check if file was created with content
                    if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                        $fileSize = (Get-Item -Path $OutputFilePath).Length
                        
                        # Check if the file contains HTML (login page) instead of binary data
                        if (Test-IsHtmlContent -FilePath $OutputFilePath) {
                            Write-Host "Downloaded file appears to be HTML (possibly login page) instead of expected binary file!" -ForegroundColor Red
                            Write-Host "Will try another download method..." -ForegroundColor Yellow
                            throw "Downloaded content is HTML instead of expected binary file"
                        }
                        
                        Write-Host "File downloaded successfully with WebClient - size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                        $downloadSuccess = $true
                    }
                    else {
                        Write-Host "File appears empty or wasn't created properly with WebClient" -ForegroundColor Yellow
                        throw "Downloaded file is empty or wasn't created"
                    }
                }
                catch {
                    Write-Host "WebClient download failed: $_" -ForegroundColor Yellow
                    Write-Host "Trying direct command-line approach (curl)..." -ForegroundColor Yellow
                    
                    # Try curl as a last resort
                    try {
                        # Format cookies for curl
                        $cookieString = ""
                        foreach ($cookie in $session.Cookies.GetCookies([System.Uri]$DownloadUrl)) {
                            $cookieString += "$($cookie.Name)=$($cookie.Value); "
                        }
                        
                        # Use curl.exe with the session cookies
                        $curlParams = @(
                            '-L',                      # Follow redirects
                            '-o', "`"$OutputFilePath`"", # Output file
                            '-H', "`"User-Agent: $userAgent`"",
                            '-H', "`"Accept: application/octet-stream`""
                        )
                        
                        if ($cookieString) {
                            $curlParams += @('-H', "`"Cookie: $cookieString`"")
                        }
                        
                        # Add auth token if available
                        if ($headers.ContainsKey("Authorization")) {
                            $curlParams += @('-H', "`"Authorization: $($headers["Authorization"])`"")
                        }
                        
                        $curlParams += "`"$DownloadUrl`""
                        
                        # Execute curl command
                        $curlCommand = "curl.exe $($curlParams -join ' ')"
                        Write-Host "Executing: $curlCommand" -ForegroundColor Yellow
                        Invoke-Expression $curlCommand
                        
                        # Check if file was created with content
                        if ((Test-Path -Path $OutputFilePath) -and ((Get-Item -Path $OutputFilePath).Length -gt 0)) {
                            $fileSize = (Get-Item -Path $OutputFilePath).Length
                            
                            # Check if the file contains HTML (login page) instead of binary data
                            if (Test-IsHtmlContent -FilePath $OutputFilePath) {
                                Write-Host "Downloaded file appears to be HTML (possibly login page) instead of expected binary file!" -ForegroundColor Red
                                throw "Downloaded content is HTML instead of expected binary file"
                            }
                            
                            Write-Host "File downloaded successfully with curl - size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green
                            $downloadSuccess = $true
                        }
                        else {
                            Write-Host "File appears empty or wasn't created properly with curl" -ForegroundColor Yellow
                            throw "Downloaded file is empty or wasn't created"
                        }
                    }
                    catch {
                        Write-Host "curl download failed: $_" -ForegroundColor Red
                        Write-Host "All download methods failed for this attempt" -ForegroundColor Red
                    }
                }
            }
        }
        catch {
            Write-Host "Error during download attempt $retryCount`: $($_)" -ForegroundColor Red
        }
        finally {
            $ProgressPreference = 'Continue'  # Reset preference
        }
        
        $retryCount++
    }
    
    if ($downloadSuccess) {
        Write-Host "Download completed successfully after $retryCount attempt(s)" -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "Download failed after $MaxRetriesLocal attempts" -ForegroundColor Red
        return $false
    }
}

# All duplicate function code has been removed

Function Test-ValidZipFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$RetryOnFailure,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelay = 1 # seconds
    )
    
    # If RetryOnFailure is enabled, try multiple times with a delay
    if ($RetryOnFailure) {
        $retryCount = 0
        $success = $false
        
        while (-not $success -and $retryCount -lt $MaxRetries) {
            if ($retryCount -gt 0) {
                Write-Host "Retry attempt $retryCount of $MaxRetries after $RetryDelay second(s)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelay
            }
            
            $success = Test-ZipFileInternal -FilePath $FilePath
            $retryCount++
            
            if ($success) {
                return $true
            }
        }
        
        if (-not $success) {
            Write-Host "ZIP validation failed after $MaxRetriesLocal attempts." -ForegroundColor Red
            return $false
        }
    }
    else {
        # Just try once
        return Test-ZipFileInternal -FilePath $FilePath
    }
}

# Internal function to validate ZIP files
function Test-ZipFileInternal {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        # Check if the file exists
        if (-not (Test-Path -Path $FilePath)) {
            Write-Host "File does not exist: $FilePath" -ForegroundColor Red
            return $false
        }
        
        # Check file size
        $fileInfo = Get-Item -Path $FilePath
        if ($fileInfo.Length -eq 0) {
            Write-Host "File is empty (0 bytes): $FilePath" -ForegroundColor Red
            return $false
        }
        
        # Check if this is a special M365 Report file (these aren't really ZIP files despite the extension)
        if ($FilePath -like "*Reports-*" -and $fileInfo.Length -gt 5KB) {
            Write-Host "Microsoft 365 Report file detected. These files use .zip extension but aren't standard ZIP files." -ForegroundColor Yellow
            Write-Host "File appears valid with size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
            Write-Host "Report files should be handled as HTML/XML exports rather than ZIP archives." -ForegroundColor Yellow
            return $true
        }
        
        # Check if this is an M365 compliance export - they sometimes use special formats
        $isM365Export = $FilePath -like "*PSTs.*" -or $FilePath -match "Exportdata|Admin_Mailbox_Export"
        
        # For M365 exports, sometimes we can just trust the file if it has a reasonable size
        if ($isM365Export -and $fileInfo.Length -gt 100KB) {
            # Special handling for Office 365 compliance exports which may have non-standard ZIP formats
            Write-Host "Microsoft 365 compliance export detected. Performing special validation..." -ForegroundColor Yellow
            
            # Special validation for M365 exports - try to read first few bytes to confirm it's a ZIP
            try {
                $fileStream = [System.IO.File]::OpenRead($FilePath)
                $buffer = New-Object byte[] 4
                $bytesRead = $fileStream.Read($buffer, 0, 4)
                $fileStream.Close()
                
                # Check for ZIP file signature (PK\003\004)
                if ($bytesRead -eq 4 -and $buffer[0] -eq 80 -and $buffer[1] -eq 75 -and $buffer[2] -eq 3 -and $buffer[3] -eq 4) {
                    Write-Host "ZIP file signature detected. File appears to be a valid ZIP archive." -ForegroundColor Green
                    return $true
                }
                else {
                    Write-Host "File does not have a valid ZIP signature." -ForegroundColor Yellow
                    # Continue with other validation methods
                }
            }
            catch {
                Write-Host "Error checking file signature: $_" -ForegroundColor Yellow
                # Continue with other validation methods
            }
        }
        
        # First try using .NET's ZipFile class
        try {
            # Add required .NET assembly if not already loaded
            if (-not ([System.Management.Automation.PSTypeName]'System.IO.Compression.ZipFile').Type) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
            }
            
            $zipFile = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
            # Try to access entries to verify the ZIP structure
            $entriesCount = $zipFile.Entries.Count
            $zipFile.Dispose()
            
            Write-Host "ZIP file validation successful: $FilePath (Contains $entriesCount entries)" -ForegroundColor Green
            return $true
        } 
        catch [System.IO.InvalidDataException] {
            Write-Host "First validation method failed (InvalidDataException). Trying alternate method..." -ForegroundColor Yellow
            # Try second method - some ZIP files may not be compatible with .NET's ZipFile
            try {
                # Try using System.IO.Packaging which can handle some ZIP formats better
                Add-Type -AssemblyName WindowsBase
                $zip = [System.IO.Packaging.ZipPackage]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                $zip.Close()
                
                Write-Host "ZIP file validation successful using alternate method: $FilePath" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "Second validation method failed. Trying third method for M365 exports..." -ForegroundColor Yellow
                
                # Third method - specifically for Microsoft 365 exports
                # Sometimes Microsoft 365 exports can only be verified with binary analysis
                try {
                    # Special case for Report files (they're not real ZIP files)
                    if ($FilePath -like "*Reports-*") {
                        # Microsoft 365 report files use .zip extension but aren't actual ZIP archives
                        # They're often HTML/XML exports with size > 5KB
                        if ($fileInfo.Length -gt 5KB) {
                            Write-Host "Microsoft 365 Report file detected with adequate size. Accepting as valid." -ForegroundColor Green
                            Write-Host "File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                            return $true
                        }
                    }
                    
                    # For PST files in ZIP format
                    if ($FilePath -like "*PSTs.*") {
                        # Check if file size seems reasonable for a PST file (at least 100KB)
                        if ($fileInfo.Length -gt 100KB) {
                            Write-Host "M365 PST export with reasonable size detected. Accepting as valid: $FilePath" -ForegroundColor Green
                            
                            # Try to rename with .pst extension if needed
                            if (-not $FilePath.EndsWith(".pst")) {
                                $pstFilePath = $FilePath -replace "\.zip$", ".pst"
                                try {
                                    Rename-Item -Path $FilePath -NewName $pstFilePath -Force
                                    Write-Host "Renamed to PST file: $pstFilePath" -ForegroundColor Green
                                }
                                catch {
                                    Write-Host "Unable to rename to PST file: $_" -ForegroundColor Yellow
                                }
                            }
                            
                            return $true
                        }
                    }
                    
                    # For other M365 exports
                    if ($isM365Export) {
                        # Check if file size seems reasonable for an export (at least 50KB)
                        if ($fileInfo.Length -gt 50KB) {
                            Write-Host "M365 export with reasonable size detected. Accepting as valid: $FilePath" -ForegroundColor Green
                            Write-Host "File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                            return $true
                        }
                    }
                }
                catch {
                    Write-Host "Third validation method failed: $_" -ForegroundColor Yellow
                }
                
                Write-Host "Invalid ZIP file (all methods failed): $FilePath" -ForegroundColor Red
                Write-Host "Error details: $_" -ForegroundColor Red
                return $false
            }
        }
        catch {
            Write-Host "Invalid ZIP file: $FilePath" -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error validating ZIP file: $_" -ForegroundColor Red
        return $false
    }
}

# Function to download export results
Function Save-ExportResults {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = $DownloadPath
    )
    
    Write-Host "Checking export status and preparing to download..." -ForegroundColor Yellow
    
    try {
        # Get export status
        $exportStatus = Get-ExportStatus -CaseId $CaseId -ExportId $ExportId
        
        # Use the display name from the status if available
        $exportDisplayName = if ($exportStatus.PSObject.Properties.Name -contains "DisplayName" -and -not [string]::IsNullOrEmpty($exportStatus.DisplayName)) {
            $exportStatus.DisplayName
        }
        else {
            $SearchName
        }
        
        # Sanitize the display name for use as a filename
        $invalidChars = [IO.Path]::GetInvalidFileNameChars()
        $sanitizedName = $exportDisplayName
        foreach ($char in $invalidChars) {
            $sanitizedName = $sanitizedName.Replace($char, '_')
        }
        
        # More flexible status checking - if the status is "unknown" but we have the export details,
        # we'll try to download anyway and let the actual download attempt determine success
        if ($exportStatus.Status -eq "succeeded" -or 
            $exportStatus.Status -eq "partiallySucceeded" -or 
            $exportStatus.Status -eq "completed" -or 
            $exportStatus.Status -eq "done" -or
            $exportStatus.Status -eq "unknown") {
            
            # If status is unknown, log this but continue
            if ($exportStatus.Status -eq "unknown") {
                Write-LogEntry -LogName $logPath -LogEntryText "Export status is unknown, but will attempt download anyway" -LogLevel "WARNING"
                Write-Host "Export status is unknown, but will attempt download anyway" -ForegroundColor Yellow
            }
            # Create folder if it doesn't exist
            if (-not (Test-Path -Path $OutputPath)) {
                Write-LogEntry -LogName $logPath -LogEntryText "Creating download folder: $OutputPath" -LogLevel "INFO"
                New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            }
            
            # Default to .zip extension for multi-file exports
            $outputFilePath = "$OutputPath\$sanitizedName.zip"
            
            # Check if we have download URLs from the export status
            $hasDownloadUrl = $false
            $downloadUrl = $null
            $downloadFiles = @()
            
            Write-Host "[DEBUG] Checking exportStatus for download URLs..." -ForegroundColor Cyan
            
            # For hashtables, check keys directly instead of PSObject.Properties
            if ($exportStatus -is [hashtable]) {
                Write-Host "[DEBUG] exportStatus keys: $($exportStatus.Keys -join ', ')" -ForegroundColor Gray
                Write-Host "[DEBUG] Has DownloadUrls key: $($exportStatus.ContainsKey('DownloadUrls'))" -ForegroundColor Gray
                Write-Host "[DEBUG] Has exportFileMetadata key: $($exportStatus.ContainsKey('exportFileMetadata'))" -ForegroundColor Gray
                
                if ($exportStatus.ContainsKey('DownloadUrls') -and $exportStatus.DownloadUrls) {
                    Write-Host "[DEBUG] DownloadUrls count: $($exportStatus.DownloadUrls.Count)" -ForegroundColor Gray
                }
                
                if ($exportStatus.ContainsKey('exportFileMetadata') -and $exportStatus.exportFileMetadata) {
                    Write-Host "[DEBUG] exportFileMetadata count: $($exportStatus.exportFileMetadata.Count)" -ForegroundColor Gray
                }
            }
            else {
                # For objects, use PSObject.Properties
                Write-Host "[DEBUG] exportStatus properties: $($exportStatus.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                Write-Host "[DEBUG] Has DownloadUrls property: $($exportStatus.PSObject.Properties.Name -contains 'DownloadUrls')" -ForegroundColor Gray
                Write-Host "[DEBUG] Has exportFileMetadata property: $($exportStatus.PSObject.Properties.Name -contains 'exportFileMetadata')" -ForegroundColor Gray
            }
            
            # Check for download URLs in either DownloadUrls (processed by Get-ExportStatus) or exportFileMetadata (raw API)
            if (($exportStatus.DownloadUrls -and $exportStatus.DownloadUrls.Count -gt 0) -or
                ($exportStatus.exportFileMetadata -and $exportStatus.exportFileMetadata.Count -gt 0)) {
                
                # Use DownloadUrls (processed) if available, otherwise fall back to exportFileMetadata (raw)
                if ($exportStatus.DownloadUrls -and $exportStatus.DownloadUrls.Count -gt 0) {
                    Write-Host "✅ Found $($exportStatus.DownloadUrls.Count) download URLs from processed DownloadUrls" -ForegroundColor Green
                    $downloadFiles = $exportStatus.DownloadUrls
                }
                else {
                    Write-Host "✅ Found $($exportStatus.exportFileMetadata.Count) download URLs from raw exportFileMetadata" -ForegroundColor Green
                    $downloadFiles = $exportStatus.exportFileMetadata
                }
                
                # Take the first available file (no PST preference - let user choose via menu)
                # Handle both processed format (FileName, DownloadUrl, Size) and raw format (fileName, downloadUrl, size)
                $firstFile = $downloadFiles[0]
                
                # Use the correct property name based on format
                $fileName = if ($firstFile.FileName) { $firstFile.FileName } else { $firstFile.fileName }
                $downloadUrl = if ($firstFile.DownloadUrl) { $firstFile.DownloadUrl } else { $firstFile.downloadUrl }
                $fileSize = if ($firstFile.Size) { $firstFile.Size } else { $firstFile.size }
                
                $hasDownloadUrl = $true
                Write-Host "Using first available file: $fileName (Size: $fileSize bytes)" -ForegroundColor Green
                
                # Update the output filename to match the original file extension
                $fileExtension = [System.IO.Path]::GetExtension($fileName)
                if (-not [string]::IsNullOrEmpty($fileExtension)) {
                    $outputFilePath = "$OutputPath\$sanitizedName$fileExtension"
                }
            }
            else {
                Write-Host "❌ No download URLs found in export status" -ForegroundColor Red
                Write-Host "This may indicate the export is not complete or the API structure has changed" -ForegroundColor Yellow
            }
            
            if ($hasDownloadUrl -and $downloadUrl) {
                
                Write-LogEntry -LogName $logPath -LogEntryText "Downloading export to: $outputFilePath" -LogLevel "INFO"
                Write-Host "Downloading export to: $outputFilePath" -ForegroundColor Yellow
                
                # Download the file
                try {
                    # Check the URL patterns to determine the right download method
                    $isPurviewUrl = $downloadUrl -like "*purview*" -or 
                    $downloadUrl -like "*purviewcases*" -or
                    $downloadUrl -like "*getAction*"
                                   
                    $isProxyUrl = $downloadUrl -like "*proxyservice.ediscovery*" -or 
                    $downloadUrl -like "*exportaedblobFileResult*"
                    
                    if ($isPurviewUrl) {
                        Write-Host "Detected Microsoft Purview URL. Using specialized handling..." -ForegroundColor Yellow
                        $downloaded = ConvertTo-M365PurviewDownload -DownloadUrl $downloadUrl -OutputFilePath $outputFilePath
                        if (-not $downloaded) {
                            throw "Failed to download from Purview URL"
                        }
                    }
                    # Proxy URLs can be downloaded directly with our enhanced download function
                    elseif ($isProxyUrl) {
                        Write-Host "Detected Microsoft 365 Compliance Export URL. Using specialized download method..." -ForegroundColor Yellow
                        $downloaded = Save-M365ComplianceFile -DownloadUrl $downloadUrl -OutputFilePath $outputFilePath
                        if (-not $downloaded -or -not $downloaded.Success) {
                            throw "Failed to download from Compliance Export URL"
                        }
                        
                        # Validate the downloaded file
                        if (-not (Test-Path -Path $outputFilePath)) {
                            throw "Failed to download file. Output path not created."
                        }
                        
                        $fileInfo = Get-Item -Path $outputFilePath
                        if ($fileInfo.Length -eq 0) {
                            throw "Downloaded file is empty (0 bytes)"
                        }
                        
                        # Special case for Report files - they use .zip extension but aren't true ZIP files
                        if ($outputFilePath -like "*Reports-*") {
                            Write-LogEntry -LogName $logPath -LogEntryText "Microsoft 365 Report file successfully downloaded: $outputFilePath" -LogLevel "INFO"
                            Write-Host "✅ Microsoft 365 Report file downloaded successfully: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                            Write-Host "Note: These Report files use .zip extension but are actually HTML/XML files." -ForegroundColor Yellow
                            
                            return @{
                                Success      = $true
                                FilePath     = $outputFilePath
                                DownloadDate = Get-Date
                                IsReport     = $true
                            }
                        }
                        
                        # Check if it's a ZIP file and validate it
                        if ($outputFilePath.EndsWith('.zip')) {
                            try {
                                # Use the enhanced validation with retry logic
                                if (Test-ValidZipFile -FilePath $outputFilePath -RetryOnFailure -MaxRetries 3 -RetryDelay 2) {
                                    Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded and validated: $outputFilePath" -LogLevel "INFO"
                                    Write-Host "✅ Download successful and ZIP file validated: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                    
                                    return @{
                                        Success      = $true
                                        FilePath     = $outputFilePath
                                        DownloadDate = Get-Date
                                        IsValidZip   = $true
                                    }
                                }
                                else {
                                    Write-LogEntry -LogName $logPath -LogEntryText "Downloaded file is not a valid ZIP file after multiple attempts: $outputFilePath" -LogLevel "ERROR"
                                    Write-Host "❌ Downloaded file is not a valid ZIP file after validation attempts." -ForegroundColor Red
                                    
                                    return @{
                                        Success    = $false
                                        FilePath   = $outputFilePath
                                        Error      = "Downloaded file is not a valid ZIP file"
                                        IsValidZip = $false
                                    }
                                }
                            }
                            catch {
                                Write-LogEntry -LogName $logPath -LogEntryText "Error validating ZIP file: $_" -LogLevel "ERROR"
                                Write-Host "Error validating ZIP file: $_" -ForegroundColor Red
                                return @{
                                    Success = $false
                                    Error   = "Error validating ZIP file: $_"
                                }
                            }
                        }
                        
                        # Regular file (not Report or ZIP)
                        Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded: $outputFilePath" -LogLevel "INFO"
                        Write-Host "✅ Download successful: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                        
                        return @{
                            Success      = $true
                            FilePath     = $outputFilePath
                            DownloadDate = Get-Date
                        }
                    }
                    else {
                        # Standard download method for non-M365 exports
                        # Get auth token headers for binary content
                        $headers = @{
                            "Accept" = "application/octet-stream"
                        }
                        
                        Write-Host "Starting standard binary download..." -ForegroundColor Yellow
                        
                        # Try multiple download methods in sequence
                        $downloadSuccess = $false
                        $downloadError = ""
                        
                        try {
                            # Method 1: WebClient
                            try {
                                $webClient = New-Object System.Net.WebClient
                                foreach ($key in $headers.Keys) {
                                    $webClient.Headers.Add($key, $headers[$key])
                                }
                                
                                $ProgressPreference = 'SilentlyContinue'
                                $webClient.DownloadFile($downloadUrl, $outputFilePath)
                                $ProgressPreference = 'Continue'
                                $downloadSuccess = $true
                            }
                            catch {
                                $downloadError = "WebClient: $_"
                                Write-Host "WebClient download failed: $_" -ForegroundColor Yellow
                            }
                            
                            # Method 2: BITS Transfer (if WebClient failed)
                            if (-not $downloadSuccess) {
                                try {
                                    if (Get-Module -ListAvailable -Name BitsTransfer) {
                                        Import-Module BitsTransfer
                                        Start-BitsTransfer -Source $downloadUrl -Destination $outputFilePath -DisplayName "Downloading export" -Priority High
                                        $downloadSuccess = $true
                                    }
                                }
                                catch {
                                    $downloadError += "`nBITS: $_"
                                    Write-Host "BITS Transfer download failed: $_" -ForegroundColor Yellow
                                }
                            }
                            
                            # Method 3: Invoke-WebRequest (if others failed)
                            if (-not $downloadSuccess) {
                                try {
                                    $ProgressPreference = 'SilentlyContinue'
                                    Invoke-WebRequest -Uri $downloadUrl -Headers $headers -OutFile $outputFilePath -Method Get -UseBasicParsing
                                    $ProgressPreference = 'Continue'
                                    $downloadSuccess = $true
                                }
                                catch {
                                    $downloadError += "`nInvoke-WebRequest: $_"
                                    Write-Host "Invoke-WebRequest download failed: $_" -ForegroundColor Red
                                    throw "All download methods failed: $downloadError"
                                }
                            }
                            
                            if (-not $downloadSuccess) {
                                throw "Download failed with all available methods: $downloadError"
                            }
                        }
                        catch {
                            # Clean up any partial downloads before re-throwing the error
                            if (Test-Path -Path $outputFilePath) {
                                Remove-Item -Path $outputFilePath -Force
                            }
                            throw # Re-throw to be caught by the outer catch block
                        }
                        
                        try {
                            # First validate file exists
                            if (-not (Test-Path -Path $outputFilePath)) {
                                throw "Failed to download file. Output path not created."
                            }
                            
                            # Then check file size
                            $fileInfo = Get-Item -Path $outputFilePath
                            if ($fileInfo.Length -eq 0) {
                                throw "Downloaded file is empty (0 bytes)"
                            }
                            
                            # Special case for Report files - they use .zip extension but aren't true ZIP files
                            if ($outputFilePath -like "*Reports-*") {
                                Write-LogEntry -LogName $logPath -LogEntryText "Microsoft 365 Report file successfully downloaded: $outputFilePath" -LogLevel "INFO"
                                Write-Host "? Microsoft 365 Report file downloaded successfully: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                Write-Host "Note: These Report files use .zip extension but are actually HTML/XML files." -ForegroundColor Yellow
                                
                                return @{
                                    Success      = $true
                                    FilePath     = $outputFilePath
                                    DownloadDate = Get-Date
                                    IsReport     = $true
                                }
                            }
                            
                            # Check if it's a ZIP file and validate it
                            if ($outputFilePath.EndsWith('.zip')) {
                                try {
                                    # Use the enhanced validation with retry logic
                                    if (Test-ValidZipFile -FilePath $outputFilePath -RetryOnFailure -MaxRetries 3 -RetryDelay 2) {
                                        Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded and validated: $outputFilePath" -LogLevel "INFO"
                                        Write-Host "? Download successful and ZIP file validated: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                        
                                        return @{
                                            Success      = $true
                                            FilePath     = $outputFilePath
                                            DownloadDate = Get-Date
                                            IsValidZip   = $true
                                        }
                                    }
                                    else {
                                        Write-LogEntry -LogName $logPath -LogEntryText "Downloaded file is not a valid ZIP file after multiple attempts: $outputFilePath" -LogLevel "ERROR"
                                        Write-Host "?? Downloaded file is not a valid ZIP file after validation attempts." -ForegroundColor Red
                                        
                                        return @{
                                            Success    = $false
                                            FilePath   = $outputFilePath
                                            Error      = "Downloaded file is not a valid ZIP file"
                                            IsValidZip = $false
                                        }
                                    }
                                }
                                catch {
                                    Write-LogEntry -LogName $logPath -LogEntryText "Error validating ZIP file: $_" -LogLevel "ERROR"
                                    Write-Host "Error validating ZIP file: $_" -ForegroundColor Red
                                    return @{
                                        Success = $false
                                        Error   = "Error validating ZIP file: $_"
                                    }
                                }
                            }
                            
                            # Regular file (not Report or ZIP)
                            Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded: $outputFilePath" -LogLevel "INFO"
                            Write-Host "? Download successful: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                            
                            return @{
                                Success      = $true
                                FilePath     = $outputFilePath
                                DownloadDate = Get-Date
                            }
                        }
                        catch {
                            Write-LogEntry -LogName $logPath -LogEntryText "Error validating download: $_" -LogLevel "ERROR"
                            Write-Host "?? Error validating download: $_" -ForegroundColor Red
                            
                            # Clean up if needed
                            if (Test-Path -Path $outputFilePath) {
                                Remove-Item -Path $outputFilePath -Force
                            }
                            
                            return @{
                                Success = $false
                                Error   = "$_"
                            }
                        }
                    }
                    catch {
                        Write-LogEntry -LogName $logPath -LogEntryText "Failed to download file: $_" -LogLevel "ERROR"
                        Write-Host "?? Download operation failed: $_" -ForegroundColor Red
                        
                        # Delete any partial downloads
                        if (Test-Path -Path $outputFilePath) {
                            Remove-Item -Path $outputFilePath -Force
                        }
                        return @{
                            Success = $false
                            Error   = "Download operation failed: $_"
                        }
                    }
                }
                catch {
                    Write-LogEntry -LogName $logPath -LogEntryText "Error in download process: $_" -LogLevel "ERROR"
                    Write-Host "Error in download process: $_" -ForegroundColor Red
                    
                    return @{
                        Success = $false
                        Error   = "Download process error: $_"
                    }
                }
            }
            else {
                Write-LogEntry -LogName $logPath -LogEntryText "Export is not ready for download. Current status: $($exportStatus.Status)" -LogLevel "WARNING"
                Write-Host "Export is not ready for download. Current status: $($exportStatus.Status)" -ForegroundColor Yellow
    
                return @{
                    Success = $false
                    Status  = $exportStatus.Status
                    Error   = "Export is not ready for download. Current status: $($exportStatus.Status)"
                }
            }
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error in Save-ExportResults function: $_" -LogLevel "ERROR"
        Write-Host "Error in Save-ExportResults function: $_" -ForegroundColor Red
        return @{
            Success = $false
            Error   = "Function error: $_"
        }
    }
}

# Function to wait for export to complete
Function Wait-ForExportCompletion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportId,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 60,
        
        [Parameter(Mandatory = $false)]
        [int]$WaitTimeSeconds = 30
    )
    
    Write-Host "Monitoring export operation with ID: $ExportId" -ForegroundColor Yellow
    Write-LogEntry -LogName $logPath -LogEntryText "Starting export monitoring with export ID: $ExportId" -LogLevel "INFO"
    
    $attempts = 0
    $status = ""
    
    do {
        $attempts++
        Write-Host "Checking export status (attempt $attempts/$MaxAttempts)..." -ForegroundColor Yellow
        
        # Wait before checking again
        if ($attempts -gt 1) {
            Write-Host "Waiting $WaitTimeSeconds seconds before next check..." -ForegroundColor Gray
            Start-Sleep -Seconds $WaitTimeSeconds
        }
        
        try {
            $exportStatus = Get-ExportStatus -CaseId $CaseId -ExportId $ExportId
            $status = $exportStatus.Status
            
            Write-LogEntry -LogName $logPath -LogEntryText "Current export status (attempt $attempts/$MaxAttempts): $status" -LogLevel "INFO"
            Write-Host "Current export status (attempt $attempts/$MaxAttempts): $status" -ForegroundColor Cyan
        }
        catch {
            Write-LogEntry -LogName $logPath -LogEntryText "Error checking export operation (attempt $attempts/$MaxAttempts): $_" -LogLevel "WARNING"
            Write-Host "Error checking export operation (attempt $attempts/$MaxAttempts): $_" -ForegroundColor Yellow
        }
        
        # Exit conditions: status complete or maximum attempts reached
        if ($status -eq "succeeded" -or $status -eq "failed" -or $status -eq "partiallySucceeded" -or $attempts -ge $MaxAttempts) {
            break
        }
    } while ($true)
    
    return @{
        Status        = $status
        Attempts      = $attempts
        MaxAttempts   = $MaxAttempts
        CompletedTime = Get-Date
    }
}

# Function to list existing eDiscovery cases
Function Get-eDiscoveryCases {
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases?`$select=displayName,id,status,createdDateTime"
        
        $casesResponse = Invoke-MgGraphRequest -Method GET -Uri $apiUrl
        return $casesResponse.value
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error retrieving eDiscovery cases: $_" -LogLevel "ERROR"
        Write-Host "Error retrieving eDiscovery cases: $_" -ForegroundColor Red
        return @()
    }
}

# Function to check if a search is ready for export by examining its statistics and related exports
Function Get-SearchReadinessStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$SearchId
    )
    
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        
        # First, check if there are any successful exports based on this search
        try {
            $exportsUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations?`$filter=action eq 'caseExport'"
            $exportsResponse = Invoke-MgGraphRequest -Method GET -Uri $exportsUrl -ErrorAction SilentlyContinue
            
            if ($exportsResponse.value -and $exportsResponse.value.Count -gt 0) {
                # Look for exports that reference this search
                $relatedExports = $exportsResponse.value | Where-Object {
                    # Check if this export is related to our search ID
                    ($_.PSObject.Properties.Name -contains "dataSourceScopes" -and $_.dataSourceScopes -like "*$SearchId*") -or
                    ($_.PSObject.Properties.Name -contains "reviewSetQuery" -and $_.reviewSetQuery -like "*$SearchId*") -or
                    ($_.PSObject.Properties.Name -contains "sourceId" -and $_.sourceId -eq $SearchId)
                }
                
                if ($relatedExports -and $relatedExports.Count -gt 0) {
                    # Check the status of the most recent related export
                    $latestExport = $relatedExports | Sort-Object createdDateTime -Descending | Select-Object -First 1
                    if ($latestExport.PSObject.Properties.Name -contains "status") {
                        switch ($latestExport.status.ToLower()) {
                            "succeeded" { return "Export Available (Status: $($latestExport.status))" }
                            "running" { return "Export In Progress" }
                            "failed" { return "Export Failed" }
                            default { return "Export Status: $($latestExport.status)" }
                        }
                    }
                }
            }
        }
        catch {
            # Continue with other status checks if export check fails
            Write-LogEntry -LogName $logPath -LogEntryText "Could not check export status for search $SearchId`: $_" -LogLevel "DEBUG"
        }
        
        # Try to get the search details
        $searchUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/searches/$SearchId"
        $searchDetails = Invoke-MgGraphRequest -Method GET -Uri $searchUrl
        
        # Check if the search has lastEstimateStatisticsOperation
        if ($searchDetails.PSObject.Properties.Name -contains "lastEstimateStatisticsOperation" -and 
            $null -ne $searchDetails.lastEstimateStatisticsOperation) {
            
            $operation = $searchDetails.lastEstimateStatisticsOperation
            
            # Check the operation status
            if ($operation.PSObject.Properties.Name -contains "status") {
                switch ($operation.status.ToLower()) {
                    "succeeded" { 
                        # Check if we have results count to confirm it's truly ready
                        if ($operation.PSObject.Properties.Name -contains "resultInfo" -and 
                            $null -ne $operation.resultInfo -and 
                            $operation.resultInfo.PSObject.Properties.Name -contains "estimatedCount") {
                            return "Ready for Export (Results: $($operation.resultInfo.estimatedCount))"
                        }
                        return "Ready for Export"
                    }
                    "running" { return "Processing Statistics..." }
                    "failed" { return "Statistics Failed" }
                    default { return "Status: $($operation.status)" }
                }
            }
        }
        
        # Try to get all operations and look for estimate operations manually
        try {
            $operationsUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations"
            $operationsResponse = Invoke-MgGraphRequest -Method GET -Uri $operationsUrl -ErrorAction SilentlyContinue
            
            if ($operationsResponse.value -and $operationsResponse.value.Count -gt 0) {
                # Look for estimate operations that might be related to this search
                $estimateOps = $operationsResponse.value | Where-Object { 
                    $_.action -like "*estimate*" -and 
                    $null -ne $_.createdDateTime 
                } | Sort-Object createdDateTime -Descending
                
                if ($estimateOps -and $estimateOps.Count -gt 0) {
                    # Take the most recent estimate operation
                    $latestEstimate = $estimateOps[0]
                    if ($latestEstimate.PSObject.Properties.Name -contains "status") {
                        switch ($latestEstimate.status.ToLower()) {
                            "succeeded" { return "Likely Ready for Export" }
                            "running" { return "Processing..." }
                            "failed" { return "Processing Failed" }
                            default { return "Status: $($latestEstimate.status)" }
                        }
                    }
                }
            }
        }
        catch {
            # Continue with time-based fallback
            Write-LogEntry -LogName $logPath -LogEntryText "Could not check operations for search $SearchId`: $_" -LogLevel "DEBUG"
        }
        
        # If no estimate operations found, check creation time and assume readiness based on age
        if ($searchDetails.PSObject.Properties.Name -contains "createdDateTime") {
            try {
                $createdDate = [DateTime]::Parse($searchDetails.createdDateTime)
                $timeSinceCreated = (Get-Date) - $createdDate
                
                if ($timeSinceCreated.TotalMinutes -lt 5) {
                    return "Recently Created - Processing"
                }
                elseif ($timeSinceCreated.TotalMinutes -lt 30) {
                    return "Likely Processing - Check Later" 
                }
                else {
                    # After 30 minutes, assume it's ready for export
                    return "Likely Ready for Export (Created: $($createdDate.ToString('MM/dd/yyyy HH:mm')))"
                }
            }
            catch {
                # If date parsing fails, fall back to ready
                return "Search Available"
            }
        }
        
        # Default fallback - be more optimistic
        return "Search Available - Try Export"
    }
    catch {
        # Don't show the detailed error to avoid cluttering output, but log it
        Write-LogEntry -LogName $logPath -LogEntryText "Could not check search readiness for SearchId $SearchId`: $_" -LogLevel "WARNING"
        return "Search Available"
    }
}

# Function to list searches in a case
Function Get-eDiscoverySearches {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId
    )
    
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        # First try without select to see all available properties
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/searches"
        Write-Host "Request URL: $apiUrl" -ForegroundColor Yellow
        
        $searchesResponse = Invoke-MgGraphRequest -Method GET -Uri $apiUrl
        Write-Host "Found $($searchesResponse.value.Count) searches" -ForegroundColor Cyan
        
        # Add diagnostic output to see what's coming back
        if ($searchesResponse.value.Count -gt 0) {
            Write-Host "First search details:" -ForegroundColor Yellow
            $firstSearch = $searchesResponse.value[0]
            Write-Host "All available properties:" -ForegroundColor Cyan
            $firstSearch.PSObject.Properties.Name | ForEach-Object {
                Write-Host "  $_`: $($firstSearch.$_)" -ForegroundColor Yellow
            }
            
            # Enhance the search objects with status information
            $enhancedSearches = @()
            foreach ($search in $searchesResponse.value) {
                # Add a copy of the search to our results
                $enhancedSearch = $search | Select-Object *
                
                # If display name is empty, try to generate one from other properties
                if ([string]::IsNullOrEmpty($enhancedSearch.displayName)) {
                    if (![string]::IsNullOrEmpty($enhancedSearch.description)) {
                        $enhancedSearch | Add-Member -MemberType NoteProperty -Name "displayName" -Value "Search: $($enhancedSearch.description)" -Force
                    }
                    else {
                        $enhancedSearch | Add-Member -MemberType NoteProperty -Name "displayName" -Value "Search ID: $($enhancedSearch.id.Substring(0, 8))" -Force
                    }
                    Write-Host "Added display name: $($enhancedSearch.displayName)" -ForegroundColor Green
                }
                
                # Add search readiness status by checking for estimate statistics
                try {
                    $searchStatus = Get-SearchReadinessStatus -CaseId $CaseId -SearchId $search.id
                    $enhancedSearch | Add-Member -MemberType NoteProperty -Name "ReadinessStatus" -Value $searchStatus -Force
                }
                catch {
                    Write-Host "Warning: Could not determine readiness for search $($search.displayName): $_" -ForegroundColor Yellow
                    $enhancedSearch | Add-Member -MemberType NoteProperty -Name "ReadinessStatus" -Value "Unknown" -Force
                }
                
                $enhancedSearches += $enhancedSearch
            }
            
            return $enhancedSearches
        }
        else {
            return $searchesResponse.value
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error retrieving searches for case ${CaseId}: $_" -LogLevel "ERROR"
        Write-Host "Error retrieving searches for case ${CaseId}: $_" -ForegroundColor Red
        return @()
    }
}

# Function to list exports in a case
Function Get-eDiscoveryExports {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId
    )
    
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        Write-Host "Using API version: $apiVersion" -ForegroundColor Yellow
        
        # Get the list of operations first to filter for exports
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations?`$select=id,action"
        $exportsResponse = Invoke-MgGraphRequest -Method GET -Uri $apiUrl
        
        # Filter operations to only include export operations
        $exportOperations = @()
        if ($null -ne $exportsResponse -and $null -ne $exportsResponse.value) {
            $exportOperations = @($exportsResponse.value | Where-Object { $_.action -like "*export*" })
            Write-Host "Found $($exportOperations.Count) export operations" -ForegroundColor Cyan
        }
        else {
            Write-Host "No export operations found" -ForegroundColor Yellow
        }
        
        # Debug the raw response to see if we're getting empty objects
        if ($DebugOutput) {
            Write-Host "DEBUG: Raw export operations:" -ForegroundColor Yellow
            $exportOperations | ForEach-Object { 
                Write-Host "  ID: $($_.id), Action: $($_.action)" -ForegroundColor Gray 
            }
        }
        
        # Create an array to store detailed export objects
        $detailedExports = @()
        
        # Get FULL details for each export by directly querying the specific operation
        foreach ($basicExport in $exportOperations) {
            try {
                # Query the individual export directly - this is what gives us the display name
                $detailUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations/$($basicExport.id)"
                if ($DebugOutput) {
                    Write-Host "Getting full export details from: $detailUrl" -ForegroundColor Yellow
                }
                
                # Get the complete export details directly - this has all the information including display name
                $exportDetail = Invoke-MgGraphRequest -Method GET -Uri $detailUrl
                
                # Debug the raw response to see what's available
                if ($DebugOutput) {
                    Write-Host "Raw export detail data:" -ForegroundColor Yellow
                    $exportDetailJson = $exportDetail | ConvertTo-Json -Depth 3 -Compress
                    Write-Host $exportDetailJson -ForegroundColor Cyan
                }
                
                # Access displayName directly without property checking
                if ($null -eq $exportDetail.displayName) {
                    # Try to get it from various possible locations
                    if ($null -ne $exportDetail.resultInfo -and $null -ne $exportDetail.resultInfo.displayName) {
                        $exportDetail | Add-Member -MemberType NoteProperty -Name "displayName" -Value $exportDetail.resultInfo.displayName -Force
                        Write-Host "Added display name from resultInfo: $($exportDetail.resultInfo.displayName)" -ForegroundColor Green
                    }
                    # Check additional data
                    elseif ($null -ne $exportDetail.additionalData -and $null -ne $exportDetail.additionalData.displayName) {
                        $exportDetail | Add-Member -MemberType NoteProperty -Name "displayName" -Value $exportDetail.additionalData.displayName -Force
                        Write-Host "Added display name from additionalData: $($exportDetail.additionalData.displayName)" -ForegroundColor Green
                    }
                    # If no display name is found, create one from other properties
                    else {
                        $createdDate = if ($exportDetail.createdDateTime) { 
                            try {
                                $dateString = $exportDetail.createdDateTime.ToString("yyyyMMdd_HHmmss")
                                $dateString
                            }
                            catch {
                                "unknown_date"
                            }
                        }
                        else { "unknown_date" }
                        
                        $exportDetail | Add-Member -MemberType NoteProperty -Name "displayName" -Value "Export_$createdDate" -Force
                        Write-Host "Created synthetic display name: Export_$createdDate" -ForegroundColor Yellow
                    }
                }
                else {
                    Write-Host "Using actual display name from API response: $($exportDetail.displayName)" -ForegroundColor Green
                }
                
                # Add the detailed export to our collection
                $detailedExports += $exportDetail
                
                if ($DebugOutput) {
                    Write-Host "? Successfully added export: $($exportDetail.displayName)" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "? Error getting details for export $($basicExport.id): $_" -ForegroundColor Red
                # Still add basic info if available
                try {
                    # Add a display name property to avoid issues in Select-Export
                    $basicExport | Add-Member -MemberType NoteProperty -Name "displayName" -Value "Export_$($basicExport.id.Substring(0, 8))" -Force
                    $detailedExports += $basicExport
                    Write-Host "Added basic export with generated name: Export_$($basicExport.id.Substring(0, 8))" -ForegroundColor Yellow
                }
                catch {
                    # If even this fails, skip the export
                    Write-Host "Could not process export $($basicExport.id), skipping" -ForegroundColor Red
                }
            }
        }
        
        # Simply return the detailed exports we've collected
        Write-Host "Returning $($detailedExports.Count) detailed exports" -ForegroundColor Yellow
        
        # Filter out any null or incomplete exports
        $validExports = $detailedExports | Where-Object { 
            $_ -ne $null -and 
            ($null -ne $_.id -or $null -ne $_.displayName -or $null -ne $_.action)
        }
        
        if ($validExports.Count -lt $detailedExports.Count) {
            Write-Host "Filtered out $($detailedExports.Count - $validExports.Count) null/invalid exports" -ForegroundColor Yellow
            $detailedExports = $validExports
        }
        
        # Verify what we're returning
        if ($DebugOutput) {
            Write-Host "DEBUG: Final export count: $($detailedExports.Count)" -ForegroundColor Cyan
            foreach ($export in $detailedExports) {
                if ($null -ne $export.displayName) {
                    Write-Host "Export ID: $($export.id), Display Name: $($export.displayName)" -ForegroundColor Green
                }
                else {
                    Write-Host "Export ID: $($export.id), No display name available" -ForegroundColor Yellow
                }
            }
        }
        
        # Return exactly what we've built - ensure it stays as an array with comma operator
        return , $detailedExports
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error retrieving exports for case ${CaseId}: $_" -LogLevel "ERROR"
        Write-Host "Error retrieving exports for case ${CaseId}: $_" -ForegroundColor Red
        return @()
    }
}

# Main menu and operation handling
Function Show-MainMenu {
    Clear-Host
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "�       Exchange Calendar Export Utility          �" -ForegroundColor Cyan
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host
    Write-Host " Select an operation:" -ForegroundColor Yellow
    Write-Host
    Write-Host " [1] Create new eDiscovery case for mailbox export" -ForegroundColor Green
    Write-Host " [2] Create export for existing case/search" -ForegroundColor Green
    Write-Host " [3] View exports and check status/attachments" -ForegroundColor Green
    Write-Host " [4] Download completed export" -ForegroundColor Green
    Write-Host " [5] List existing cases" -ForegroundColor Green
    Write-Host " [6] List existing searches in a case" -ForegroundColor Green
    Write-Host " [0] Exit" -ForegroundColor Green
    Write-Host
    Write-Host -NoNewline " Enter your choice [0-6]: "
    
    $choice = Read-Host
    return $choice
}

# Function to list export attachments
Function Get-ExportAttachments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportId
    )
    
    try {
        $apiVersion = if ($Version -eq "prod") { "v1.0" } else { "beta" }
        $apiUrl = "https://graph.microsoft.com/$apiVersion/security/cases/ediscoveryCases/$CaseId/operations/$ExportId"
        Write-Host "Fetching export attachments from: $apiUrl" -ForegroundColor Yellow
        $exportDetails = Invoke-MgGraphRequest -Method GET -Uri $apiUrl
        
        # Just dump the raw object to verify we have the exportFileMetadata
        Write-Host "Raw export response for attachments:" -ForegroundColor Yellow
        $rawJson = $exportDetails | ConvertTo-Json -Depth 3
        Write-Host $rawJson -ForegroundColor Cyan
        
        # Check for the export files/attachments - direct property access
        if ($null -ne $exportDetails.exportFileMetadata -and $exportDetails.exportFileMetadata.Count -gt 0) {
            Write-Host "Found $($exportDetails.exportFileMetadata.Count) attachments in exportFileMetadata" -ForegroundColor Green
            return $exportDetails.exportFileMetadata
        }
        # Check by validating properties first
        elseif ($exportDetails.PSObject.Properties.Name -contains "exportFileMetadata") {
            if ($null -ne $exportDetails.exportFileMetadata -and $exportDetails.exportFileMetadata.Count -gt 0) {
                Write-Host "Found $($exportDetails.exportFileMetadata.Count) attachments in exportFileMetadata" -ForegroundColor Green
                return $exportDetails.exportFileMetadata
            }
        }
        # Check alternate locations - using safer property access
        elseif ($exportDetails.PSObject.Properties.Name -contains "resultInfo" -and 
            $null -ne $exportDetails.resultInfo -and
            $exportDetails.resultInfo.PSObject.Properties.Name -contains "additionalProperties" -and
            $null -ne $exportDetails.resultInfo.additionalProperties -and
            $exportDetails.resultInfo.additionalProperties.PSObject.Properties.Name -contains "exportFiles") {
            
            $exportFiles = $exportDetails.resultInfo.additionalProperties.exportFiles
            if ($null -ne $exportFiles -and $exportFiles.Count -gt 0) {
                Write-Host "Found $($exportFiles.Count) attachments in resultInfo.additionalProperties.exportFiles" -ForegroundColor Green
                return $exportFiles
            }
        }
        # Explicitly check for "exportFiles" property
        elseif ($exportDetails.PSObject.Properties.Name -contains "exportFiles") {
            if ($null -ne $exportDetails.exportFiles -and $exportDetails.exportFiles.Count -gt 0) {
                Write-Host "Found $($exportDetails.exportFiles.Count) attachments in exportFiles" -ForegroundColor Green
                return $exportDetails.exportFiles
            }
        }
        
        # If we've reached here, we need to scan the raw data more carefully
        Write-Host "Standard property checks failed to find attachments, performing deeper scan..." -ForegroundColor Yellow
        
        # Convert the object to JSON and search for specific patterns
        $jsonData = $exportDetails | ConvertTo-Json -Depth 10
        if ($jsonData -match '"exportFileMetadata"\s*:\s*\[') {
            Write-Host "Found exportFileMetadata in JSON data, attempting to extract manually" -ForegroundColor Yellow
            
            try {
                # Try to extract directly from the raw response
                $pattern = '"exportFileMetadata"\s*:\s*(\[.*?\])'
                if ($jsonData -match $pattern) {
                    $match = $matches[1]
                    $extractedJson = $match
                    $attachments = $extractedJson | ConvertFrom-Json
                    Write-Host "Manually extracted $($attachments.Count) attachments" -ForegroundColor Green
                    return $attachments
                }
            }
            catch {
                Write-Host "Error manually extracting attachments: $_" -ForegroundColor Red
            }
        }
        
        Write-Host "No attachments found for this export." -ForegroundColor Yellow
        return @()
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error retrieving export attachments: $_" -LogLevel "ERROR"
        Write-Host "Error retrieving export attachments: $_" -ForegroundColor Red
        return @()
    }
}

# Function to select a case from a list
Function Select-Case {
    $cases = Get-eDiscoveryCases
    
    if ($cases.Count -eq 0) {
        Write-Host "No eDiscovery cases found." -ForegroundColor Red
        return $null
    }
    
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "�              Available Cases                    �" -ForegroundColor Cyan
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host
    Write-Host "Select a case by number:" -ForegroundColor Yellow
    Write-Host
    
    for ($i = 0; $i -lt $cases.Count; $i++) {
        Write-Host " [$($i+1)] $($cases[$i].displayName) (ID: $($cases[$i].id))"
    }
    
    Write-Host " [0] Cancel" -ForegroundColor Yellow
    Write-Host
    Write-Host -NoNewline " Select a case [0-$($cases.Count)]: "
    
    $choice = Read-Host
    
    # Convert choice to integer and validate
    try {
        $choiceNum = [int]$choice
        if ($choiceNum -eq 0) {
            return $null  # User cancelled
        }
        elseif ($choiceNum -ge 1 -and $choiceNum -le $cases.Count) {
            return $cases[$choiceNum - 1]  # Return the actual case object
        }
        else {
            Write-Host "Invalid choice. Please select a number between 0 and $($cases.Count)." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Invalid input. Please enter a number." -ForegroundColor Red
        return $null
    }
}

# Function to select a search from a list
# Function to select a search from a list
Function Select-Search {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CaseId
    )
    
    $searches = Get-eDiscoverySearches -CaseId $CaseId
    
    if ($DebugOutput) {
        Write-Host "DEBUG: Retrieved $($searches.Count) searches from Get-eDiscoverySearches" -ForegroundColor Yellow
    }
    
    if ($searches.Count -eq 0) {
        Write-Host "No searches found in this case." -ForegroundColor Red
        return $null
    }
    
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "�              Available Searches                 �" -ForegroundColor Cyan
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host
    
    for ($i = 0; $i -lt $searches.Count; $i++) {
        # Format the creation date if available
        $createdDate = if ($searches[$i].createdDateTime) {
            try {
                if ($searches[$i].createdDateTime -is [DateTime]) {
                    $searches[$i].createdDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                }
                elseif ($searches[$i].createdDateTime -is [string] -and $searches[$i].createdDateTime -match "T") {
                    $searches[$i].createdDateTime.Substring(0, 19).Replace("T", " ")
                }
                else {
                    $searches[$i].createdDateTime
                }
            }
            catch {
                "Unknown date"
            }
        }
        else {
            "Unknown date"
        }
        
        # Show search description if available
        $descriptionInfo = if (![string]::IsNullOrEmpty($searches[$i].description)) {
            " ($($searches[$i].description))"
        }
        else {
            ""
        }
        
        # Make sure we always have a display name to show
        $displayName = if (![string]::IsNullOrEmpty($searches[$i].displayName)) {
            $searches[$i].displayName
        }
        elseif (![string]::IsNullOrEmpty($searches[$i].description)) {
            "Search: $($searches[$i].description)"
        }
        else {
            "Search ID: $($searches[$i].id.Substring(0, 8))"
        }
        
        Write-Host " [$($i+1)] $displayName$descriptionInfo (Created: $createdDate) (ID: $($searches[$i].id))"
    }
    
    Write-Host " [0] Cancel" -ForegroundColor Yellow
    Write-Host
    Write-Host -NoNewline " Select a search [0-$($searches.Count)]: "
    
    $choice = Read-Host
    
    if ($choice -eq "0") {
        return $null
    }
    
    try {
        $index = [int]$choice - 1
        if ($index -ge 0 -and $index -lt $searches.Count) {
            return $searches[$index]
        }
        else {
            Write-Host "Invalid selection." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Invalid input. Please enter a number." -ForegroundColor Red
        return $null
    }
}

# Function to select an export from a list
Function Select-Export {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CaseId
    )
    
    # Get the detailed export list using our improved function
    $exports = @()  # Initialize as empty array
    try {
        $result = Get-eDiscoveryExports -CaseId $CaseId
        # Ensure we have a non-empty array
        if ($null -ne $result -and ($result -is [array] -or $result -is [System.Collections.ArrayList])) {
            $exports = @($result)  # Force to array with @ operator
        }
        elseif ($null -ne $result) {
            # Single item result (not an array)
            $exports = @($result)  # Create single-item array
        }
    }
    catch {
        Write-Host "Error getting exports: $_" -ForegroundColor Red
        $exports = @()  # Empty array on error
    }
    
    if ($DebugOutput) {
        Write-Host "DEBUG: Raw received $($exports.Count) exports from Get-eDiscoveryExports" -ForegroundColor Yellow
    }
    
    # Filter out any null or invalid exports
    $validExports = @($exports | Where-Object { 
            $_ -ne $null -and 
            ($null -ne $_.id -or $null -ne $_.displayName -or $null -ne $_.action)
        })
    
    if ($DebugOutput) {
        Write-Host "DEBUG: After filtering, $($validExports.Count) valid exports remain" -ForegroundColor Yellow
    }
    $exports = $validExports
    
    if ($exports.Count -eq 0) {
        Write-Host "No exports found in this case." -ForegroundColor Red
        return $null
    }
    
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host "�              Available Exports                  �" -ForegroundColor Cyan
    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
    Write-Host
    
    for ($i = 0; $i -lt $exports.Count; $i++) {
        # Extract creation date, ensure it's formatted properly
        $createdDate = if ($exports[$i].createdDateTime) { 
            # Check if createdDateTime is a string or DateTime object and format accordingly
            if ($exports[$i].createdDateTime -is [DateTime]) {
                $exports[$i].createdDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            } 
            elseif ($exports[$i].createdDateTime -is [string]) {
                if ($exports[$i].createdDateTime -match "T") {
                    $exports[$i].createdDateTime.Substring(0, 19).Replace("T", " ")
                }
                else {
                    $exports[$i].createdDateTime
                }
            }
            else {
                "Unknown date"
            }
        }
        else { 
            "Unknown date" 
        }
        
        # Extract status - try direct property access to each possible property
        $status = "unknown"
        if ($null -ne $exports[$i].status) {
            $status = $exports[$i].status
        }
        elseif ($null -ne $exports[$i].Status) {
            $status = $exports[$i].Status
        }
        elseif ($null -ne $exports[$i].state) {
            $status = $exports[$i].state
        }
        elseif ($null -ne $exports[$i].operationStatus) {
            $status = $exports[$i].operationStatus
        }
        
        # Get the display name - using direct property access and falling back to ID if needed
        # Debug: Check what's in the export object
        if ($DebugOutput) {
            Write-Host "DEBUG: Export $i properties:" -ForegroundColor Yellow
            $exports[$i] | Format-List -Property * | Out-String | Write-Host -ForegroundColor Gray
        }
        
        $displayName = if ($null -ne $exports[$i].displayName -and -not [string]::IsNullOrEmpty($exports[$i].displayName)) {
            if ($DebugOutput) {
                Write-Host "DEBUG: Using displayName: $($exports[$i].displayName)" -ForegroundColor Green
            }
            $exports[$i].displayName
        } 
        else {
            if ($DebugOutput) {
                Write-Host "DEBUG: No displayName found, using fallback" -ForegroundColor Yellow
            }
            # Safely handle null ID values
            if ($null -ne $exports[$i].id -and -not [string]::IsNullOrEmpty($exports[$i].id)) {
                "Export_$($exports[$i].id.Substring(0, [Math]::Min(8, $exports[$i].id.Length)))"
            }
            else {
                "Export_$($i)" # Use index as fallback if ID is missing
            }
        }
        
        # Write out the export details
        $idDisplay = if ($null -ne $exports[$i].id -and -not [string]::IsNullOrEmpty($exports[$i].id)) { 
            $exports[$i].id 
        }
        else { 
            "Unknown" 
        }
        Write-Host " [$($i+1)] Name: $displayName (Created: $createdDate) (Status: $status) (ID: $idDisplay)"
    }
    
    Write-Host " [0] Cancel" -ForegroundColor Yellow
    Write-Host
    Write-Host -NoNewline " Select an export [0-$($exports.Count)]: "
    
    $choice = Read-Host
    
    if ($choice -eq "0") {
        return $null
    }
    
    try {
        $index = [int]$choice - 1
        if ($index -ge 0 -and $index -lt $exports.Count) {
            return $exports[$index]
        }
        else {
            Write-Host "Invalid selection." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Invalid input. Please enter a number." -ForegroundColor Red
        return $null
    }
}

# Process operations based on command line parameters or menu selection
if ($Operation -eq "menu" -and -not $CaseId -and -not $SearchId -and -not $ExportId) {
    # Show interactive menu if no specific operation or IDs are provided
    $exit = $false
    
    while (-not $exit) {
        $choice = Show-MainMenu
        
        switch ($choice) {
            "0" {
                $exit = $true
                Write-Host "Exiting script." -ForegroundColor Yellow
            }
            "1" {
                # Create new case
                Write-Host "Creating new eDiscovery case for mailbox export..." -ForegroundColor Yellow
                $result = New-CalendarExportCase -UserEmail $EmailAddress -ContentTypeFilter $ContentType
                
                if ($result) {
                    Write-Host "Case and search created successfully." -ForegroundColor Green
                    Write-Host "Case ID: $($result.CaseId)" -ForegroundColor Cyan
                    Write-Host "Search ID: $($result.SearchId)" -ForegroundColor Cyan
                    Write-Host
                    Write-Host "To export this search later, use:" -ForegroundColor Yellow
                    Write-Host ".\exchange-cal-export.ps1 -Operation export -CaseId $($result.CaseId) -SearchId $($result.SearchId)" -ForegroundColor Yellow
                    
                    # Prompt to start export now
                    Write-Host
                    $startExport = Read-Host "Do you want to start the export now? (y/n)"
                    if ($startExport -eq "y" -or $startExport -eq "Y") {
                        $exportResult = Export-CalendarResults -CaseId $result.CaseId -SearchId $result.SearchId -SearchName $result.Name
                        
                        if ($exportResult -and $exportResult.ExportId) {
                            Write-Host "Export started successfully." -ForegroundColor Green
                            Write-Host "Export ID: $($exportResult.ExportId)" -ForegroundColor Cyan
                            Write-Host
                            Write-Host "To check export status later, use:" -ForegroundColor Yellow
                            Write-Host ".\exchange-cal-export.ps1 -Operation export -CaseId $($result.CaseId) -ExportId $($exportResult.ExportId)" -ForegroundColor Yellow
                        }
                    }
                }
                
                Write-Host "Press any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "2" {
                # Create export for existing case/search
                try {
                    $case = Select-Case
                    if ($case) {
                        $search = Select-Search -CaseId $case.id
                        if ($search) {
                            $exportResult = Export-CalendarResults -CaseId $case.id -SearchId $search.id -SearchName $search.displayName
                            
                            if ($exportResult -and $exportResult.ExportId) {
                                Write-Host "Export started successfully." -ForegroundColor Green
                                Write-Host "Export ID: $($exportResult.ExportId)" -ForegroundColor Cyan
                                Write-Host
                                Write-Host "To check export status later, use:" -ForegroundColor Yellow
                                Write-Host ".\exchange-cal-export.ps1 -Operation export -CaseId $($case.id) -ExportId $($exportResult.ExportId)" -ForegroundColor Yellow
                            }
                        }
                    }
                }
                catch {
                    Write-Host "Error during export operation: $_" -ForegroundColor Red
                    Write-LogEntry -LogName $logPath -LogEntryText "Error during export operation: $_" -LogLevel "ERROR"
                }
                
                Write-Host "Press any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "3" {
                # Check export status and view attachments
                $case = Select-Case
                if ($case) {
                    $export = Select-Export -CaseId $case.id
                    if ($export) {
                        $exportStatus = Get-ExportStatus -CaseId $case.id -ExportId $export.id
                        
                        Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                        Write-Host "�              Export Details                     �" -ForegroundColor Cyan
                        Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                        Write-Host
                        Write-Host "Export Name: $($exportStatus.DisplayName)" -ForegroundColor Green
                        Write-Host "Export Status: $($exportStatus.Status)" -ForegroundColor Cyan
                        Write-Host "Created By: $($exportStatus.CreatedBy)" -ForegroundColor Cyan
                        Write-Host "Created Date: $($exportStatus.CreatedDate)" -ForegroundColor Cyan
                        Write-Host
                        
                        # Get the export details directly from the status result
                        $exportDetails = $exportStatus.ExportDetails
                        
                        # Direct access to exportFileMetadata without property checking
                        $attachments = $null
                        if ($null -ne $exportDetails.exportFileMetadata) {
                            $attachments = $exportDetails.exportFileMetadata
                            Write-Host "Using $($attachments.Count) attachments found directly in exportFileMetadata" -ForegroundColor Green
                        }
                        else {
                            # If not found, fall back to Get-ExportAttachments
                            Write-Host "Direct attachment access failed, using Get-ExportAttachments..." -ForegroundColor Yellow
                            $attachments = Get-ExportAttachments -CaseId $case.id -ExportId $export.id
                        }
                        
                        if ($attachments.Count -eq 0) {
                            Write-Host "No attachments found for this export." -ForegroundColor Red
                        }
                        else {
                            Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                            Write-Host "�              Export Attachments                 �" -ForegroundColor Cyan
                            Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                            Write-Host
                            
                            for ($i = 0; $i -lt $attachments.Count; $i++) {
                                # Extract the file size in a readable format
                                # Debug: Display raw attachment data
                                if ($DebugOutput) {
                                    Write-Host "DEBUG: Attachment $($i+1) data:" -ForegroundColor Yellow
                                    $attachments[$i] | Format-List -Property * | Out-String | Write-Host -ForegroundColor Gray
                                }
                                
                                # Direct property access for file size
                                $fileSize = if ($null -ne $attachments[$i].size) {
                                    $sizeInBytes = $attachments[$i].size
                                    if ($sizeInBytes -ge 1GB) {
                                        "{0:N2} GB" -f ($sizeInBytes / 1GB)
                                    }
                                    elseif ($sizeInBytes -ge 1MB) {
                                        "{0:N2} MB" -f ($sizeInBytes / 1MB)
                                    }
                                    elseif ($sizeInBytes -ge 1KB) {
                                        "{0:N2} KB" -f ($sizeInBytes / 1KB)
                                    }
                                    else {
                                        "$sizeInBytes bytes"
                                    }
                                }
                                else {
                                    "Unknown size"
                                }
                                
                                # Get the file name - direct property access
                                $fileName = if ($null -ne $attachments[$i].fileName) {
                                    $attachments[$i].fileName
                                }
                                else {
                                    "Attachment $($i+1)"
                                }
                                
                                # Get the download URL - direct property access
                                $hasDownloadUrl = $null -ne $attachments[$i].downloadUrl -and -not [string]::IsNullOrEmpty($attachments[$i].downloadUrl)
                                $downloadStatus = if ($hasDownloadUrl) {
                                    "Available for download"
                                }
                                else {
                                    "Download URL not available"
                                }
                                
                                Write-Host " [$($i+1)] $fileName" -ForegroundColor Green
                                Write-Host "     Size: $fileSize" -ForegroundColor Cyan
                                Write-Host "     Status: $downloadStatus" -ForegroundColor Cyan
                                Write-Host "----------------------------------------" -ForegroundColor Gray
                            }
                            
                            Write-Host
                            $downloadChoice = Read-Host "Enter attachment number to download [1-$($attachments.Count)] or 0 to cancel"
                            
                            if ($downloadChoice -ne "0" -and $downloadChoice -match '^\d+$') {
                                $index = [int]$downloadChoice - 1
                                if ($index -ge 0 -and $index -lt $attachments.Count) {
                                    if ($null -ne $attachments[$index].downloadUrl -and 
                                        -not [string]::IsNullOrEmpty($attachments[$index].downloadUrl)) {
                                        
                                        $fileName = if ($null -ne $attachments[$index].fileName) {
                                            $attachments[$index].fileName
                                        }
                                        else {
                                            "Export_$($export.id)_Attachment_$($index+1)"
                                        }
                                        
                                        # Create folder if it doesn't exist
                                        if (-not (Test-Path -Path $DownloadPath)) {
                                            Write-LogEntry -LogName $logPath -LogEntryText "Creating download folder: $DownloadPath" -LogLevel "INFO"
                                            New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
                                        }
                                        
                                        $outputFilePath = "$DownloadPath\$fileName"
                                        
                                        Write-Host "Downloading $fileName to $outputFilePath..." -ForegroundColor Yellow
                                        
                                        try {
                                            Write-Host "Attempting specialized M365 Compliance download..." -ForegroundColor Yellow
                                            
                                            # Check the URL patterns to determine the right download method
                                            $isPurviewUrl = $attachments[$index].downloadUrl -like "*purview*" -or 
                                            $attachments[$index].downloadUrl -like "*purviewcases*" -or
                                            $attachments[$index].downloadUrl -like "*getAction*"
                                   
                                            $isProxyUrl = $attachments[$index].downloadUrl -like "*proxyservice.ediscovery*" -or 
                                            $attachments[$index].downloadUrl -like "*exportaedblobFileResult*"
                                            
                                            $downloadSuccess = $false
                                            
                                            # Purview URLs need special handling to extract the real download URL
                                            if ($isPurviewUrl) {
                                                Write-Host "Detected Microsoft Purview URL. Using specialized handling..." -ForegroundColor Yellow
                                                $downloadSuccess = ConvertTo-M365PurviewDownload -DownloadUrl $attachments[$index].downloadUrl -OutputFilePath $outputFilePath
                                            }
                                            # Proxy URLs can be downloaded directly with our enhanced download function
                                            elseif ($isProxyUrl) {
                                                Write-Host "Detected Microsoft 365 Compliance Export URL. Using specialized download method..." -ForegroundColor Yellow
                                                $downloadSuccess = Save-M365ComplianceFile -DownloadUrl $attachments[$index].downloadUrl -OutputFilePath $outputFilePath
                                            }
                                            else {
                                                # Standard download method for non-M365 exports
                                                # Get auth token headers for binary content
                                                $headers = @{
                                                    "Accept" = "application/octet-stream"
                                                }
                                                
                                                Write-Host "Starting standard binary download..." -ForegroundColor Yellow
                                                
                                                # Try multiple methods in sequence
                                                try {
                                                    # Method 1: WebClient
                                                    $webClient = New-Object System.Net.WebClient
                                                    foreach ($key in $headers.Keys) {
                                                        $webClient.Headers.Add($key, $headers[$key])
                                                    }
                                                    
                                                    $ProgressPreference = 'SilentlyContinue'
                                                    $webClient.DownloadFile($attachments[$index].downloadUrl, $outputFilePath)
                                                    $ProgressPreference = 'Continue'
                                                    $downloadSuccess = $true
                                                }
                                                catch {
                                                    Write-Host "WebClient download failed: $_" -ForegroundColor Yellow
                                                    
                                                    # Method 2: BITS Transfer
                                                    try {
                                                        if (Get-Module -ListAvailable -Name BitsTransfer) {
                                                            Import-Module BitsTransfer
                                                            Start-BitsTransfer -Source $attachments[$index].downloadUrl -Destination $outputFilePath -DisplayName "Downloading attachment" -Priority High
                                                            $downloadSuccess = $true
                                                        }
                                                        else {
                                                            # Method 3: Invoke-WebRequest
                                                            $ProgressPreference = 'SilentlyContinue'
                                                            Invoke-WebRequest -Uri $attachments[$index].downloadUrl -Headers $headers -OutFile $outputFilePath -Method Get -UseBasicParsing
                                                            $ProgressPreference = 'Continue'
                                                            $downloadSuccess = $true
                                                        }
                                                    }
                                                    catch {
                                                        Write-Host "All download methods failed: $_" -ForegroundColor Red
                                                        $downloadSuccess = $false
                                                    }
                                                }
                                            }
                                            
                                            # Validate the downloaded file
                                            if (Test-Path -Path $outputFilePath) {
                                                $fileInfo = Get-Item -Path $outputFilePath
                                                if ($fileInfo.Length -gt 0) {
                                                    # Special case for Report files - they use .zip extension but aren't true ZIP files
                                                    if ($outputFilePath -like "*Reports-*") {
                                                        Write-LogEntry -LogName $logPath -LogEntryText "Microsoft 365 Report file successfully downloaded: $outputFilePath" -LogLevel "INFO"
                                                        Write-Host "? Microsoft 365 Report file downloaded successfully: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                                        Write-Host "Note: These Report files use .zip extension but are actually HTML/XML files." -ForegroundColor Yellow
                                                        
                                                        return @{
                                                            Success      = $true
                                                            FilePath     = $outputFilePath
                                                            DownloadDate = Get-Date
                                                            IsReport     = $true
                                                        }
                                                    }
                                                    # Check if it's a ZIP file and validate it
                                                    elseif ($outputFilePath.EndsWith('.zip')) {
                                                        # Use the enhanced validation with retry logic
                                                        if (Test-ValidZipFile -FilePath $outputFilePath -RetryOnFailure -MaxRetries 3 -RetryDelay 2) {
                                                            Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded and validated: $outputFilePath" -LogLevel "INFO"
                                                            Write-Host "? Download successful and ZIP file validated: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                                            
                                                            return @{
                                                                Success      = $true
                                                                FilePath     = $outputFilePath
                                                                DownloadDate = Get-Date
                                                                IsValidZip   = $true
                                                            }
                                                        }
                                                        else {
                                                            Write-LogEntry -LogName $logPath -LogEntryText "Downloaded file is not a valid ZIP file after multiple attempts: $outputFilePath" -LogLevel "ERROR"
                                                            Write-Host "?? Downloaded file is not a valid ZIP file after validation attempts." -ForegroundColor Red
                                                            
                                                            return @{
                                                                Success    = $false
                                                                FilePath   = $outputFilePath
                                                                Error      = "Downloaded file is not a valid ZIP file"
                                                                IsValidZip = $false
                                                            }
                                                        }
                                                    }
                                                    else {
                                                        Write-LogEntry -LogName $logPath -LogEntryText "File successfully downloaded: $outputFilePath" -LogLevel "INFO"
                                                        Write-Host "? Download successful: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
                                                            
                                                        return @{
                                                            Success      = $true
                                                            FilePath     = $outputFilePath
                                                            DownloadDate = Get-Date
                                                        }
                                                    }
                                                }
                                                else {
                                                    Write-LogEntry -LogName $logPath -LogEntryText "Downloaded file is empty (0 bytes): $outputFilePath" -LogLevel "ERROR"
                                                    Write-Host "?? Downloaded file is empty (0 bytes)." -ForegroundColor Red
                                                    
                                                    return @{
                                                        Success  = $false
                                                        FilePath = $outputFilePath
                                                        Error    = "Downloaded file is empty (0 bytes)"
                                                    }
                                                }
                                            }
                                            else {
                                                Write-LogEntry -LogName $logPath -LogEntryText "Failed to download file. Output path not created: $outputFilePath" -LogLevel "ERROR"
                                                Write-Host "? Failed to download file. Output path not created." -ForegroundColor Red
                                                
                                                return @{
                                                    Success  = $false
                                                    FilePath = $outputFilePath
                                                    Error    = "Download failed - file not created"
                                                }
                                            }
                                        }
                                        catch {
                                            Write-LogEntry -LogName $logPath -LogEntryText "Error downloading file: $_" -LogLevel "ERROR"
                                            Write-Host "Error downloading file: $_" -ForegroundColor Red
                                            
                                            return @{
                                                Success = $false
                                                Error   = $_
                                            }
                                        }
                                    }
                                    else {
                                        Write-Host "Download URL not available for this attachment." -ForegroundColor Red
                                    }
                                }
                                else {
                                    Write-Host "Invalid attachment number." -ForegroundColor Red
                                }
                            }
                        }
                    }
                }
                
                Write-Host "Press any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "4" {
                # Download export
                $case = Select-Case
                if ($case) {
                    $export = Select-Export -CaseId $case.id
                    if ($export) {
                        $downloadResult = Save-ExportResults -CaseId $case.id -ExportId $export.id -SearchName "Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                        
                        if ($downloadResult -and $downloadResult.Success) {
                            Write-Host "Download completed successfully." -ForegroundColor Green
                            Write-Host "File saved to: $($downloadResult.FilePath)" -ForegroundColor Cyan
                        }
                    }
                }
                
                Write-Host "Press any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "5" {
                # List cases
                $cases = Get-eDiscoveryCases
                
                if ($cases.Count -eq 0) {
                    Write-Host "No eDiscovery cases found." -ForegroundColor Red
                }
                else {
                    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                    Write-Host "�              Available Cases                    �" -ForegroundColor Cyan
                    Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                    Write-Host
                    
                    foreach ($case in $cases) {
                        Write-Host "Name: $($case.displayName)" -ForegroundColor Green
                        Write-Host "ID: $($case.id)" -ForegroundColor Cyan
                        Write-Host "Status: $($case.status)" -ForegroundColor Cyan
                        Write-Host "Created: $($case.createdDateTime)" -ForegroundColor Cyan
                        Write-Host "----------------------------------------" -ForegroundColor Gray
                    }
                }
                
                Write-Host "Press any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "6" {
                # List searches in a case
                try {
                    $case = Select-Case
                    if ($case) {
                        Write-Host "Retrieving searches for case: $($case.displayName)" -ForegroundColor Yellow
                            
                        $searches = Get-eDiscoverySearches -CaseId $case.id
                        if ($searches.Count -eq 0) {
                            Write-Host "No searches found in this case." -ForegroundColor Red
                        }
                        else {
                            Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                            Write-Host "�              Available Searches                 �" -ForegroundColor Cyan
                            Write-Host "+-------------------------------------------------+" -ForegroundColor Cyan
                            Write-Host
                            
                            foreach ($search in $searches) {
                                try {
                                    $statusDisplay = "Available"
                                    $statusColor = "Cyan"
                                                
                                    if ($search.PSObject.Properties.Name -contains "ReadinessStatus") {
                                        $statusDisplay = $search.ReadinessStatus
                                                    
                                        # Color code the status based on readiness
                                        if ($statusDisplay -like "*Export Available*" -or $statusDisplay -like "*Ready for Export*" -or $statusDisplay -like "*Likely Ready*") {
                                            $statusColor = "Green"
                                        }
                                        elseif ($statusDisplay -like "*Processing*" -or $statusDisplay -like "*Running*" -or $statusDisplay -like "*In Progress*") {
                                            $statusColor = "Yellow"
                                        }
                                        elseif ($statusDisplay -like "*Failed*") {
                                            $statusColor = "Red"
                                        }
                                        elseif ($statusDisplay -like "*Unknown*" -or $statusDisplay -like "*Check Manually*") {
                                            $statusColor = "Gray"
                                        }
                                        else {
                                            # For other status messages, use default color
                                            $statusColor = "Cyan"
                                        }
                                    }
                                    else {
                                        # Fallback status checking (legacy)
                                        if ($search.PSObject.Properties.Name -contains "status") {
                                            $statusDisplay = $search.status
                                        }
                                        elseif ($search.PSObject.Properties.Name -contains "lastEstimateStatisticsOperation" -and 
                                            $search.lastEstimateStatisticsOperation -and 
                                            $search.lastEstimateStatisticsOperation.status) {
                                            $statusDisplay = "Estimate: $($search.lastEstimateStatisticsOperation.status)"
                                        }
                                        elseif ($search.PSObject.Properties.Name -contains "lastModifiedDateTime") {
                                            $statusDisplay = "Modified: $($search.lastModifiedDateTime)"
                                        }
                                    }
                                                
                                    Write-Host "Name: $($search.displayName)" -ForegroundColor Green
                                    Write-Host "ID: $($search.id)" -ForegroundColor Cyan
                                    Write-Host "Status: $statusDisplay" -ForegroundColor $statusColor
                                    Write-Host "Created: $($search.createdDateTime)" -ForegroundColor Cyan
                                                
                                    # Show additional debug info if available
                                    if ($search.PSObject.Properties.Name -contains "description" -and 
                                        ![string]::IsNullOrEmpty($search.description)) {
                                        Write-Host "Description: $($search.description)" -ForegroundColor Gray
                                    }
                                                
                                    Write-Host "----------------------------------------" -ForegroundColor Gray
                                }
                                catch {
                                    Write-LogEntry -LogName $logPath -LogEntryText "Error processing search details: $_" -LogLevel "ERROR"
                                    Write-Host "Error processing search details: $_" -ForegroundColor Red
                                    Write-Host "----------------------------------------" -ForegroundColor Gray
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-LogEntry -LogName $logPath -LogEntryText "Error listing searches: $_" -LogLevel "ERROR"
                    Write-Host "Error listing searches: $_" -ForegroundColor Red
                }
                finally {
                    Write-Host "Press any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
            }
            default {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}
else {
    # Handle command line operations
    try {
        switch ($Operation) {
            "create" {
                try {
                    # Create new case and search
                    $result = New-CalendarExportCase -UserEmail $EmailAddress -ContentTypeFilter $ContentType
                
                    if ($result) {
                        Write-Host "Case and search created successfully." -ForegroundColor Green
                        Write-Host "Case ID: $($result.CaseId)" -ForegroundColor Cyan
                        Write-Host "Search ID: $($result.SearchId)" -ForegroundColor Cyan
                    }
                    else {
                        throw "Failed to create case and search"
                    }
                }
                catch {
                    Write-LogEntry -LogName $logPath -LogEntryText "Error creating case: $_" -LogLevel "ERROR"
                    Write-Host "Error creating case: $_" -ForegroundColor Red
                }
            }
            "export" {
                try {
                    # Handle export operation
                    if ($CaseId -and $SearchId) {
                        # Start a new export
                        $userSearchName = ($EmailAddress.split("@")[0]).replace(".", "_")
                        $searchName = ($userSearchName) + '_Mailbox'
                    
                        $exportResult = Export-CalendarResults -CaseId $CaseId -SearchId $SearchId -SearchName $searchName
                    
                        if ($exportResult -and $exportResult.ExportId) {
                            Write-Host "Export started successfully." -ForegroundColor Green
                            Write-Host "Export ID: $($exportResult.ExportId)" -ForegroundColor Cyan
                        }
                        else {
                            throw "Failed to start export"
                        }
                    }
                    elseif ($CaseId -and $ExportId) {
                        try {
                            # Check export status
                            $exportStatus = Get-ExportStatus -CaseId $CaseId -ExportId $ExportId
                                
                            Write-Host "Export Status: $($exportStatus.Status)" -ForegroundColor Cyan
                            Write-Host "Created By: $($exportStatus.CreatedBy)" -ForegroundColor Cyan
                            Write-Host "Created Date: $($exportStatus.CreatedDate)" -ForegroundColor Cyan
                                
                            # Wait for export completion if requested
                            if ($exportStatus.Status -ne "succeeded" -and $exportStatus.Status -ne "partiallySucceeded") {
                                $wait = Read-Host "Export is not ready. Do you want to wait for completion? (y/n)"
                                if ($wait -eq "y" -or $wait -eq "Y") {
                                    $completion = Wait-ForExportCompletion -CaseId $CaseId -ExportId $ExportId
                                    if ($completion) {
                                        Write-Host "Export completed with status: $($completion.Status)" -ForegroundColor Cyan
                                        Write-Host "Completed at: $($completion.CompletedTime)" -ForegroundColor Cyan
                                    }
                                    else {
                                        Write-Host "Failed to monitor export completion" -ForegroundColor Red
                                    }
                                }
                            }
                        }
                        catch {
                            Write-LogEntry -LogName $logPath -LogEntryText "Error checking export status: $_" -LogLevel "ERROR"
                            Write-Host "Error checking export status: $_" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "For export operation, you must provide either:" -ForegroundColor Red
                        Write-Host "- Both CaseId and SearchId to start a new export" -ForegroundColor Red
                        Write-Host "- Both CaseId and ExportId to check export status" -ForegroundColor Red
                    }
                }
                catch {
                    Write-LogEntry -LogName $logPath -LogEntryText "Error during export operation: $_" -LogLevel "ERROR"
                    Write-Host "Error during export operation: $_" -ForegroundColor Red
                }
            }
            "download" {
                try {
                    # Download export
                    if ($CaseId -and $ExportId) {
                        $userSearchName = ($EmailAddress.split("@")[0]).replace(".", "_")
                        $searchName = ($userSearchName) + '_Mailbox'
                        
                        $downloadResult = Save-ExportResults -CaseId $CaseId -ExportId $ExportId -SearchName $searchName -OutputPath $DownloadPath
                        
                        if ($downloadResult -and $downloadResult.Success) {
                            Write-Host "Download completed successfully." -ForegroundColor Green
                            Write-Host "File saved to: $($downloadResult.FilePath)" -ForegroundColor Cyan
                        }
                        else {
                            $errorMsg = if ($downloadResult.Error) { $downloadResult.Error } else { "Download failed" }
                            throw $errorMsg
                        }
                    }
                    else {
                        Write-Host "For download operation, you must provide both CaseId and ExportId" -ForegroundColor Red
                    }
                }
                catch {
                    Write-LogEntry -LogName $logPath -LogEntryText "Error during download: $_" -LogLevel "ERROR"
                    Write-Host "Error during download: $_" -ForegroundColor Red
                }
            }
            default {
                Write-Host "Unknown operation: $Operation" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error in command-line operation: $_" -LogLevel "ERROR"
        Write-Host "Error in operation: $_" -ForegroundColor Red
    }
}

Write-LogEntry -LogName $logPath -LogEntryText "Script completed." -LogLevel "INFO"
Write-Host "Script execution completed. Check the log file for details: $logPath" -ForegroundColor Green
