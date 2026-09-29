#requires -Version 5.1

<#
.SYNOPSIS
    Exchange Online Public Folder -> Mailbox root-folder-only migration engine using raw EWS SOAP.

.DESCRIPTION
    Copy-only migration utility for an Exchange Online Public Folder.

    ONLY the selected source folder itself is processed.
    Child/subfolders are deliberately ignored.

    Items found directly inside the selected source folder are migrated
    according to their Outlook/MAPI ItemClass.

    Currently migrates:
      - Mail
      - Public Folder posts
      - Calendar appointments
      - Contacts
      - Tasks
      - Sticky Notes
      - Journal entries

    Detects, but deliberately does not yet migrate:
      - Meeting requests
      - Meeting cancellations
      - Meeting responses
      - Distribution lists / contact groups
      - InfoPath items
      - Documents
      - Reports
      - Sharing messages
      - Unknown item classes

    IMPORTANT:
      - The source is strictly read-only.
      - No source CreateItem, UpdateItem, DeleteItem, MoveItem or CopyItem calls exist.
      - Destination duplicate checks use only native Exchange/EWS item properties.
      - No application-specific marker/property is written to migrated items.
      - No destination folders are created, reused, or renamed.
      - Child/subfolders are NEVER scanned.

    Authentication:
      - OAuth 2.0 client_credentials against Microsoft Entra ID.
      - EWS OAuth requests use ExchangeImpersonation.
      - ExchangeOnlineManagement is used only when the Public Folder mailbox
        needs to be discovered automatically.

    Configuration:
      The script loads .\config.txt and expects:
        $TenantId
        $ClientId
        $DestinationMailbox

      Optional:
        $SourcePublicFolderMailbox
        $EwsEndpoint
        $CalendarViewStart
        $CalendarViewEnd

    Example config.txt:

        $TenantId = "00000000-0000-0000-0000-000000000000"
        $ClientId = "00000000-0000-0000-0000-000000000000"
        $DestinationMailbox = "mailbox1@testmaxit.onmicrosoft.com"
        $SourcePublicFolderMailbox = "PublicFolder_94bed040@testmaxit.onmicrosoft.com"
        $EwsEndpoint = "https://outlook.office365.com/EWS/Exchange.asmx"
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# STEP 01: CONFIGURATION
# ============================================================

$TenantId                  = $null
$ClientId                  = $null
$DestinationMailbox        = $null
$SourcePublicFolderMailbox = $null
$EwsEndpoint               = 'https://outlook.office365.com/EWS/Exchange.asmx'
$CalendarViewStart         = $null
$CalendarViewEnd           = $null

if (-not (Test-Path '.\config.txt')) {
    throw 'config.txt was not found in the current directory.'
}

$config = Get-Content '.\config.txt' -Raw
Invoke-Expression $config

if ([string]::IsNullOrWhiteSpace($TenantId)) {
    throw 'TenantId is missing from config.txt.'
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw 'ClientId is missing from config.txt.'
}

if ([string]::IsNullOrWhiteSpace($DestinationMailbox)) {
    throw 'DestinationMailbox is missing from config.txt.'
}

if ([string]::IsNullOrWhiteSpace($EwsEndpoint)) {
    $EwsEndpoint = 'https://outlook.office365.com/EWS/Exchange.asmx'
}

Write-Host '========== STEP 01: CONFIGURATION =========='
Write-Host "Destination mailbox: $DestinationMailbox"
Write-Host "EWS endpoint        : $EwsEndpoint"

if ($SourcePublicFolderMailbox) {
    Write-Host "Source PF mailbox   : $SourcePublicFolderMailbox"
} else {
    Write-Host 'Source PF mailbox   : automatic discovery'
}

# ============================================================
# SCRIPT STATE
# ============================================================

$SourceHeaders            = @{}
$DestinationHeaders       = @{}
$SourceImpersonation      = $null
$DestinationImpersonation = $null
$Token                    = $null
$SecretText               = $null

$Stats = [ordered]@{
    FoldersDiscovered = 0

    MailDetected  = 0
    MailCopied    = 0
    MailSkipped   = 0
    MailFailed    = 0

    CalendarDetected = 0
    CalendarCopied   = 0
    CalendarSkipped  = 0
    CalendarFailed   = 0

    ContactDetected = 0
    ContactCopied   = 0
    ContactSkipped  = 0
    ContactFailed   = 0

    TaskDetected = 0
    TaskCopied   = 0
    TaskSkipped  = 0
    TaskFailed   = 0

    NoteDetected = 0
    NoteCopied   = 0
    NoteSkipped  = 0
    NoteFailed   = 0

    JournalDetected = 0
    JournalCopied   = 0
    JournalSkipped  = 0
    JournalFailed   = 0

    DetectionOnly = 0
}

$AllFolders         = @()
$MailFolders        = @()
$CalendarFolders    = @()
$OtherTypedFolders  = @()

$AllNormalItems     = @()
$AllCalendarItems   = @()
$AllOtherItems      = @()

# NEW:
# Unified collection of typed items discovered from BOTH:
#   - IPF.Note folders
#   - Other typed Public Folder containers
#
# This is the important correction.
$AllTypedItems      = @()

$DetectionOnlyItems = @()

$DestinationInboxId    = $null
$DestinationCalendarId = $null
$DestinationContactsId = $null
$DestinationTasksId    = $null
$DestinationNotesId    = $null
$DestinationJournalId  = $null

# ============================================================
# STEP 02: EXCHANGE ONLINE CONNECTION
# ============================================================

Write-Host '========== STEP 02: EXCHANGE ONLINE CONNECTION =========='

function Test-ExchangeOnlineConnection {
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-ExchangeOnlineConnection {

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'ExchangeOnlineManagement is not installed. Install-Module ExchangeOnlineManagement or set SourcePublicFolderMailbox in config.txt.'
    }

    if (-not (Get-Module -Name ExchangeOnlineManagement)) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    }

    if (-not (Test-ExchangeOnlineConnection)) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
}

if (-not $SourcePublicFolderMailbox) {

    Ensure-ExchangeOnlineConnection

    $rootPf = Get-PublicFolderMailbox -ErrorAction Stop |
        Where-Object { $_.IsRootPublicFolderMailbox -eq $true } |
        Select-Object -First 1

    if (-not $rootPf) {
        throw 'Could not discover the root Public Folder mailbox.'
    }

    $SourcePublicFolderMailbox = [string]$rootPf.PrimarySmtpAddress

    if ([string]::IsNullOrWhiteSpace($SourcePublicFolderMailbox)) {
        throw 'The root Public Folder mailbox has no usable PrimarySmtpAddress.'
    }
}

Write-Host "Source Public Folder mailbox: $SourcePublicFolderMailbox"

# ============================================================
# STEP 03: AUTHENTICATION
# ============================================================

Write-Host '========== STEP 03: AUTHENTICATION =========='

$Secret = Read-Host 'Client Secret' -AsSecureString

$SecureBSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)

try {
    $SecretText = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($SecureBSTR)
}
finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($SecureBSTR)
}

$TokenResponse = Invoke-RestMethod `
    "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Method Post `
    -Body @{
        client_id     = $ClientId
        client_secret = $SecretText
        scope         = 'https://outlook.office365.com/.default'
        grant_type    = 'client_credentials'
    } `
    -ErrorAction Stop

if (-not $TokenResponse.access_token) {
    throw 'OAuth token response did not contain an access_token.'
}

$Token = $TokenResponse.access_token
$SecretText = $null

Write-Host 'OAuth authentication successful.'

# ============================================================
# STEP 04: EWS SETUP
# ============================================================

Write-Host '========== STEP 04: EWS SETUP =========='

$SourceHeaders = @{
    Authorization           = "Bearer $Token"
    'X-AnchorMailbox'       = $SourcePublicFolderMailbox
    'X-PublicFolderMailbox' = $SourcePublicFolderMailbox
}

$DestinationHeaders = @{
    Authorization     = "Bearer $Token"
    'X-AnchorMailbox' = $DestinationMailbox
}

$SourceImpersonation = @"
<t:ExchangeImpersonation>
    <t:ConnectingSID>
        <t:SmtpAddress>$([System.Security.SecurityElement]::Escape($SourcePublicFolderMailbox))</t:SmtpAddress>
    </t:ConnectingSID>
</t:ExchangeImpersonation>
"@

$DestinationImpersonation = @"
<t:ExchangeImpersonation>
    <t:ConnectingSID>
        <t:SmtpAddress>$([System.Security.SecurityElement]::Escape($DestinationMailbox))</t:SmtpAddress>
    </t:ConnectingSID>
</t:ExchangeImpersonation>
"@

$NamespaceManager = New-Object System.Xml.XmlNamespaceManager(
    (New-Object System.Xml.NameTable)
)

$NamespaceManager.AddNamespace(
    's',
    'http://schemas.xmlsoap.org/soap/envelope/'
)

$NamespaceManager.AddNamespace(
    'm',
    'http://schemas.microsoft.com/exchange/services/2006/messages'
)

$NamespaceManager.AddNamespace(
    't',
    'http://schemas.microsoft.com/exchange/services/2006/types'
)

function ConvertTo-XmlSafe {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ''
    }

    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-NodeText {
    param(
        [System.Xml.XmlNode]$Node,
        [string]$XPath
    )

    if (-not $Node) {
        return $null
    }

    $n = $Node.SelectSingleNode($XPath, $NamespaceManager)

    if ($n) {
        return $n.InnerText
    }

    return $null
}

function Convert-EwsBodyToPlainText {
    param(
        [string]$Body
    )

    if ([string]::IsNullOrWhiteSpace($Body)) {
        return ''
    }

    $Text = [System.Net.WebUtility]::HtmlDecode($Body)

    $Text = [regex]::Replace(
        $Text,
        '(?is)<\s*(br\s*/?|/p|/div|/li|/tr|/h[1-6])\s*>',
        "`r`n"
    )

    $Text = [regex]::Replace(
        $Text,
        '(?is)<\s*li\b[^>]*>',
        '- '
    )

    $Text = [regex]::Replace(
        $Text,
        '(?is)<[^>]+>',
        ''
    )

    $Text = [System.Net.WebUtility]::HtmlDecode($Text)

    $Text = $Text.Replace([char]0x00A0, ' ')

    $Text = [regex]::Replace(
        $Text,
        '[ \t]+',
        ' '
    )

    $Text = [regex]::Replace(
        $Text,
        '(\r?\n){3,}',
        "`r`n`r`n"
    )

    return $Text.Trim()
}

function New-EwsEnvelope {
    param(
        [string]$Body,

        [ValidateSet('Source','Destination')]
        [string]$Target
    )

    $Impersonation = if ($Target -eq 'Source') {
        $SourceImpersonation
    }
    else {
        $DestinationImpersonation
    }

    return @"
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <s:Header>
        <t:RequestServerVersion Version="Exchange2016"/>
        $Impersonation
    </s:Header>
    <s:Body>
        $Body
    </s:Body>
</s:Envelope>
"@
}

function Invoke-EwsRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Body,

        [Parameter(Mandatory)]
        [ValidateSet('Source','Destination')]
        [string]$Target,

        [Parameter(Mandatory)]
        [string]$OperationName
    )

    $Headers = if ($Target -eq 'Source') {
        $SourceHeaders
    }
    else {
        $DestinationHeaders
    }

    $Soap = New-EwsEnvelope -Body $Body -Target $Target

    $MaxAttempts = 3

    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {

        try {

            $Response = Invoke-WebRequest `
                -Uri $EwsEndpoint `
                -Headers $Headers `
                -ContentType 'text/xml; charset=utf-8' `
                -Method Post `
                -Body $Soap `
                -ErrorAction Stop

            [xml]$Xml = $Response.Content

            break
        }
        catch {

            if ($Attempt -eq $MaxAttempts) {

                $ErrorBody = ''

                if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $ErrorBody = $_.ErrorDetails.Message
                }

                if ($ErrorBody) {
                    throw "EWS $OperationName failed after $MaxAttempts attempts: $ErrorBody"
                }

                throw "EWS $OperationName failed after $MaxAttempts attempts: $($_.Exception.Message)"
            }

            Start-Sleep -Seconds ([Math]::Pow(2, $Attempt - 1))
        }
    }

    $Fault = $Xml.SelectSingleNode(
        '//s:Fault',
        $NamespaceManager
    )

    if ($Fault) {

        $FaultText = $Fault.SelectSingleNode(
            '*[local-name()="faultstring"]'
        )

        if ($FaultText) {
            throw "EWS $OperationName SOAP Fault: $($FaultText.InnerText)"
        }

        throw "EWS $OperationName SOAP Fault: $($Fault.InnerText)"
    }

    $ResponseMessages = $Xml.SelectNodes(
        '//*[@ResponseClass]',
        $NamespaceManager
    )

    foreach ($Message in $ResponseMessages) {

        if ($Message.GetAttribute('ResponseClass') -eq 'Error') {

            $ResponseCode = $Message.SelectSingleNode(
                'm:ResponseCode',
                $NamespaceManager
            )

            $MessageText = $Message.SelectSingleNode(
                'm:MessageText',
                $NamespaceManager
            )

            $CodeText = if ($ResponseCode) {
                $ResponseCode.InnerText
            }
            else {
                'Unknown'
            }

            $Text = if ($MessageText) {
                $MessageText.InnerText
            }
            else {
                'Unknown EWS error'
            }

            throw "EWS $OperationName failed: $CodeText - $Text"
        }
    }

    return $Xml
}

# ============================================================
# STEP 05: PUBLIC FOLDER ROOT
# ============================================================

Write-Host '========== STEP 05: PUBLIC FOLDER ROOT =========='

$RootBody = @"
<m:GetFolder>
    <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
    </m:FolderShape>
    <m:FolderIds>
        <t:DistinguishedFolderId Id="publicfoldersroot"/>
    </m:FolderIds>
</m:GetFolder>
"@

$RootXml = Invoke-EwsRequest `
    -Body $RootBody `
    -Target Source `
    -OperationName 'GetFolder(PublicFolderRoot)'

$RootFolderNode = $RootXml.SelectSingleNode(
    '//m:GetFolderResponseMessage/m:Folders/*',
    $NamespaceManager
)

if (-not $RootFolderNode) {
    throw 'Could not resolve the Public Folder root.'
}

Write-Host 'Public Folder root resolved.'

# ============================================================
# STEP 06: FIND PUBLICFOLDER SOURCE ROOT
# ============================================================

Write-Host '========== STEP 06: FIND PUBLICFOLDER SOURCE ROOT =========='

function Get-EwsChildFolders {

    param(
        [Parameter(Mandatory)]
        [string]$ParentXml,

        [Parameter(Mandatory)]
        [ValidateSet('Source','Destination')]
        [string]$Target
    )

    $Results = New-Object System.Collections.Generic.List[object]

    $Offset = 0
    $PageSize = 500

    while ($true) {

        $Body = @"
<m:FindFolder Traversal="Shallow">
    <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="folder:ChildFolderCount"/>
            <t:FieldURI FieldURI="folder:TotalCount"/>
            <t:FieldURI FieldURI="folder:FolderClass"/>
        </t:AdditionalProperties>

    </m:FolderShape>

    <m:IndexedPageFolderView
        MaxEntriesReturned="$PageSize"
        Offset="$Offset"
        BasePoint="Beginning"/>

    <m:ParentFolderIds>
        $ParentXml
    </m:ParentFolderIds>

</m:FindFolder>
"@

        $Xml = Invoke-EwsRequest `
            -Body $Body `
            -Target $Target `
            -OperationName 'FindFolder'

        $Root = $Xml.SelectSingleNode(
            '//m:FindFolderResponseMessage/m:RootFolder',
            $NamespaceManager
        )

        $FolderNodes = $Xml.SelectNodes(
            '//m:FindFolderResponseMessage/m:RootFolder/t:Folders/*',
            $NamespaceManager
        )

        foreach ($Node in $FolderNodes) {

            $IdNode = $Node.SelectSingleNode(
                't:FolderId',
                $NamespaceManager
            )

            if (-not $IdNode) {
                continue
            }

            $Results.Add(
                [PSCustomObject]@{
                    DisplayName      = Get-NodeText $Node 't:DisplayName'
                    FolderId         = $IdNode.GetAttribute('Id')
                    ChangeKey        = $IdNode.GetAttribute('ChangeKey')
                    FolderClass      = Get-NodeText $Node 't:FolderClass'
                    FolderType       = $Node.LocalName
                    ChildFolderCount = [int](Get-NodeText $Node 't:ChildFolderCount')
                    TotalCount       = [int](Get-NodeText $Node 't:TotalCount')
                }
            )
        }

        $Last = $true

        if ($Root) {

            $Attr = $Root.GetAttribute(
                'IncludesLastItemInRange'
            )

            if ($Attr) {
                $Last = [bool]::Parse($Attr)
            }
        }

        if ($Last) {
            break
        }

        $Offset += $PageSize
    }

    return $Results
}

function Resolve-SourceFolder {

    param(
        [string]$Path
    )

    if (
        [string]::IsNullOrWhiteSpace($Path) -or
        $Path -eq '\'
    ) {

        $Children = Get-EwsChildFolders `
            -ParentXml '<t:DistinguishedFolderId Id="publicfoldersroot"/>' `
            -Target Source

        if ($Children.Count -eq 1) {
            return $Children[0]
        }

        if ($Children.Count -eq 0) {
            throw 'No top-level Public Folder was found.'
        }

        throw "Multiple top-level Public Folders were found: $(($Children | ForEach-Object DisplayName) -join ', '). Set SourcePath explicitly."
    }

    $Segments = $Path.Trim('\') -split '\\'

    $ParentXml =
        '<t:DistinguishedFolderId Id="publicfoldersroot"/>'

    $Current = $null

    foreach ($Segment in $Segments) {

        $Children = Get-EwsChildFolders `
            -ParentXml $ParentXml `
            -Target Source

        $Current = $Children |
            Where-Object {
                $_.DisplayName -eq $Segment
            } |
            Select-Object -First 1

        if (-not $Current) {
            throw "Source folder '$Path' could not be resolved. Missing segment '$Segment'."
        }

        $ParentXml =
            "<t:FolderId Id=`"$($Current.FolderId)`" ChangeKey=`"$($Current.ChangeKey)`"/>"
    }

    return $Current
}

$SourcePath = '\PublicFolder'

if (
    Get-Variable `
        -Name SourcePath `
        -Scope Global `
        -ErrorAction SilentlyContinue
) {

    if ($Global:SourcePath) {
        $SourcePath = $Global:SourcePath
    }
}

$SourceRoot = Resolve-SourceFolder -Path $SourcePath

$SourceRootPath = "\$($SourceRoot.DisplayName)"

Write-Host "Source root: $SourceRootPath"

# ============================================================
# STEP 07: FOLDER DISCOVERY (ROOT + TYPED SUBFOLDERS ONLY)
# ============================================================

Write-Host '========== STEP 07: FOLDER DISCOVERY (ROOT + TYPED SUBFOLDERS ONLY) =========='

$TypedFolderClasses = @(
    'IPF.Appointment*',
    'IPF.Contact*',
    'IPF.Task*',
    'IPF.StickyNote*',
    'IPF.Journal*',
    'IPF.Activity*'
)

function Get-EligibleFolders {
    param(
        [Parameter(Mandatory)][PSCustomObject]$RootFolder,
        [Parameter(Mandatory)][string]$RootPath
    )

    $List = New-Object System.Collections.Generic.List[object]

    # 1. Always include the Root Public Folder itself
    $List.Add([PSCustomObject]@{
        DisplayName      = $RootFolder.DisplayName
        FolderId         = $RootFolder.FolderId
        ChangeKey        = $RootFolder.ChangeKey
        FolderClass      = $RootFolder.FolderClass
        FolderType       = $RootFolder.FolderType
        ChildFolderCount = $RootFolder.ChildFolderCount
        TotalCount       = $RootFolder.TotalCount
        Path             = $RootPath
        IsRoot           = $true
    })

    # 2. Inspect direct child folders to discover typed folders
    #    (Calendar, Contacts, Tasks, Notes, Journal)
    if ($RootFolder.ChildFolderCount -gt 0) {
        $ParentXmlString = "<t:FolderId Id=`"$($RootFolder.FolderId)`" ChangeKey=`"$($RootFolder.ChangeKey)`"/>"

        $Children = Get-EwsChildFolders -ParentXml $ParentXmlString -Target Source

        foreach ($Child in $Children) {
            $IsTyped = $false

            foreach ($Pattern in $TypedFolderClasses) {
                if ($Child.FolderClass -like $Pattern) {
                    $IsTyped = $true
                    break
                }
            }

            if ($IsTyped) {
                # Include typed subfolder
                $List.Add([PSCustomObject]@{
                    DisplayName      = $Child.DisplayName
                    FolderId         = $Child.FolderId
                    ChangeKey        = $Child.ChangeKey
                    FolderClass      = $Child.FolderClass
                    FolderType       = $Child.FolderType
                    ChildFolderCount = $Child.ChildFolderCount
                    TotalCount       = $Child.TotalCount
                    Path             = "$RootPath\$($Child.DisplayName)"
                    IsRoot           = $false
                })

                Write-Host "INCLUDED TYPED SUBFOLDER: $($Child.DisplayName) [$($Child.FolderClass)]"
            }
            else {
                # Exclude normal mail subfolders
                Write-Host "EXCLUDED NORMAL SUBFOLDER: $($Child.DisplayName) [$($Child.FolderClass)]"
            }
        }
    }

    return $List
}

$AllFolders = @(Get-EligibleFolders -RootFolder $SourceRoot -RootPath $SourceRootPath)
$Stats.FoldersDiscovered = $AllFolders.Count

Write-Host "Source root: $SourceRootPath"
Write-Host "Total eligible folders (root + typed subfolders): $($AllFolders.Count)"

# ============================================================
# STEP 08: SELECTED FOLDER INFORMATION
# ============================================================

Write-Host '========== STEP 08: SELECTED FOLDER INFORMATION =========='

Write-Host "Selected source root    : $SourceRootPath"
Write-Host "Root folder class       : $($SourceRoot.FolderClass)"
Write-Host "Root folder type        : $($SourceRoot.FolderType)"
Write-Host "Root direct item count  : $($SourceRoot.TotalCount)"
Write-Host "Root child folder count : $($SourceRoot.ChildFolderCount)"
Write-Host "Subfolders discovered   : $($AllFolders.Count - 1)"
Write-Host ''
Write-Host 'IMPORTANT: All subfolders are scanned so their contents show up'
Write-Host 'in the discovery/debug output below, but the COPY steps only'
Write-Host "migrate items whose parent folder is the root ($SourceRootPath)."
Write-Host 'Items sitting in any subfolder are detected and logged, but skipped.'

# ============================================================
# STEP 09: ITEM CLASSIFICATION
# ============================================================

Write-Host '========== STEP 09: ITEM CLASSIFICATION =========='

function Get-EwsItemType {

    param(
        [string]$ItemClass
    )

    switch -Wildcard ($ItemClass) {

        'IPM.Schedule.Meeting.Request' {
            return 'MeetingRequest'
        }

        'IPM.Schedule.Meeting.Canceled' {
            return 'MeetingCancellation'
        }

        'IPM.Schedule.Meeting.Resp*' {
            return 'MeetingResponse'
        }

        'IPM.Appointment*' {
            return 'Calendar'
        }

        'IPM.Contact*' {
            return 'Contact'
        }

        'IPM.DistList*' {
            return 'DistributionList'
        }

        'IPM.Task*' {
            return 'Task'
        }

        'IPM.StickyNote*' {
            return 'Note'
        }

        'IPM.Activity*' {
            return 'Journal'
        }

        'IPM.Post*' {
            return 'Post'
        }

        'IPM.Note*' {
            return 'MailOrPost'
        }

        'IPM.InfoPath*' {
            return 'InfoPath'
        }

        'IPM.Document*' {
            return 'Document'
        }

        'IPM.Report*' {
            return 'Report'
        }

        'IPM.Sharing*' {
            return 'Sharing'
        }

        default {
            return 'Unknown'
        }
    }
}

# ============================================================
# STEP 10: DISCOVER ITEMS IN ROOT + ALL SUBFOLDERS
# ============================================================

Write-Host '========== STEP 10: ITEM DISCOVERY (ROOT + SUBFOLDERS) =========='

function Get-EwsFolderItems {

    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Folder
    )

    $Results = New-Object System.Collections.Generic.List[object]
    $Offset = 0
    $PageSize = 500

    while ($true) {

        $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:ItemClass"/>
            <t:FieldURI FieldURI="item:DateTimeSent"/>
            <t:FieldURI FieldURI="item:DateTimeReceived"/>
            <t:FieldURI FieldURI="item:Size"/>
            <t:FieldURI FieldURI="item:HasAttachments"/>
            <t:FieldURI FieldURI="message:InternetMessageId"/>
            <t:FieldURI FieldURI="message:IsRead"/>
            <t:FieldURI FieldURI="message:Sender"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:IndexedPageItemView
        MaxEntriesReturned="$PageSize"
        Offset="$Offset"
        BasePoint="Beginning"/>

    <m:ParentFolderIds>

        <t:FolderId
            Id="$($Folder.FolderId)"
            ChangeKey="$($Folder.ChangeKey)"/>

    </m:ParentFolderIds>

</m:FindItem>
"@

        $Xml = Invoke-EwsRequest `
            -Body $Body `
            -Target Source `
            -OperationName 'FindItem(FolderItems)'

        $Root = $Xml.SelectSingleNode(
            '//m:FindItemResponseMessage/m:RootFolder',
            $NamespaceManager
        )

        $Nodes = $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

        foreach ($Node in $Nodes) {

            $IdNode = $Node.SelectSingleNode('t:ItemId', $NamespaceManager)
            if (-not $IdNode) { continue }

            $Class = Get-NodeText -Node $Node -XPath 't:ItemClass'

            $Results.Add([PSCustomObject]@{
                FolderPath        = $Folder.Path
                FolderId          = $Folder.FolderId
                ItemId            = $IdNode.GetAttribute('Id')
                ChangeKey         = $IdNode.GetAttribute('ChangeKey')
                Subject           = Get-NodeText -Node $Node -XPath 't:Subject'
                ItemClass         = $Class
                ItemType          = Get-EwsItemType -ItemClass $Class
                DateTimeSent      = Get-NodeText -Node $Node -XPath 't:DateTimeSent'
                DateTimeReceived  = Get-NodeText -Node $Node -XPath 't:DateTimeReceived'
                InternetMessageId = Get-NodeText -Node $Node -XPath 't:InternetMessageId'
                Sender            = Get-NodeText -Node $Node -XPath 't:Sender/t:Mailbox/t:EmailAddress'
                IsRead            = if (Get-NodeText -Node $Node -XPath 't:IsRead') {
                    [bool]::Parse((Get-NodeText -Node $Node -XPath 't:IsRead'))
                }
                else {
                    $false
                }
                HasAttachments   = if (Get-NodeText -Node $Node -XPath 't:HasAttachments') {
                    [bool]::Parse((Get-NodeText -Node $Node -XPath 't:HasAttachments'))
                }
                else {
                    $false
                }
            })
        }

        $Last = $true

        if ($Root) {
            $Attr = $Root.GetAttribute('IncludesLastItemInRange')
            if ($Attr) {
                $Last = [bool]::Parse($Attr)
            }
        }

        if ($Last) { break }

        $Offset += $PageSize
    }

    return $Results
}

$SourceItemsList = New-Object System.Collections.Generic.List[object]

foreach ($Folder in $AllFolders) {
    $FolderItems = Get-EwsFolderItems -Folder $Folder

    foreach ($FolderItem in $FolderItems) {
        $SourceItemsList.Add($FolderItem)
    }
}

# Converted via .ToArray() to prevent PowerShell array wrapper interop exception
$AllSourceItems = $SourceItemsList.ToArray()

Write-Host "Total items discovered across root + all subfolders: $($AllSourceItems.Count)"
# ============================================================
# STEP 11: CLASSIFY DISCOVERED ITEMS
# ============================================================

Write-Host '========== STEP 11: DISCOVERED ITEM SUMMARY =========='

# Define items belonging directly to the root folder (Mail/Post scope)
$RootOnlyItems = @(
    $AllSourceItems | Where-Object { $_.FolderPath -eq $SourceRootPath }
)

$AllNormalItems = @(
    $RootOnlyItems |
        Where-Object {
            $_.ItemType -in @(
                'MailOrPost',
                'Post',
                'MeetingRequest',
                'MeetingCancellation',
                'MeetingResponse',
                'DistributionList',
                'InfoPath',
                'Document',
                'Report',
                'Sharing',
                'Unknown'
            )
        }
)

$AllCalendarItems = @(
    $AllSourceItems |
        Where-Object {
            $_.ItemType -eq 'Calendar'
        }
)

$AllTypedItems = @(
    $AllSourceItems |
        Where-Object {
            $_.ItemType -in @(
                'Contact',
                'Task',
                'Note',
                'Journal'
            )
        }
)

$MailPostCount = @(
    $RootOnlyItems |
        Where-Object {
            $_.ItemType -in @(
                'MailOrPost',
                'Post'
            )
        }
).Count

$CalendarCount = @($AllCalendarItems).Count

$ContactCount = @(
    $AllTypedItems | Where-Object { $_.ItemType -eq 'Contact' }
).Count

$TaskCount = @(
    $AllTypedItems | Where-Object { $_.ItemType -eq 'Task' }
).Count

$NoteCount = @(
    $AllTypedItems | Where-Object { $_.ItemType -eq 'Note' }
).Count

$JournalCount = @(
    $AllTypedItems | Where-Object { $_.ItemType -eq 'Journal' }
).Count

Write-Host ''
Write-Host "Mail/Post items (Root Folder) : $MailPostCount"
Write-Host "Calendar items                : $CalendarCount"
Write-Host "Contact items                 : $ContactCount"
Write-Host "Task items                    : $TaskCount"
Write-Host "Sticky Note items             : $NoteCount"
Write-Host "Journal items                 : $JournalCount"
Write-Host "Total processed items         : $($AllSourceItems.Count)"

# ============================================================
# STEP 12: DETECTION-ONLY ITEMS
# ============================================================

Write-Host '========== STEP 12: DETECTION-ONLY ITEMS =========='

$DetectionOnlyItems = @(
    $AllSourceItems |
        Where-Object {
            $_.ItemType -in @(
                'MeetingRequest',
                'MeetingCancellation',
                'MeetingResponse',
                'DistributionList',
                'InfoPath',
                'Document',
                'Report',
                'Sharing',
                'Unknown'
            )
        }
)

$Stats.DetectionOnly =
    $DetectionOnlyItems.Count

# ============================================================
# DEBUG: SHOW EVERY DIRECT ITEM CLASS
# ============================================================

Write-Host ''
Write-Host '========== DEBUG: DIRECT SOURCE ITEMS =========='

foreach ($Item in $AllSourceItems) {
    Write-Host (
        'Folder="{0}" | Subject="{1}" | ItemClass="{2}" | ItemType="{3}"' -f `
        $Item.FolderPath,
        $Item.Subject,
        $Item.ItemClass,
        $Item.ItemType
    )
}

Write-Host ''
Write-Host '========== DEBUG: ITEM CLASS COUNTS =========='

$AllSourceItems |
    Group-Object ItemClass |
    Sort-Object Count -Descending |
    ForEach-Object {
        Write-Host (
            '{0} : {1}' -f `
            $_.Name,
            $_.Count
        )
    }

Write-Host ''
Write-Host '========== DEBUG: ITEM TYPE COUNTS =========='

$AllSourceItems |
    Group-Object ItemType |
    Sort-Object Count -Descending |
    ForEach-Object {
        Write-Host (
            '{0} : {1}' -f `
            $_.Name,
            $_.Count
        )
    }

# ============================================================
# STEP 13: DISCOVERY SUMMARY
# ============================================================

Write-Host '========== STEP 13: DISCOVERY SUMMARY =========='

$MigratableNormalCount = @(
    $RootOnlyItems | Where-Object { $_.ItemType -in @('MailOrPost','Post') }
).Count

$MigratableTypedCount = $AllTypedItems.Count

Write-Host ''
Write-Host "Total discovered items               : $($AllSourceItems.Count)"
Write-Host "Migratable mail/post (root folder)   : $MigratableNormalCount"
Write-Host "Migratable calendar items            : $($AllCalendarItems.Count)"
Write-Host "Migratable typed items               : $MigratableTypedCount"
Write-Host "Detection-only items                 : $($DetectionOnlyItems.Count)"
Write-Host ''
Write-Host "Folders scanned                      : $($AllFolders.Count)"

# ============================================================
# STEP 14: DISCOVERY SUMMARY
# ============================================================

Write-Host '========== STEP 14: DISCOVERY SUMMARY =========='

$DetectionOnlyItems = @(
    @($AllNormalItems) +
    @($AllOtherItems) |
    Where-Object {
        $_.ItemType -in @(
            'MeetingRequest',
            'MeetingCancellation',
            'MeetingResponse',
            'DistributionList',
            'InfoPath',
            'Document',
            'Report',
            'Sharing',
            'Unknown'
        )
    }
)

$Stats.DetectionOnly =
    $DetectionOnlyItems.Count

$MigratableNormalCount = @(
    $AllNormalItems |
    Where-Object {
        $_.ItemType -in @(
            'MailOrPost',
            'Post'
        )
    }
).Count

$MigratableTypedCount =
    $AllTypedItems.Count

Write-Host "Total discovered       : $(($AllNormalItems.Count + $AllCalendarItems.Count + $AllOtherItems.Count))"
Write-Host "Migratable mail/post   : $MigratableNormalCount"
Write-Host "Migratable calendar    : $($AllCalendarItems.Count)"
Write-Host "Migratable typed items : $MigratableTypedCount"
Write-Host "Detection-only items   : $($DetectionOnlyItems.Count)"

# ============================================================
# STEP 15: DESTINATION INBOX
# ============================================================

Write-Host '========== STEP 15: DESTINATION INBOX =========='

function Get-DistinguishedFolderId {

    param(
        [Parameter(Mandatory)]
        [string]$Id
    )

    $Body = @"
<m:GetFolder>
    <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
    </m:FolderShape>

    <m:FolderIds>
        <t:DistinguishedFolderId Id="$Id"/>
    </m:FolderIds>

</m:GetFolder>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName "GetFolder($Id)"

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetFolderResponseMessage/m:Folders/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw "Destination distinguished folder '$Id' could not be resolved."
    }

    $IdNode =
        $Node.SelectSingleNode(
            't:FolderId',
            $NamespaceManager
        )

    return $IdNode.GetAttribute('Id')
}

$DestinationInboxId =
    Get-DistinguishedFolderId -Id 'inbox'

Write-Host "Inbox: $DestinationInboxId"

# ============================================================
# STEP 16: DESTINATION CALENDAR
# ============================================================

Write-Host '========== STEP 16: DESTINATION CALENDAR =========='

$DestinationCalendarId =
    Get-DistinguishedFolderId -Id 'calendar'

Write-Host "Calendar: $DestinationCalendarId"

# ============================================================
# STEP 17: DESTINATION CONTACTS
# ============================================================

Write-Host '========== STEP 17: DESTINATION CONTACTS =========='

$DestinationContactsId =
    Get-DistinguishedFolderId -Id 'contacts'

Write-Host "Contacts: $DestinationContactsId"

# ============================================================
# STEP 18: DESTINATION TASKS
# ============================================================

Write-Host '========== STEP 18: DESTINATION TASKS =========='

$DestinationTasksId =
    Get-DistinguishedFolderId -Id 'tasks'

Write-Host "Tasks: $DestinationTasksId"

# ============================================================
# STEP 19: DESTINATION NOTES
# ============================================================

Write-Host '========== STEP 19: DESTINATION NOTES =========='

$DestinationNotesId =
    Get-DistinguishedFolderId -Id 'notes'

Write-Host "Notes: $DestinationNotesId"

# ============================================================
# STEP 20: DESTINATION JOURNAL
# ============================================================

Write-Host '========== STEP 20: DESTINATION JOURNAL =========='

$DestinationJournalId =
    Get-DistinguishedFolderId -Id 'journal'

Write-Host "Journal: $DestinationJournalId"

# ============================================================
# STEP 21: DESTINATION FOLDER MAPPING
# ============================================================

Write-Host '========== STEP 21: DESTINATION FOLDER MAPPING =========='

Write-Host 'No destination folder hierarchy will be created.'

Write-Host 'All source mail/posts -> Inbox'
Write-Host 'All source calendars -> Calendar'
Write-Host 'All source contacts -> Contacts'
Write-Host 'All source tasks -> Tasks'
Write-Host 'All source notes -> Notes'
Write-Host 'All source journals -> Journal'

# ============================================================
# STEP 22: MAIL DUPLICATE CHECK
# ============================================================

Write-Host '========== STEP 22: MAIL DUPLICATE CHECK =========='

function Find-DestinationMessageByInternetMessageId {

    param(
        [string]$FolderId,
        [string]$InternetMessageId
    )

    if ([string]::IsNullOrWhiteSpace($InternetMessageId)) {
        return $null
    }

    $SafeId =
        ConvertTo-XmlSafe $InternetMessageId

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="message:InternetMessageId"/>
            <t:FieldURI FieldURI="message:IsRead"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:Restriction>
        <t:IsEqualTo>
            <t:FieldURI FieldURI="message:InternetMessageId"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeId"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>
    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$FolderId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'FindItem(MailDuplicate)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        return $null
    }

    $IdNode =
        $Node.SelectSingleNode(
            't:ItemId',
            $NamespaceManager
        )

    return [PSCustomObject]@{
        ItemId    = $IdNode.GetAttribute('Id')
        ChangeKey = $IdNode.GetAttribute('ChangeKey')
    }
}

function Find-DestinationMessageFallback {

    param(
        [string]$FolderId,
        [string]$Subject,
        [string]$DateTimeSent,
        [string]$Sender
    )

    $SafeSubject =
        ConvertTo-XmlSafe $Subject

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:DateTimeSent"/>
            <t:FieldURI FieldURI="message:Sender"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:Restriction>

        <t:IsEqualTo>

            <t:FieldURI FieldURI="item:Subject"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeSubject"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>

    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$FolderId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'FindItem(MailFallbackDuplicate)'

    $Nodes =
        $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

    foreach ($Node in $Nodes) {

        $CandidateSent =
            Get-NodeText $Node 't:DateTimeSent'

        $CandidateSender =
            Get-NodeText $Node 't:Sender/t:Mailbox/t:EmailAddress'

        if (
            $CandidateSent -eq $DateTimeSent -and
            (
                (-not $Sender) -or
                $CandidateSender -eq $Sender
            )
        ) {

            $IdNode =
                $Node.SelectSingleNode(
                    't:ItemId',
                    $NamespaceManager
                )

            return [PSCustomObject]@{
                ItemId    = $IdNode.GetAttribute('Id')
                ChangeKey = $IdNode.GetAttribute('ChangeKey')
            }
        }
    }

    return $null
}

# ============================================================
# STEP 23: GET MAIL MIME
# ============================================================

Write-Host '========== STEP 23: GET MAIL MIME =========='

function Remove-XUnsentMimeHeader {

    param(
        [Parameter(Mandatory)]
        [string]$MimeBase64
    )

    $Bytes =
        [Convert]::FromBase64String(
            $MimeBase64
        )

    $Latin1 =
        [System.Text.Encoding]::GetEncoding(28591)

    $MimeText =
        $Latin1.GetString($Bytes)

    $HeaderEnd =
        $MimeText.IndexOf(
            "`r`n`r`n",
            [System.StringComparison]::Ordinal
        )

    $SeparatorLength = 4

    if ($HeaderEnd -lt 0) {

        $HeaderEnd =
            $MimeText.IndexOf(
                "`n`n",
                [System.StringComparison]::Ordinal
            )

        $SeparatorLength = 2
    }

    if ($HeaderEnd -lt 0) {
        return $MimeBase64
    }

    $HeaderAndSeparator =
        $MimeText.Substring(
            0,
            $HeaderEnd + $SeparatorLength
        )

    $BodyText =
        $MimeText.Substring(
            $HeaderEnd + $SeparatorLength
        )

    $CleanHeader =
        [regex]::Replace(
            $HeaderAndSeparator,
            '(?im)^X-Unsent[ \t]*:[^\r\n]*(?:\r\n|\n)',
            ''
        )

    if ($CleanHeader -eq $HeaderAndSeparator) {
        return $MimeBase64
    }

    $CleanBytes =
        $Latin1.GetBytes(
            $CleanHeader + $BodyText
        )

    return [Convert]::ToBase64String(
        $CleanBytes
    )
}

function Get-EwsMime {

    param(
        [string]$ItemId,
        [string]$ChangeKey
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:IncludeMimeContent>true</t:IncludeMimeContent>
    </m:ItemShape>

    <m:ItemIds>
        <t:ItemId
            Id="$ItemId"
            ChangeKey="$ChangeKey"/>
    </m:ItemIds>

</m:GetItem>
"@

    for ($Attempt = 1; $Attempt -le 3; $Attempt++) {

        try {

            $Xml = Invoke-EwsRequest `
                -Body $Body `
                -Target Source `
                -OperationName 'GetItem(MIME)'

            $MimeNode =
                $Xml.SelectSingleNode(
                    '//m:GetItemResponseMessage/m:Items/*/t:MimeContent',
                    $NamespaceManager
                )

            if (-not $MimeNode) {
                throw 'MimeContent was not returned.'
            }

            $Mime = $MimeNode.InnerText

            if ([string]::IsNullOrWhiteSpace($Mime)) {
                throw 'MimeContent was empty.'
            }

            $Mime =
                Remove-XUnsentMimeHeader `
                    -MimeBase64 $Mime

            [void][Convert]::FromBase64String($Mime)

            return $Mime
        }
        catch {

            if ($Attempt -eq 3) {
                throw
            }

            Start-Sleep `
                -Seconds ([Math]::Pow(2, $Attempt - 1))
        }
    }
}

# ============================================================
# STEP 24: PREPARE MAIL
# ============================================================

Write-Host '========== STEP 24: PREPARE MAIL =========='

function Create-DestinationMessageFromMime {

    param(
        [string]$FolderId,
        [string]$Mime
    )

    $Body = @"
<m:CreateItem MessageDisposition="SaveOnly">

    <m:SavedItemFolderId>
        <t:FolderId Id="$FolderId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:Message>

            <t:MimeContent>$Mime</t:MimeContent>

            <t:ExtendedProperty>
                <t:ExtendedFieldURI
                    PropertyTag="0x0E07"
                    PropertyType="Integer"/>
                <t:Value>1</t:Value>
            </t:ExtendedProperty>

        </t:Message>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Mail)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Mail was created but no destination ItemId was returned.'
    }

    return [PSCustomObject]@{
        ItemId    = $IdNode.GetAttribute('Id')
        ChangeKey = $IdNode.GetAttribute('ChangeKey')
    }
}

function Set-DestinationMessageReadState {

    param(
        [string]$ItemId,
        [string]$ChangeKey,
        [bool]$IsRead
    )

    $ReadText =
        if ($IsRead) {
            'true'
        }
        else {
            'false'
        }

    $Body = @"
<m:UpdateItem
    MessageDisposition="SaveOnly"
    ConflictResolution="AlwaysOverwrite">

    <m:ItemChanges>

        <t:ItemChange>

            <t:ItemId
                Id="$ItemId"
                ChangeKey="$ChangeKey"/>

            <t:Updates>

                <t:SetItemField>

                    <t:FieldURI
                        FieldURI="message:IsRead"/>

                    <t:Message>
                        <t:IsRead>$ReadText</t:IsRead>
                    </t:Message>

                </t:SetItemField>

            </t:Updates>

        </t:ItemChange>

    </m:ItemChanges>

</m:UpdateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'UpdateItem(IsRead)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:UpdateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if ($IdNode) {

        return [PSCustomObject]@{
            ItemId    = $IdNode.GetAttribute('Id')
            ChangeKey = $IdNode.GetAttribute('ChangeKey')
        }
    }

    return [PSCustomObject]@{
        ItemId    = $ItemId
        ChangeKey = $ChangeKey
    }
}

# ============================================================
# STEP 25: CREATE MAIL
# ============================================================

Write-Host '========== STEP 25: CREATE MAIL =========='

function Copy-MailItem {

    param(
        [PSCustomObject]$Item
    )

    $Stats.MailDetected++

    $DestinationFolderId =
        $DestinationInboxId

    $Existing = $null

    if ($Item.InternetMessageId) {

        $Existing =
            Find-DestinationMessageByInternetMessageId `
                -FolderId $DestinationFolderId `
                -InternetMessageId $Item.InternetMessageId
    }
    else {

        $Existing =
            Find-DestinationMessageFallback `
                -FolderId $DestinationFolderId `
                -Subject $Item.Subject `
                -DateTimeSent $Item.DateTimeSent `
                -Sender $Item.Sender
    }

    if ($Existing) {

        $Stats.MailSkipped++

        Write-Host "SKIPPED MAIL: $($Item.Subject)"

        return
    }

    try {

        $Mime =
            Get-EwsMime `
                -ItemId $Item.ItemId `
                -ChangeKey $Item.ChangeKey

        $Created =
            Create-DestinationMessageFromMime `
                -FolderId $DestinationFolderId `
                -Mime $Mime

        $null =
            Set-DestinationMessageReadState `
                -ItemId $Created.ItemId `
                -ChangeKey $Created.ChangeKey `
                -IsRead $Item.IsRead

        $Stats.MailCopied++

        Write-Host "COPIED MAIL: $($Item.Subject)"
    }
    catch {

        $Stats.MailFailed++

        Write-Host "FAILED MAIL: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 26: COPY MAIL
# ============================================================

Write-Host '========== STEP 26: COPY MAIL =========='

foreach (
    $Item in (
        $RootOnlyItems |
        Where-Object { $_.ItemType -in @('MailOrPost','Post') }
    )
) {
    Copy-MailItem -Item $Item
}
# ============================================================
# STEP 27: CALENDAR DUPLICATE CHECK
# ============================================================

Write-Host '========== STEP 27: CALENDAR DUPLICATE CHECK =========='

function Convert-ToDateTimeKey {

    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    try {

        return (
            [DateTimeOffset]::Parse($Value)
        ).ToUniversalTime().ToString('o')
    }
    catch {
        return $Value
    }
}

function Find-DestinationCalendarDuplicate {

    param(
        [PSCustomObject]$Item
    )

    $SafeSubject =
        ConvertTo-XmlSafe $Item.Subject

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="calendar:Start"/>
            <t:FieldURI FieldURI="calendar:End"/>
            <t:FieldURI FieldURI="calendar:Location"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:Restriction>

        <t:IsEqualTo>

            <t:FieldURI FieldURI="item:Subject"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeSubject"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>

    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$DestinationCalendarId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'FindItem(CalendarDuplicate)'

    $Nodes =
        $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

    $SourceStart =
        Convert-ToDateTimeKey $Item.Start

    $SourceEnd =
        Convert-ToDateTimeKey $Item.End

    $SourceLoc =
        if ($Item.Location) {
            $Item.Location
        }
        else {
            ''
        }

    foreach ($Node in $Nodes) {

        $CandidateStart =
            Convert-ToDateTimeKey (
                Get-NodeText $Node 't:Start'
            )

        $CandidateEnd =
            Convert-ToDateTimeKey (
                Get-NodeText $Node 't:End'
            )

        $CandidateLoc =
            if (Get-NodeText $Node 't:Location') {
                Get-NodeText $Node 't:Location'
            }
            else {
                ''
            }

        if (
            $CandidateStart -eq $SourceStart -and
            $CandidateEnd -eq $SourceEnd -and
            $CandidateLoc -eq $SourceLoc
        ) {

            $IdNode =
                $Node.SelectSingleNode(
                    't:ItemId',
                    $NamespaceManager
                )

            return [PSCustomObject]@{
                ItemId    = $IdNode.GetAttribute('Id')
                ChangeKey = $IdNode.GetAttribute('ChangeKey')
            }
        }
    }

    return $null
}

# ============================================================
# STEP 28: GET CALENDAR DETAILS
# ============================================================

Write-Host '========== STEP 28: GET CALENDAR DETAILS =========='

function Get-EwsCalendarDetails {

    param(
        [PSCustomObject]$Item
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>
            <t:FieldURI FieldURI="item:ItemClass"/>
            <t:FieldURI FieldURI="calendar:Start"/>
            <t:FieldURI FieldURI="calendar:End"/>
            <t:FieldURI FieldURI="calendar:Location"/>
            <t:FieldURI FieldURI="calendar:IsAllDayEvent"/>
            <t:FieldURI FieldURI="calendar:IsRecurring"/>
            <t:FieldURI FieldURI="calendar:CalendarItemType"/>
            <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus"/>
            <t:FieldURI FieldURI="calendar:UID"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>
        <t:ItemId
            Id="$($Item.ItemId)"
            ChangeKey="$($Item.ChangeKey)"/>
    </m:ItemIds>

</m:GetItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Source `
        -OperationName 'GetItem(CalendarDetails)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetItemResponseMessage/m:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw 'Calendar GetItem returned no item.'
    }

    return [PSCustomObject]@{
        Subject =
            Get-NodeText $Node 't:Subject'

        Body =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Node 't:Body'
            )

        ItemClass =
            Get-NodeText $Node 't:ItemClass'

        Start =
            Get-NodeText $Node 't:Start'

        End =
            Get-NodeText $Node 't:End'

        Location =
            Get-NodeText $Node 't:Location'

        IsAllDayEvent =
            if (Get-NodeText $Node 't:IsAllDayEvent') {
                [bool]::Parse(
                    (Get-NodeText $Node 't:IsAllDayEvent')
                )
            }
            else {
                $false
            }

        IsRecurring =
            if (Get-NodeText $Node 't:IsRecurring') {
                [bool]::Parse(
                    (Get-NodeText $Node 't:IsRecurring')
                )
            }
            else {
                $false
            }

        CalendarItemType =
            Get-NodeText $Node 't:CalendarItemType'

        LegacyFreeBusy =
            Get-NodeText $Node 't:LegacyFreeBusyStatus'

        UID =
            Get-NodeText $Node 't:UID'
    }
}

# ============================================================
# STEP 29: CREATE CALENDAR
# ============================================================

Write-Host '========== STEP 29: CREATE CALENDAR =========='

function Create-DestinationCalendarItem {

    param(
        [PSCustomObject]$Details
    )

    $SafeSubject =
        ConvertTo-XmlSafe $Details.Subject

    $SafeBody =
        ConvertTo-XmlSafe $Details.Body

    $SafeStart =
        ConvertTo-XmlSafe $Details.Start

    $SafeEnd =
        ConvertTo-XmlSafe $Details.End

    $SafeLoc =
        ConvertTo-XmlSafe $Details.Location

    $Busy =
        if ($Details.LegacyFreeBusy) {
            $Details.LegacyFreeBusy
        }
        else {
            'Busy'
        }

    $Body = @"
<m:CreateItem SendMeetingInvitations="SendToNone">

    <m:SavedItemFolderId>
        <t:FolderId Id="$DestinationCalendarId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:CalendarItem>

            <t:Subject>$SafeSubject</t:Subject>

            <t:Body BodyType="Text">$SafeBody</t:Body>

            <t:Start>$SafeStart</t:Start>
            <t:End>$SafeEnd</t:End>

            <t:IsAllDayEvent>$($Details.IsAllDayEvent.ToString().ToLower())</t:IsAllDayEvent>

            <t:LegacyFreeBusyStatus>$Busy</t:LegacyFreeBusyStatus>

            <t:Location>$SafeLoc</t:Location>

        </t:CalendarItem>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Calendar)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Calendar item was created but no destination ItemId was returned.'
    }

    return [PSCustomObject]@{
        ItemId    = $IdNode.GetAttribute('Id')
        ChangeKey = $IdNode.GetAttribute('ChangeKey')
    }
}

function Copy-CalendarItem {

    param(
        [PSCustomObject]$Item
    )

    $Stats.CalendarDetected++

    # FIX: The selected-folder-only discovery (STEP 10) only fetches base
    # item fields, so $Item.Start / .End / .Location / .IsRecurring /
    # .CalendarType do not exist on this object under Set-StrictMode.
    # The old code read those properties directly here -> terminating error
    # -> whole script aborted on the first calendar item.
    #
    # Fetch full calendar details FIRST, then run duplicate/recurring
    # checks against $Details, and do it all inside one try/catch so a
    # failure on one item can never kill the run.
    try {

        $Details =
            Get-EwsCalendarDetails `
                -Item $Item

        $Existing =
            Find-DestinationCalendarDuplicate `
                -Item $Details

        if ($Existing) {

            $Stats.CalendarSkipped++

            Write-Host "SKIPPED CALENDAR: $($Item.Subject)"

            return
        }

        if (
            $Details.IsRecurring -or
            $Details.CalendarItemType -in @(
                'RecurringMaster',
                'Occurrence',
                'Exception'
            )
        ) {

            $Stats.CalendarFailed++

            Write-Host "FAILED CALENDAR: $($Item.Subject)"
            Write-Host '    Recurring calendar items are detected but not yet migrated by this version.'

            return
        }

        $null =
            Create-DestinationCalendarItem `
                -Details $Details

        $Stats.CalendarCopied++

        Write-Host "COPIED CALENDAR: $($Item.Subject)"
    }
    catch {

        $Stats.CalendarFailed++

        Write-Host "FAILED CALENDAR: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}
# ============================================================
# STEP 30: COPY CALENDAR
# ============================================================

Write-Host '========== STEP 30: COPY CALENDAR =========='

foreach ($Item in $AllCalendarItems) {

    Copy-CalendarItem -Item $Item
}

# ============================================================
# STEP 31: CONTACT DUPLICATE CHECK
# ============================================================

Write-Host '========== STEP 31: CONTACT DUPLICATE CHECK =========='

function Get-ContactEmailsFromNode {

    param(
        [System.Xml.XmlNode]$Node
    )

    if (-not $Node) {
        return @()
    }

    return @(
        $Node.SelectNodes(
            't:EmailAddresses/t:Entry',
            $NamespaceManager
        ) |
        ForEach-Object {
            $_.InnerText
        } |
        Where-Object {
            $_
        }
    )
}

function Find-DestinationContactDuplicate {

    param(
        [PSCustomObject]$Contact
    )

    $SafeDisplayName =
        ConvertTo-XmlSafe $Contact.DisplayName

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="contacts:DisplayName"/>
            <t:FieldURI FieldURI="contacts:EmailAddresses"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:Restriction>

        <t:IsEqualTo>

            <t:FieldURI FieldURI="contacts:DisplayName"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeDisplayName"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>

    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$DestinationContactsId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'FindItem(ContactDuplicate)'

    $Nodes =
        $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

    $SourceEmails = @(
        $Contact.EmailAddresses |
        ForEach-Object {
            $_.ToLowerInvariant()
        }
    )

    foreach ($Node in $Nodes) {

        $CandidateEmails = @(
            Get-ContactEmailsFromNode -Node $Node |
            ForEach-Object {
                $_.ToLowerInvariant()
            }
        )

        $EmailMatch = $false

        if ($SourceEmails.Count -gt 0) {

            $EmailMatch =
                @(
                    $SourceEmails |
                    Where-Object {
                        $_ -in $CandidateEmails
                    }
                ).Count -gt 0
        }

        if (
            $EmailMatch -or
            $SourceEmails.Count -eq 0
        ) {

            $IdNode =
                $Node.SelectSingleNode(
                    't:ItemId',
                    $NamespaceManager
                )

            return [PSCustomObject]@{
                ItemId    = $IdNode.GetAttribute('Id')
                ChangeKey = $IdNode.GetAttribute('ChangeKey')
            }
        }
    }

    return $null
}

# ============================================================
# STEP 32: GET CONTACT DETAILS
# ============================================================

Write-Host '========== STEP 32: GET CONTACT DETAILS =========='

function Get-EwsContactDetails {

    param(
        [PSCustomObject]$Item
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Body"/>

            <t:FieldURI FieldURI="contacts:DisplayName"/>
            <t:FieldURI FieldURI="contacts:GivenName"/>
            <t:FieldURI FieldURI="contacts:MiddleName"/>
            <t:FieldURI FieldURI="contacts:Surname"/>
            <t:FieldURI FieldURI="contacts:FileAs"/>
            <t:FieldURI FieldURI="contacts:CompanyName"/>
            <t:FieldURI FieldURI="contacts:JobTitle"/>
            <t:FieldURI FieldURI="contacts:Department"/>
            <t:FieldURI FieldURI="contacts:EmailAddresses"/>
            <t:FieldURI FieldURI="contacts:PhoneNumbers"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$($Item.ItemId)"
            ChangeKey="$($Item.ChangeKey)"/>

    </m:ItemIds>

</m:GetItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Source `
        -OperationName 'GetItem(Contact)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetItemResponseMessage/m:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw 'Contact GetItem returned no item.'
    }

    $Emails = @(
        $Node.SelectNodes(
            't:EmailAddresses/t:Entry',
            $NamespaceManager
        ) |
        ForEach-Object {
            [PSCustomObject]@{
                Key   = $_.GetAttribute('Key')
                Value = $_.InnerText
            }
        }
    )

    $Phones = @(
        $Node.SelectNodes(
            't:PhoneNumbers/t:Entry',
            $NamespaceManager
        ) |
        ForEach-Object {
            [PSCustomObject]@{
                Key   = $_.GetAttribute('Key')
                Value = $_.InnerText
            }
        }
    )

    return [PSCustomObject]@{

        DisplayName =
            Get-NodeText $Node 't:DisplayName'

        GivenName =
            Get-NodeText $Node 't:GivenName'

        MiddleName =
            Get-NodeText $Node 't:MiddleName'

        Surname =
            Get-NodeText $Node 't:Surname'

        FileAs =
            Get-NodeText $Node 't:FileAs'

        CompanyName =
            Get-NodeText $Node 't:CompanyName'

        JobTitle =
            Get-NodeText $Node 't:JobTitle'

        Department =
            Get-NodeText $Node 't:Department'

        Body =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Node 't:Body'
            )

        EmailEntries = $Emails

        PhoneEntries = $Phones

        EmailAddresses = @(
            $Emails |
            ForEach-Object {
                $_.Value
            } |
            Where-Object {
                $_
            }
        )
    }
}

# ============================================================
# STEP 33: CREATE CONTACT
# ============================================================

Write-Host '========== STEP 33: CREATE CONTACT =========='

function Create-DestinationContact {

    param(
        [PSCustomObject]$Details
    )

    $Fields =
        New-Object System.Collections.Generic.List[string]

    foreach (
        $Name in @(
            'DisplayName',
            'GivenName',
            'MiddleName',
            'Surname',
            'FileAs',
            'CompanyName',
            'JobTitle',
            'Department'
        )
    ) {

        $Value = $Details.$Name

        if ($Value) {

            $Safe =
                ConvertTo-XmlSafe $Value

            $Fields.Add(
                "<t:$Name>$Safe</t:$Name>"
            )
        }
    }

    if ($Details.EmailEntries.Count -gt 0) {

        $EmailXml =
            New-Object System.Text.StringBuilder

        [void]$EmailXml.Append(
            '<t:EmailAddresses>'
        )

        foreach ($Entry in $Details.EmailEntries) {

            if ($Entry.Value) {

                $Key =
                    ConvertTo-XmlSafe $Entry.Key

                $Value =
                    ConvertTo-XmlSafe $Entry.Value

                [void]$EmailXml.Append(
                    "<t:Entry Key=`"$Key`">$Value</t:Entry>"
                )
            }
        }

        [void]$EmailXml.Append(
            '</t:EmailAddresses>'
        )

        $Fields.Add(
            $EmailXml.ToString()
        )
    }

    if ($Details.PhoneEntries.Count -gt 0) {

        $PhoneXml =
            New-Object System.Text.StringBuilder

        [void]$PhoneXml.Append(
            '<t:PhoneNumbers>'
        )

        foreach ($Entry in $Details.PhoneEntries) {

            if ($Entry.Value) {

                $Key =
                    ConvertTo-XmlSafe $Entry.Key

                $Value =
                    ConvertTo-XmlSafe $Entry.Value

                [void]$PhoneXml.Append(
                    "<t:Entry Key=`"$Key`">$Value</t:Entry>"
                )
            }
        }

        [void]$PhoneXml.Append(
            '</t:PhoneNumbers>'
        )

        $Fields.Add(
            $PhoneXml.ToString()
        )
    }

    if ($Details.Body) {

        $Fields.Add(
            "<t:Body BodyType=`"Text`">$(ConvertTo-XmlSafe $Details.Body)</t:Body>"
        )
    }

    $Body = @"
<m:CreateItem>

    <m:SavedItemFolderId>
        <t:FolderId Id="$DestinationContactsId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:Contact>

            $($Fields -join "`n            ")

        </t:Contact>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Contact)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Contact was created but no destination ItemId was returned.'
    }

    return $IdNode.GetAttribute('Id')
}

# ============================================================
# STEP 34: COPY CONTACTS
# ============================================================

Write-Host '========== STEP 34: COPY CONTACTS =========='

# IMPORTANT:
# Use $AllTypedItems instead of $AllOtherItems.
#
# This fixes Contacts stored directly in an IPF.Note Public Folder.

foreach (
    $Item in (
        $AllTypedItems |
        Where-Object {
            $_.ItemType -eq 'Contact'
        }
    )
) {

    $Stats.ContactDetected++

    try {

        $Details =
            Get-EwsContactDetails `
                -Item $Item

        $Existing =
            Find-DestinationContactDuplicate `
                -Contact $Details

        if ($Existing) {

            $Stats.ContactSkipped++

            Write-Host "SKIPPED CONTACT: $($Details.DisplayName)"

            continue
        }

        $null =
            Create-DestinationContact `
                -Details $Details

        $Stats.ContactCopied++

        Write-Host "COPIED CONTACT: $($Details.DisplayName)"
    }
    catch {

        $Stats.ContactFailed++

        Write-Host "FAILED CONTACT: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 35: TASK DETAILS
# ============================================================

Write-Host '========== STEP 35: TASK DETAILS =========='

function Get-EwsTaskDetails {

    param(
        [PSCustomObject]$Item
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>

            <t:FieldURI FieldURI="task:StartDate"/>
            <t:FieldURI FieldURI="task:DueDate"/>
            <t:FieldURI FieldURI="task:Status"/>
            <t:FieldURI FieldURI="task:PercentComplete"/>
            <t:FieldURI FieldURI="task:CompleteDate"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$($Item.ItemId)"
            ChangeKey="$($Item.ChangeKey)"/>

    </m:ItemIds>

</m:GetItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Source `
        -OperationName 'GetItem(Task)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetItemResponseMessage/m:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw 'Task GetItem returned no item.'
    }

    return [PSCustomObject]@{

        Subject =
            Get-NodeText $Node 't:Subject'

        Body =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Node 't:Body'
            )

        StartDate =
            Get-NodeText $Node 't:StartDate'

        DueDate =
            Get-NodeText $Node 't:DueDate'

        Status =
            Get-NodeText $Node 't:Status'

        PercentComplete =
            Get-NodeText $Node 't:PercentComplete'

        CompleteDate =
            Get-NodeText $Node 't:CompleteDate'
    }
}

function Find-DestinationTaskDuplicate {

    param(
        [PSCustomObject]$Task
    )

    $SafeSubject =
        ConvertTo-XmlSafe $Task.Subject

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="task:StartDate"/>
            <t:FieldURI FieldURI="task:DueDate"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:Restriction>

        <t:IsEqualTo>

            <t:FieldURI FieldURI="item:Subject"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeSubject"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>

    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$DestinationTasksId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'FindItem(TaskDuplicate)'

    $Nodes =
        $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )

    $SourceStart =
        Convert-ToDateTimeKey $Task.StartDate

    $SourceDue =
        Convert-ToDateTimeKey $Task.DueDate

    foreach ($Node in $Nodes) {

        $CandidateStart =
            Convert-ToDateTimeKey (
                Get-NodeText $Node 't:StartDate'
            )

        $CandidateDue =
            Convert-ToDateTimeKey (
                Get-NodeText $Node 't:DueDate'
            )

        if (
            $CandidateStart -eq $SourceStart -and
            $CandidateDue -eq $SourceDue
        ) {

            $IdNode =
                $Node.SelectSingleNode(
                    't:ItemId',
                    $NamespaceManager
                )

            return [PSCustomObject]@{
                ItemId    = $IdNode.GetAttribute('Id')
                ChangeKey = $IdNode.GetAttribute('ChangeKey')
            }
        }
    }

    return $null
}

# ============================================================
# STEP 36: CREATE TASK
# ============================================================

Write-Host '========== STEP 36: CREATE TASK =========='

function Create-DestinationTask {

    param(
        [PSCustomObject]$Task
    )

    $StartXml =
        if ($Task.StartDate) {
            "<t:StartDate>$(ConvertTo-XmlSafe $Task.StartDate)</t:StartDate>"
        }
        else {
            ''
        }

    $DueXml =
        if ($Task.DueDate) {
            "<t:DueDate>$(ConvertTo-XmlSafe $Task.DueDate)</t:DueDate>"
        }
        else {
            ''
        }

    $CompleteXml =
        if ($Task.CompleteDate) {
            "<t:CompleteDate>$(ConvertTo-XmlSafe $Task.CompleteDate)</t:CompleteDate>"
        }
        else {
            ''
        }

    $PercentXml =
        if ($Task.PercentComplete) {
            "<t:PercentComplete>$(ConvertTo-XmlSafe $Task.PercentComplete)</t:PercentComplete>"
        }
        else {
            ''
        }

    $Status =
        if ($Task.Status) {
            $Task.Status
        }
        else {
            'NotStarted'
        }

    $BodyXml =
        if ($Task.Body) {
            "<t:Body BodyType=`"Text`">$(ConvertTo-XmlSafe $Task.Body)</t:Body>"
        }
        else {
            ''
        }

    $Body = @"
<m:CreateItem>

    <m:SavedItemFolderId>
        <t:FolderId Id="$DestinationTasksId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:Task>

            <t:Subject>$(ConvertTo-XmlSafe $Task.Subject)</t:Subject>

            $BodyXml

            $StartXml
            $DueXml

            <t:Status>$Status</t:Status>

            $PercentXml
            $CompleteXml

        </t:Task>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Task)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Task was created but no destination ItemId was returned.'
    }

    return $IdNode.GetAttribute('Id')
}

# ============================================================
# STEP 37: COPY TASKS
# ============================================================

Write-Host '========== STEP 37: COPY TASKS =========='

foreach (
    $Item in (
        $AllTypedItems |
        Where-Object {
            $_.ItemType -eq 'Task'
        }
    )
) {

    $Stats.TaskDetected++

    try {

        $Details =
            Get-EwsTaskDetails `
                -Item $Item

        $Existing =
            Find-DestinationTaskDuplicate `
                -Task $Details

        if ($Existing) {

            $Stats.TaskSkipped++

            Write-Host "SKIPPED TASK: $($Details.Subject)"

            continue
        }

        $null =
            Create-DestinationTask `
                -Task $Details

        $Stats.TaskCopied++

        Write-Host "COPIED TASK: $($Details.Subject)"
    }
    catch {

        $Stats.TaskFailed++

        Write-Host "FAILED TASK: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 38: NOTE DETAILS
# ============================================================

Write-Host '========== STEP 38: NOTE DETAILS =========='

function Get-EwsNoteDetails {

    param(
        [PSCustomObject]$Item
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>
            <t:FieldURI FieldURI="item:ItemClass"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$($Item.ItemId)"
            ChangeKey="$($Item.ChangeKey)"/>

    </m:ItemIds>

</m:GetItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Source `
        -OperationName 'GetItem(Note)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetItemResponseMessage/m:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw 'Note GetItem returned no item.'
    }

    return [PSCustomObject]@{

        Subject =
            Get-NodeText $Node 't:Subject'

        Body =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Node 't:Body'
            )

        ItemClass =
            Get-NodeText $Node 't:ItemClass'
    }
}

function Get-DestinationItemsBySubject {

    param(
        [string]$FolderId,
        [string]$Subject,
        [string]$OperationName
    )

    $SafeSubject =
        ConvertTo-XmlSafe $Subject

    $Body = @"
<m:FindItem Traversal="Shallow">

    <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
    </m:ItemShape>

    <m:Restriction>

        <t:IsEqualTo>

            <t:FieldURI FieldURI="item:Subject"/>

            <t:FieldURIOrConstant>
                <t:Constant Value="$SafeSubject"/>
            </t:FieldURIOrConstant>

        </t:IsEqualTo>

    </m:Restriction>

    <m:ParentFolderIds>
        <t:FolderId Id="$FolderId"/>
    </m:ParentFolderIds>

</m:FindItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName $OperationName

    return @(
        $Xml.SelectNodes(
            '//m:FindItemResponseMessage/m:RootFolder/t:Items/*',
            $NamespaceManager
        )
    )
}

function Find-DestinationNoteDuplicate {

    param(
        [PSCustomObject]$Note
    )

    $Nodes =
        Get-DestinationItemsBySubject `
            -FolderId $DestinationNotesId `
            -Subject $Note.Subject `
            -OperationName 'FindItem(NoteDuplicate)'

    foreach ($Node in $Nodes) {

        $IdNode =
            $Node.SelectSingleNode(
                't:ItemId',
                $NamespaceManager
            )

        if (-not $IdNode) {
            continue
        }

        $Temp = [PSCustomObject]@{
            ItemId =
                $IdNode.GetAttribute('Id')

            ChangeKey =
                $IdNode.GetAttribute('ChangeKey')
        }

        $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>
            <t:FieldURI FieldURI="item:ItemClass"/>
        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$($Temp.ItemId)"
            ChangeKey="$($Temp.ChangeKey)"/>

    </m:ItemIds>

</m:GetItem>
"@

        $Xml = Invoke-EwsRequest `
            -Body $Body `
            -Target Destination `
            -OperationName 'GetItem(NoteDuplicate)'

        $Candidate =
            $Xml.SelectSingleNode(
                '//m:GetItemResponseMessage/m:Items/*',
                $NamespaceManager
            )

        if (-not $Candidate) {
            continue
        }

        $CandidateBody =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Candidate 't:Body'
            )

        if ($CandidateBody -eq $Note.Body) {
            return $Temp
        }
    }

    return $null
}

# ============================================================
# STEP 39: CREATE NOTES
# ============================================================

Write-Host '========== STEP 39: CREATE NOTES =========='

function Create-DestinationNote {

    param(
        [PSCustomObject]$Note
    )

    $Body = @"
<m:CreateItem MessageDisposition="SaveOnly">

    <m:SavedItemFolderId>
        <t:FolderId Id="$DestinationNotesId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:Message>

            <t:Subject>$(ConvertTo-XmlSafe $Note.Subject)</t:Subject>

            <t:Body BodyType="Text">
                $(ConvertTo-XmlSafe $Note.Body)
            </t:Body>

            <t:ExtendedProperty>

                <t:ExtendedFieldURI
                    PropertyTag="0x001A"
                    PropertyType="String"/>

                <t:Value>IPM.StickyNote</t:Value>

            </t:ExtendedProperty>

        </t:Message>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Note)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Note was created but no destination ItemId was returned.'
    }

    return $IdNode.GetAttribute('Id')
}

# ============================================================
# STEP 40: COPY NOTES
# ============================================================

Write-Host '========== STEP 40: COPY NOTES =========='

foreach (
    $Item in (
        $AllTypedItems |
        Where-Object {
            $_.ItemType -eq 'Note'
        }
    )
) {

    $Stats.NoteDetected++

    try {

        $Details =
            Get-EwsNoteDetails `
                -Item $Item

        $Existing =
            Find-DestinationNoteDuplicate `
                -Note $Details

        if ($Existing) {

            $Stats.NoteSkipped++

            Write-Host "SKIPPED NOTE: $($Details.Subject)"

            continue
        }

        $null =
            Create-DestinationNote `
                -Note $Details

        $Stats.NoteCopied++

        Write-Host "COPIED NOTE: $($Details.Subject)"
    }
    catch {

        $Stats.NoteFailed++

        Write-Host "FAILED NOTE: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 41: JOURNAL DETAILS
# ============================================================

Write-Host '========== STEP 41: JOURNAL DETAILS =========='

function Get-EwsJournalDetails {

    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Item
    )

    $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$($Item.ItemId)"
            ChangeKey="$($Item.ChangeKey)"/>

    </m:ItemIds>

</m:GetItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Source `
        -OperationName 'GetItem(Journal)'

    $Node =
        $Xml.SelectSingleNode(
            '//m:GetItemResponseMessage/m:Items/*',
            $NamespaceManager
        )

    if (-not $Node) {
        throw 'Journal GetItem returned no item.'
    }

    $MessageClass =
        Get-NodeText `
            -Node $Node `
            -XPath 't:ItemClass'

    if ([string]::IsNullOrWhiteSpace($MessageClass)) {
        $MessageClass = $Item.ItemClass
    }

    if ([string]::IsNullOrWhiteSpace($MessageClass)) {
        $MessageClass = 'IPM.Activity'
    }

    return [PSCustomObject]@{

        Subject =
            Get-NodeText $Node 't:Subject'

        Body =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Node 't:Body'
            )

        MessageClass =
            $MessageClass
    }
}

# ============================================================
# STEP 42: JOURNAL DUPLICATE CHECK
# ============================================================

Write-Host '========== STEP 42: JOURNAL DUPLICATE CHECK =========='

function Find-DestinationJournalDuplicate {

    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Journal
    )

    $Nodes =
        Get-DestinationItemsBySubject `
            -FolderId $DestinationJournalId `
            -Subject $Journal.Subject `
            -OperationName 'FindItem(JournalDuplicate)'

    foreach ($Node in $Nodes) {

        $IdNode =
            $Node.SelectSingleNode(
                't:ItemId',
                $NamespaceManager
            )

        if (-not $IdNode) {
            continue
        }

        $CandidateId =
            $IdNode.GetAttribute('Id')

        $CandidateChangeKey =
            $IdNode.GetAttribute('ChangeKey')

        $Body = @"
<m:GetItem>

    <m:ItemShape>

        <t:BaseShape>IdOnly</t:BaseShape>

        <t:AdditionalProperties>

            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="item:Body"/>
            <t:FieldURI FieldURI="item:ItemClass"/>

        </t:AdditionalProperties>

    </m:ItemShape>

    <m:ItemIds>

        <t:ItemId
            Id="$CandidateId"
            ChangeKey="$CandidateChangeKey"/>

    </m:ItemIds>

</m:GetItem>
"@

        try {

            $Xml = Invoke-EwsRequest `
                -Body $Body `
                -Target Destination `
                -OperationName 'GetItem(JournalDuplicate)'
        }
        catch {

            Write-Host "    Could not inspect Journal candidate: $($_.Exception.Message)"

            continue
        }

        $Candidate =
            $Xml.SelectSingleNode(
                '//m:GetItemResponseMessage/m:Items/*',
                $NamespaceManager
            )

        if (-not $Candidate) {
            continue
        }

        $CandidateClass =
            Get-NodeText $Candidate 't:ItemClass'

        $CandidateSubject =
            Get-NodeText $Candidate 't:Subject'

        $CandidateBody =
            Convert-EwsBodyToPlainText (
                Get-NodeText $Candidate 't:Body'
            )

        if (
            $CandidateClass -eq 'IPM.Activity' -and
            $CandidateSubject -eq $Journal.Subject -and
            $CandidateBody -eq $Journal.Body
        ) {

            return [PSCustomObject]@{
                ItemId    = $CandidateId
                ChangeKey = $CandidateChangeKey
            }
        }
    }

    return $null
}

# ============================================================
# STEP 43: CREATE + COPY JOURNALS
# ============================================================

Write-Host '========== STEP 43: COPY JOURNALS =========='

function Create-DestinationJournal {

    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Journal
    )

    $Body = @"
<m:CreateItem MessageDisposition="SaveOnly">

    <m:SavedItemFolderId>
        <t:FolderId Id="$DestinationJournalId"/>
    </m:SavedItemFolderId>

    <m:Items>

        <t:Message>

            <t:Subject>$(ConvertTo-XmlSafe $Journal.Subject)</t:Subject>

            <t:Body BodyType="Text">
                $(ConvertTo-XmlSafe $Journal.Body)
            </t:Body>

            <t:ExtendedProperty>

                <t:ExtendedFieldURI
                    PropertyTag="0x001A"
                    PropertyType="String"/>

                <t:Value>IPM.Activity</t:Value>

            </t:ExtendedProperty>

        </t:Message>

    </m:Items>

</m:CreateItem>
"@

    $Xml = Invoke-EwsRequest `
        -Body $Body `
        -Target Destination `
        -OperationName 'CreateItem(Journal)'

    $IdNode =
        $Xml.SelectSingleNode(
            '//m:CreateItemResponseMessage/m:Items/*/t:ItemId',
            $NamespaceManager
        )

    if (-not $IdNode) {
        throw 'Journal was created but no destination ItemId was returned.'
    }

    return $IdNode.GetAttribute('Id')
}

foreach (
    $Item in (
        $AllTypedItems |
        Where-Object {
            $_.ItemType -eq 'Journal'
        }
    )
) {

    $Stats.JournalDetected++

    try {

        Write-Host ''
        Write-Host "Journal: $($Item.Subject)"

        $Details =
            Get-EwsJournalDetails `
                -Item $Item

        Write-Host "    Class: $($Details.MessageClass)"

        $Existing =
            Find-DestinationJournalDuplicate `
                -Journal $Details

        if ($Existing) {

            $Stats.JournalSkipped++

            Write-Host "SKIPPED JOURNAL: $($Details.Subject)"

            continue
        }

        $null =
            Create-DestinationJournal `
                -Journal $Details

        $Stats.JournalCopied++

        Write-Host "COPIED JOURNAL: $($Details.Subject)"
    }
    catch {

        $Stats.JournalFailed++

        Write-Host "FAILED JOURNAL: $($Item.Subject)"
        Write-Host "    $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 44: REMAINING ITEM TYPES
# ============================================================

Write-Host '========== STEP 44: REMAINING ITEM TYPES =========='

if ($DetectionOnlyItems.Count -eq 0) {

    Write-Host 'Detection-only: None'
}
else {

    $Grouped =
        $DetectionOnlyItems |
        Group-Object ItemType

    foreach ($Group in $Grouped) {

        Write-Host "$($Group.Name): $($Group.Count)"

        foreach ($Item in $Group.Group) {

            Write-Host "    $($Item.FolderPath) -> $($Item.Subject) [$($Item.ItemClass)]"
        }
    }
}

# ============================================================
# STEP 45: FINAL MIGRATION SUMMARY
# ============================================================

Write-Host '========== STEP 45: FINAL MIGRATION SUMMARY =========='

Write-Host ''

Write-Host '============================================================'
Write-Host 'MIGRATION SUMMARY'
Write-Host '============================================================'

Write-Host "Source PF mailbox : $SourcePublicFolderMailbox"
Write-Host "Destination       : $DestinationMailbox"

Write-Host ''

Write-Host "Source folders discovered      : $($Stats.FoldersDiscovered)"
Write-Host "Destination folders created    : 0"
Write-Host "Destination hierarchy mirrored : NO"

Write-Host ''

Write-Host 'MAIL'
Write-Host "  Detected : $($Stats.MailDetected)"
Write-Host "  Copied   : $($Stats.MailCopied)"
Write-Host "  Skipped  : $($Stats.MailSkipped)"
Write-Host "  Failed   : $($Stats.MailFailed)"

Write-Host ''

Write-Host 'CALENDAR'
Write-Host "  Detected : $($Stats.CalendarDetected)"
Write-Host "  Copied   : $($Stats.CalendarCopied)"
Write-Host "  Skipped  : $($Stats.CalendarSkipped)"
Write-Host "  Failed   : $($Stats.CalendarFailed)"

Write-Host ''

Write-Host 'CONTACTS'
Write-Host "  Detected : $($Stats.ContactDetected)"
Write-Host "  Copied   : $($Stats.ContactCopied)"
Write-Host "  Skipped  : $($Stats.ContactSkipped)"
Write-Host "  Failed   : $($Stats.ContactFailed)"

Write-Host ''

Write-Host 'TASKS'
Write-Host "  Detected : $($Stats.TaskDetected)"
Write-Host "  Copied   : $($Stats.TaskCopied)"
Write-Host "  Skipped  : $($Stats.TaskSkipped)"
Write-Host "  Failed   : $($Stats.TaskFailed)"

Write-Host ''

Write-Host 'NOTES'
Write-Host "  Detected : $($Stats.NoteDetected)"
Write-Host "  Copied   : $($Stats.NoteCopied)"
Write-Host "  Skipped  : $($Stats.NoteSkipped)"
Write-Host "  Failed   : $($Stats.NoteFailed)"

Write-Host ''

Write-Host 'JOURNALS'
Write-Host "  Detected : $($Stats.JournalDetected)"
Write-Host "  Copied   : $($Stats.JournalCopied)"
Write-Host "  Skipped  : $($Stats.JournalSkipped)"
Write-Host "  Failed   : $($Stats.JournalFailed)"

Write-Host ''

Write-Host "Detection-only : $($Stats.DetectionOnly)"

Write-Host ''

Write-Host 'Source modifications: 0'
Write-Host 'Source deletions:     0'

Write-Host ''

Write-Host 'SOURCE WAS NOT MODIFIED.'

Write-Host '============================================================'

# ============================================================
# STEP 46: CLEANUP
# ============================================================

Write-Host '========== STEP 46: CLEANUP =========='

$Token = $null
$SecretText = $null

Write-Host 'Migration script finished.'