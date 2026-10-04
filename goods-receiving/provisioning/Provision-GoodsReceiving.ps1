#Requires -Version 7.2
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }
<#
.SYNOPSIS
    Creates the SharePoint list, document library, columns, indexes, version
    history settings and groups used by the Goods Receiving web part.

.DESCRIPTION
    The script reads the same configuration files as the web part:

        src/config/fields.json    every field of a receiving record
        src/config/solution.json  list / library / group names, document types

    It is safe to run more than once. Anything that already exists is left
    as it is, with two exceptions that only ever ADD things:
      * new choice values in fields.json are added to existing choice columns
      * missing indexes and missing default-view columns are added

    Nothing is ever deleted and no column type is ever changed.

.PARAMETER SiteUrl
    Full URL of the SharePoint site, e.g. https://contoso.sharepoint.com/sites/warehouse

.PARAMETER ClientId
    Application (client) ID of the Entra ID app registration that PnP.PowerShell
    uses for interactive sign-in. See DEPLOYMENT.md, section "Provisioning".
    Can also be supplied through the ENTRAID_APP_ID environment variable.

.PARAMETER UseExistingConnection
    Skip Connect-PnPOnline and use the connection you already opened.

.PARAMETER SkipPermissions
    Do not break permission inheritance on the list and library and do not
    grant the two groups access. Use this if your site owner manages
    permissions manually.

.PARAMETER SkipPage
    Do not create the "Goods Receiving" page. By default the script creates a
    published page (banner + intro text + the web part) and adds it to the
    site navigation. This needs the .sppkg to be deployed first (DEPLOYMENT.md,
    section 4); if it is not, the script skips the page and tells you to re-run it.

.PARAMETER ConfigFolder
    Folder that contains fields.json and solution.json. Defaults to ../src/config.

.EXAMPLE
    ./Provision-GoodsReceiving.ps1 -SiteUrl https://contoso.sharepoint.com/sites/warehouse -ClientId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/warehouse -Interactive -ClientId <id>
    ./Provision-GoodsReceiving.ps1 -SiteUrl https://contoso.sharepoint.com/sites/warehouse -UseExistingConnection
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    [string] $ClientId = $env:ENTRAID_APP_ID,

    [switch] $UseExistingConnection,

    [switch] $SkipPermissions,

    [switch] $SkipPage,

    [string] $ConfigFolder = (Join-Path $PSScriptRoot '..' 'src' 'config'),

    [string] $ManifestPath = (Join-Path $PSScriptRoot '..' 'src' 'webparts' 'goodsReceiving' 'GoodsReceivingWebPart.manifest.json')
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------
function Write-Step([string] $Message) { Write-Host "`n==> $Message" -ForegroundColor Cyan }
function Write-Created([string] $Message) { Write-Host "    [created] $Message" -ForegroundColor Green }
function Write-Updated([string] $Message) { Write-Host "    [updated] $Message" -ForegroundColor Yellow }
function Write-Exists([string] $Message) { Write-Host "    [exists]  $Message" -ForegroundColor DarkGray }
function Write-Warn([string] $Message) { Write-Host "    [warning] $Message" -ForegroundColor Magenta }

# Returns $Object.$Name when present and not null, otherwise $Default.
# The JSON files use optional properties, so this keeps the code readable.
function Get-Value($Object, [string] $Name, $Default = $null) {
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Name -and $null -ne $Object.$Name) {
        return $Object.$Name
    }
    return $Default
}

function ConvertTo-XmlText([string] $Text) {
    return [System.Security.SecurityElement]::Escape($Text)
}

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$fieldsPath = Join-Path $ConfigFolder 'fields.json'
$solutionPath = Join-Path $ConfigFolder 'solution.json'
foreach ($path in @($fieldsPath, $solutionPath)) {
    if (-not (Test-Path $path)) { throw "Configuration file not found: $path" }
}
$fieldsConfig = Get-Content $fieldsPath -Raw | ConvertFrom-Json
$solution = Get-Content $solutionPath -Raw | ConvertFrom-Json
$fields = @($fieldsConfig.fields | Sort-Object { [double]$_.order })

$indexedCount = @($fields | Where-Object { Get-Value $_ 'indexed' $false }).Count
if ($indexedCount -gt 18) {
    # SharePoint allows 20 indexes per list; keep headroom for the Created/Modified defaults.
    throw "fields.json marks $indexedCount fields as indexed. SharePoint allows at most 20 indexes per list."
}

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------
Write-Step "Connecting to $SiteUrl"
if ($UseExistingConnection) {
    $connection = Get-PnPConnection
    if ($connection.Url.TrimEnd('/') -ne $SiteUrl.TrimEnd('/')) {
        throw "The existing PnP connection points to $($connection.Url), not $SiteUrl."
    }
} else {
    if (-not $ClientId) {
        throw 'Provide -ClientId (or set the ENTRAID_APP_ID environment variable). See DEPLOYMENT.md, section "Provisioning".'
    }
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
}
Write-Exists "Connected as $((Get-PnPCurrentUser).Email)"

# ---------------------------------------------------------------------------
# Lists
# ---------------------------------------------------------------------------
function Get-OrCreateList([string] $Title, [string] $Url, [string] $Template, [string] $Description) {
    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($list) {
        Write-Exists "'$Title'"
        return $list
    }
    # Not added to the site navigation on purpose: people should work through the
    # Goods Receiving page, where the app applies the workflow rules.
    New-PnPList -Title $Title -Url $Url -Template $Template | Out-Null
    $list = Get-PnPList -Identity $Title
    if ($Description) {
        Set-PnPList -Identity $list -Description $Description | Out-Null
    }
    Write-Created "'$Title' ($Template)"
    return $list
}

function Set-ListVersioning($List, [int] $MajorVersionLimit) {
    $enabled = Get-PnPProperty -ClientObject $List -Property EnableVersioning
    if ($enabled) {
        Write-Exists "Version history already enabled on '$($List.Title)'"
        return
    }
    Set-PnPList -Identity $List -EnableVersioning $true -MajorVersions $MajorVersionLimit | Out-Null
    Write-Updated "Version history enabled on '$($List.Title)' (keeps $MajorVersionLimit major versions)"
}

# ---------------------------------------------------------------------------
# Fields
# ---------------------------------------------------------------------------
function Get-ListField($List, [string] $InternalName) {
    try {
        return Get-PnPField -List $List -Identity $InternalName -ErrorAction Stop
    } catch {
        return $null
    }
}

# Builds the CAML <Field> schema for one entry of fields.json.
function New-FieldSchemaXml($Field, [hashtable] $LookupListIds) {
    $name = $Field.internalName
    $validation = Get-Value $Field 'validation'
    $attributes = [ordered]@{
        Name        = $name
        StaticName  = $name
        DisplayName = $Field.displayName
        # Only fields that are needed even for drafts are required at SharePoint level.
        # Everything else is enforced by the app when the record is submitted.
        Required    = if (Get-Value $Field 'requiredForDraft' $false) { 'TRUE' } else { 'FALSE' }
    }
    $description = Get-Value $Field 'description'
    if ($description) { $attributes['Description'] = $description }
    $inner = ''
    $default = Get-Value $Field 'defaultValue'

    switch ($Field.type) {
        'Text' {
            $attributes['Type'] = 'Text'
            $attributes['MaxLength'] = [string][Math]::Min([int](Get-Value $validation 'maxLength' 255), 255)
        }
        'Note' {
            $attributes['Type'] = 'Note'
            $attributes['NumLines'] = '6'
            $attributes['RichText'] = 'FALSE'
            $attributes['AppendOnly'] = 'FALSE'
        }
        'Number' {
            $attributes['Type'] = 'Number'
            $attributes['Decimals'] = [string](Get-Value $Field 'decimals' 0)
            $min = Get-Value $validation 'min'
            $max = Get-Value $validation 'max'
            if ($null -ne $min) { $attributes['Min'] = [string]$min }
            if ($null -ne $max) { $attributes['Max'] = [string]$max }
            # "Greater than" / "less than" rules become a column validation formula so
            # they are enforced by SharePoint too (blank stays allowed for drafts).
            $conditions = @()
            $ref = "[$($Field.displayName)]"
            $minExclusive = Get-Value $validation 'minExclusive'
            $maxExclusive = Get-Value $validation 'maxExclusive'
            if ($null -ne $minExclusive) { $conditions += "$ref>$minExclusive" }
            if ($null -ne $maxExclusive) { $conditions += "$ref<$maxExclusive" }
            if ($conditions.Count -gt 0) {
                $rule = if ($conditions.Count -eq 1) { $conditions[0] } else { "AND($($conditions -join ','))" }
                $message = "$($Field.displayName) must be"
                if ($null -ne $minExclusive) { $message += " greater than $minExclusive" }
                if ($null -ne $minExclusive -and $null -ne $maxExclusive) { $message += ' and' }
                if ($null -ne $maxExclusive) { $message += " less than $maxExclusive" }
                $message += '.'
                $inner += "<Validation Message=`"$(ConvertTo-XmlText $message)`">$(ConvertTo-XmlText "=OR(ISBLANK($ref),$rule)")</Validation>"
            }
        }
        'Date' {
            $attributes['Type'] = 'DateTime'
            $attributes['Format'] = 'DateOnly'
            $attributes['FriendlyDisplayFormat'] = 'Disabled'
            if ($default -eq '[today]') { $inner += '<Default>[today]</Default>' }
        }
        'DateTime' {
            $attributes['Type'] = 'DateTime'
            $attributes['Format'] = 'DateTime'
            $attributes['FriendlyDisplayFormat'] = 'Disabled'
            if ($default -eq '[today]') { $inner += '<Default>[today]</Default>' }
        }
        'Choice' {
            $attributes['Type'] = 'Choice'
            $attributes['Format'] = 'Dropdown'
            $attributes['FillInChoice'] = 'FALSE'
            $inner += '<CHOICES>'
            foreach ($choice in $Field.choices) { $inner += "<CHOICE>$(ConvertTo-XmlText $choice)</CHOICE>" }
            $inner += '</CHOICES>'
            if ($default) { $inner += "<Default>$(ConvertTo-XmlText $default)</Default>" }
        }
        'User' {
            $attributes['Type'] = 'User'
            $attributes['List'] = 'UserInfo'
            $attributes['ShowField'] = 'ImnName'
            $attributes['UserSelectionMode'] = 'PeopleOnly'
            $attributes['UserSelectionScope'] = '0'
        }
        'Lookup' {
            $lookup = Get-Value $Field 'lookup'
            $listId = $LookupListIds[$lookup.listTitle]
            $attributes['Type'] = 'Lookup'
            $attributes['List'] = "{$listId}"
            $attributes['ShowField'] = $lookup.showField
        }
        default {
            throw "Field '$name' has unsupported type '$($Field.type)'."
        }
    }

    $attributeText = ($attributes.GetEnumerator() | ForEach-Object { '{0}="{1}"' -f $_.Key, (ConvertTo-XmlText ([string]$_.Value)) }) -join ' '
    return "<Field $attributeText>$inner</Field>"
}

# Adds choices from fields.json that the existing column does not have yet.
function Sync-ChoiceField($List, $ExistingField, $Field) {
    [xml] $schema = $ExistingField.SchemaXml
    $current = @($schema.SelectNodes('//CHOICE') | ForEach-Object { $_.InnerText })
    $missing = @($Field.choices | Where-Object { $current -notcontains $_ })
    if ($missing.Count -eq 0) { return }
    $choicesNode = $schema.SelectSingleNode('/Field/CHOICES')
    if (-not $choicesNode) {
        $choicesNode = $schema.CreateElement('CHOICES')
        [void] $schema.DocumentElement.AppendChild($choicesNode)
    }
    foreach ($choice in $missing) {
        $node = $schema.CreateElement('CHOICE')
        $node.InnerText = $choice
        [void] $choicesNode.AppendChild($node)
    }
    Set-PnPField -List $List -Identity $Field.internalName -Values @{ SchemaXml = $schema.OuterXml } | Out-Null
    Write-Updated "$($Field.displayName): added choice(s) $($missing -join ', ')"
}

function Set-FieldIndex($List, [string] $InternalName, [string] $DisplayName) {
    $field = Get-ListField $List $InternalName
    if ($field.Indexed) { return }
    Set-PnPField -List $List -Identity $InternalName -Values @{ Indexed = $true } | Out-Null
    Write-Updated "$DisplayName indexed"
}

function Set-ListFields($List, $Fields) {
    # Resolve lookup target lists once.
    $lookupListIds = @{}
    foreach ($field in @($Fields | Where-Object { $_.type -eq 'Lookup' })) {
        $title = $field.lookup.listTitle
        if (-not $lookupListIds.ContainsKey($title)) {
            $target = Get-PnPList -Identity $title -ErrorAction SilentlyContinue
            if ($target) { $lookupListIds[$title] = $target.Id }
        }
    }

    foreach ($field in $Fields) {
        $name = $field.internalName

        if (Get-Value $field 'builtIn' $false) {
            # Title is SharePoint's built-in column; it stores the Record ID.
            if ($name -eq 'Title') {
                $title = Get-ListField $List 'Title'
                $values = @{}
                if ($title.Title -ne $field.displayName) { $values['Title'] = $field.displayName }
                if ($title.Required) { $values['Required'] = $false }
                if ($values.Count -gt 0) {
                    Set-PnPField -List $List -Identity 'Title' -Values $values | Out-Null
                    Write-Updated "Title column renamed to '$($field.displayName)' and made optional (filled in automatically)"
                } else {
                    Write-Exists $field.displayName
                }
                if (Get-Value $field 'indexed' $false) { Set-FieldIndex $List 'Title' $field.displayName }
            }
            continue
        }

        if ($field.type -eq 'Lookup' -and -not $lookupListIds.ContainsKey($field.lookup.listTitle)) {
            Write-Warn "$($field.displayName): lookup list '$($field.lookup.listTitle)' not found - column skipped. Create that list first."
            continue
        }

        $existing = Get-ListField $List $name
        if ($existing) {
            if ($existing.InternalName -ne $name) {
                Write-Warn "$($field.displayName): a column with title '$name' exists but its internal name is '$($existing.InternalName)'. Skipped - check fields.json."
                continue
            }
            Write-Exists $field.displayName
            if ($field.type -eq 'Choice') { Sync-ChoiceField $List $existing $field }
        } else {
            $xml = New-FieldSchemaXml $field $lookupListIds
            Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            Write-Created "$($field.displayName) ($($field.type))"
        }

        if (Get-Value $field 'indexed' $false) { Set-FieldIndex $List $name $field.displayName }
    }
}

# Adds missing columns to the default view (never removes any).
function Add-DefaultViewFields($List, [string[]] $InternalNames) {
    $view = Get-PnPView -List $List -Includes ViewFields | Where-Object { $_.DefaultView } | Select-Object -First 1
    if (-not $view) { return }
    $current = @($view.ViewFields)
    $missing = @($InternalNames | Where-Object { $current -notcontains $_ })
    if ($missing.Count -eq 0) { return }
    Set-PnPView -List $List -Identity $view.Id -Fields ($current + $missing) | Out-Null
    Write-Updated "Default view '$($view.Title)': added $($missing -join ', ')"
}

# ---------------------------------------------------------------------------
# 1. Records list
# ---------------------------------------------------------------------------
Write-Step "List '$($solution.list.title)'"
$list = Get-OrCreateList $solution.list.title $solution.list.url 'GenericList' (Get-Value $solution.list 'description')
Set-ListVersioning $list ([int](Get-Value $solution.list 'majorVersionLimit' 500))
# Files go to the document library, not to list item attachments.
if (Get-PnPProperty -ClientObject $list -Property EnableAttachments) {
    Set-PnPList -Identity $list -EnableAttachments $false | Out-Null
    Write-Updated 'List item attachments disabled (files are stored in the document library)'
}
try {
    # Grid ("Edit in grid view") editing would bypass the app's validation rules.
    Set-PnPList -Identity $list -DisableGridEditing $true | Out-Null
} catch {
    Write-Warn "Could not disable grid editing: $($_.Exception.Message)"
}

Write-Step 'Columns and indexes'
Set-ListFields $list $fields

$tableColumns = @($fields | Where-Object { $_.showInTable -and -not (Get-Value $_ 'builtIn' $false) } | ForEach-Object { $_.internalName })
Add-DefaultViewFields $list (@('LinkTitle') + $tableColumns)

# ---------------------------------------------------------------------------
# 2. Documents library
# ---------------------------------------------------------------------------
$lib = $solution.library
Write-Step "Library '$($lib.title)'"
$library = Get-OrCreateList $lib.title $lib.url 'DocumentLibrary' (Get-Value $lib 'description')
Set-ListVersioning $library ([int](Get-Value $lib 'majorVersionLimit' 500))

$libraryFields = @(
    [pscustomobject]@{
        internalName = $lib.recordLookupField.internalName
        displayName  = $lib.recordLookupField.displayName
        type         = 'Lookup'
        lookup       = [pscustomobject]@{ listTitle = $solution.list.title; showField = 'Title' }
        indexed      = $true
        description  = 'The goods receiving record this document belongs to.'
    },
    [pscustomobject]@{
        internalName = $lib.documentTypeField.internalName
        displayName  = $lib.documentTypeField.displayName
        type         = 'Choice'
        choices      = @($lib.documentTypes)
        indexed      = $true
    }
)
Write-Step 'Library columns'
Set-ListFields $library $libraryFields
Add-DefaultViewFields $library @($lib.documentTypeField.internalName, $lib.recordLookupField.internalName)

# ---------------------------------------------------------------------------
# 3. Groups
# ---------------------------------------------------------------------------
Write-Step 'SharePoint groups'
$ownerGroupTitle = $null
try { $ownerGroupTitle = (Get-PnPGroup -AssociatedOwnerGroup).Title } catch { $ownerGroupTitle = $null }

function Get-OrCreateGroup([string] $Name, [string] $Description) {
    $group = $null
    try { $group = Get-PnPGroup -Identity $Name -ErrorAction Stop } catch { $group = $null }
    if ($group) {
        Write-Exists "'$Name'"
        return $group
    }
    if ($ownerGroupTitle) {
        $group = New-PnPGroup -Title $Name -Description $Description -Owner $ownerGroupTitle
    } else {
        $group = New-PnPGroup -Title $Name -Description $Description
    }
    Write-Created "'$Name'"
    return $group
}

$receiversGroup = Get-OrCreateGroup $solution.groups.receivers 'Warehouse staff who record incoming material shipments.'
$supervisorsGroup = Get-OrCreateGroup $solution.groups.supervisors 'Supervisors who approve or reject goods receiving records.'

# ---------------------------------------------------------------------------
# 4. Permissions
# ---------------------------------------------------------------------------
if ($SkipPermissions) {
    Write-Step 'Permissions skipped (-SkipPermissions)'
} else {
    Write-Step 'Permissions'
    # Look up the built-in level by type, not by name: names are translated on non-English sites.
    $contribute = Get-PnPRoleDefinition | Where-Object { $_.RoleTypeKind -eq 'Contributor' } | Select-Object -First 1
    if (-not $contribute) { throw 'Could not find the built-in Contribute permission level on this site.' }

    # Supervisors get Contribute + "Approve Items" + "Override List Behaviors".
    #  - Approve Items: the web part also uses it to recognise supervisors who are
    #    members through a nested Microsoft 365 / security group.
    #  - Override List Behaviors (CancelCheckout): lets supervisors edit everyone's
    #    records if you switch on item-level permissions (see README, "Security model").
    $supervisorLevelName = $solution.groups.supervisorPermissionLevel
    $supervisorLevel = $null
    try { $supervisorLevel = Get-PnPRoleDefinition -Identity $supervisorLevelName -ErrorAction Stop } catch { $supervisorLevel = $null }
    if ($supervisorLevel) {
        Write-Exists "Permission level '$supervisorLevelName'"
    } else {
        Add-PnPRoleDefinition -RoleName $supervisorLevelName -Clone $contribute.Name -Include ApproveItems, CancelCheckout `
            -Description 'Contribute plus Approve Items and Override List Behaviors. Used by the Goods Receiving supervisors group.' | Out-Null
        Write-Created "Permission level '$supervisorLevelName' (Contribute + Approve Items + Override List Behaviors)"
    }

    foreach ($target in @($list, $library)) {
        $unique = Get-PnPProperty -ClientObject $target -Property HasUniqueRoleAssignments
        if (-not $unique) {
            # Copy the existing assignments so site owners and visitors keep their access.
            Set-PnPList -Identity $target -BreakRoleInheritance -CopyRoleAssignments | Out-Null
            Write-Updated "'$($target.Title)': stopped inheriting permissions (existing access copied)"
        }
        $supervisorRole = if ($target.Id -eq $list.Id) { $supervisorLevelName } else { $contribute.Name }
        Set-PnPListPermission -Identity $target -Group $receiversGroup.Title -AddRole $contribute.Name | Out-Null
        Set-PnPListPermission -Identity $target -Group $supervisorsGroup.Title -AddRole $supervisorRole | Out-Null
        Write-Updated "'$($target.Title)': '$($receiversGroup.Title)' = $($contribute.Name), '$($supervisorsGroup.Title)' = $supervisorRole"
    }
}

# ---------------------------------------------------------------------------
# 5. Page with the web part
# ---------------------------------------------------------------------------
# Reads the web part ID from its manifest (the first "id" in the file).
function Get-WebPartId([string] $Path) {
    if (-not (Test-Path $Path)) { return $null }
    $match = [regex]::Match((Get-Content $Path -Raw), '"id"\s*:\s*"([0-9a-fA-F-]{36})"')
    if ($match.Success) { return $match.Groups[1].Value }
    return $null
}

function Get-ExistingPage([string] $Name) {
    try { return Get-PnPPage -Identity $Name -ErrorAction Stop } catch { return $null }
}

$pageConfig = Get-Value $solution 'page'
$pageResult = 'skipped'
if ($SkipPage) {
    Write-Step 'Page skipped (-SkipPage)'
} elseif (-not $pageConfig) {
    Write-Step 'Page skipped (no "page" section in solution.json)'
} else {
    $pageName = $pageConfig.name
    $pageTitle = Get-Value $pageConfig 'title' $pageName
    Write-Step "Page '$pageTitle' (SitePages/$pageName.aspx)"

    if (Get-ExistingPage $pageName) {
        Write-Exists "Page '$pageTitle' (left unchanged)"
        $pageResult = 'exists'
    } else {
        # A page has to exist before SharePoint can list the web parts available for it.
        Add-PnPPage -Name $pageName -LayoutType Article | Out-Null
        Set-PnPPage -Identity $pageName -Title $pageTitle | Out-Null
        try {
            # Banner style: the title on a coloured block (no picture needed).
            Set-PnPPage -Identity $pageName -HeaderLayoutType ColorBlock | Out-Null
        } catch {
            Write-Warn "Banner style not changed (the page keeps the default title area): $($_.Exception.Message)"
        }

        $webPartId = Get-WebPartId $ManifestPath
        $component = Get-PnPPageComponent -Page $pageName -ListAvailable |
            Where-Object {
                ($webPartId -and ([string]$_.Id).Trim('{', '}') -eq $webPartId) -or
                $_.Name -eq 'GoodsReceivingWebPart' -or $_.Name -eq 'Goods Receiving'
            } |
            Select-Object -First 1

        if (-not $component) {
            Remove-PnPPage -Identity $pageName -Force
            Write-Warn 'The Goods Receiving web part is not available on this site yet, so the page was not created.'
            Write-Warn 'Deploy goods-receiving.sppkg to the App Catalog (DEPLOYMENT.md, section 4), then run this script again.'
            $pageResult = 'missing-webpart'
        } else {
            # Banner = the page's title area. Below it: one section with intro text and the web part.
            Add-PnPPageSection -Page $pageName -SectionTemplate OneColumn -Order 1 | Out-Null
            $introText = Get-Value $pageConfig 'introText'
            if ($introText) {
                Add-PnPPageTextPart -Page $pageName -Section 1 -Column 1 -Order 1 -Text "<p>$(ConvertTo-XmlText $introText)</p>" | Out-Null
            }
            Add-PnPPageWebPart -Page $pageName -Component $component -Section 1 -Column 1 -Order 2 | Out-Null
            Set-PnPPage -Identity $pageName -CommentsEnabled:$false | Out-Null
            Set-PnPPage -Identity $pageName -Publish | Out-Null
            Write-Created "Page '$pageTitle' with the Goods Receiving web part (published)"
            $pageResult = 'created'
        }
    }

    if ($pageResult -ne 'missing-webpart' -and (Get-Value $pageConfig 'addToNavigation' $true)) {
        $pageUrl = "SitePages/$pageName.aspx"
        $existingNode = Get-PnPNavigationNode -Location QuickLaunch |
            Where-Object { $_.Title -eq $pageTitle -or ([string]$_.Url).EndsWith($pageUrl) } |
            Select-Object -First 1
        if ($existingNode) {
            Write-Exists "Navigation link '$pageTitle'"
        } else {
            Add-PnPNavigationNode -Location QuickLaunch -Title $pageTitle -Url $pageUrl -First | Out-Null
            Write-Created "Navigation link '$pageTitle' (top of the site navigation)"
        }
    }
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Step 'Finished'
$pageStep = switch ($pageResult) {
    'missing-webpart' { "Deploy goods-receiving.sppkg (DEPLOYMENT.md, section 4) and run this script again to create the page." }
    'skipped' { "Add the Goods Receiving web part to a page (DEPLOYMENT.md, section 6)." }
    default { "Open the page: $($SiteUrl.TrimEnd('/'))/SitePages/$($pageConfig.name).aspx" }
}
Write-Host @"
    Next steps:
      1. Add people to '$($solution.groups.receivers)' and '$($solution.groups.supervisors)'
         (Site settings > People and groups). Add people directly or through a security group.
      2. Check Site settings > Regional settings: the time zone must be the plant's time zone.
      3. $pageStep
"@
