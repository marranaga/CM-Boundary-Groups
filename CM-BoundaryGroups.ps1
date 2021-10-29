$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable 'ConfigurationManager')) {
    try {
        Import-Module (Join-Path (Split-Path $ENV:SMS_ADMIN_UI_PATH -Parent) 'ConfigurationManager.psd1') -ErrorAction Stop
    } catch {
        Throw [System.Management.Automation.ItemNotFoundException] 'Failed to locate the ConfigurationManager.psd1 file'
    }
}

if (-not ($Settings = Get-Content "$PSScriptRoot\Settings.json" | ConvertFrom-Json)) {
    Throw [System.Management.Automation.ItemNotFoundException] 'Failed to locate the Settings.json file'
}

if (-not (Get-PSDrive -Name $Settings.SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $Settings.SiteCode -PSProvider 'CMSite' -Root $Settings.ServerAddress -Description "SCCM Site" -ErrorAction Stop
}


Push-Location "$($Settings.SiteCode):"

$RefreshSchedule = New-CMSchedule -RecurInterval Days -RecurCount 1 -Start (New-Object DateTime (Get-Random ([datetime]::maxValue.ticks))).ToLongTimeString()
$FolderRoot = "$($Settings.SiteCode):\DeviceCollection\$($Settings.RootFolderPath)".Trim('\')


$FolderRoot = "$($Settings.SiteCode):\DeviceCollection\$($Settings.RootFolderPath)".Trim('\')
Write-Host ('Folder Root: {0}' -f $FolderRoot)

if (-not (Test-Path $FolderRoot)) {
    Write-Host ('Folder root does not exist. Creating {0}' -f $FolderRoot)
    New-Item -Path $FolderRoot -ItemType Directory -Force
}

$BoundaryGroups = Get-CMBoundaryGroup
Write-Host ('Processing {0} Boundary Groups' -f $BoundaryGroups.Count)

foreach ($BoundaryGroup in $BoundaryGroups) {
    $CollectionName = $Settings.Prefix + $BoundaryGroup.Name
    Write-Host ('Collection Name: {0}' -f $CollectionName)

    if (-not ($ExistingCollection = Get-CMDeviceCollection -Name $CollectionName)) {
        $NewCMDeviceCollection = @{
            Name                   = $CollectionName
            Comment                = "All systems under the boundary group: $($BoundaryGroup.Name)"
            LimitingCollectionName = $Settings.LimitingCollectionName
            RefreshSchedule        = $RefreshSchedule
            RefreshType            = 'Periodic'
        }
        $Collection = New-CMDeviceCollection @NewCMDeviceCollection
        $CollectionQuery = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.ResourceId in (select resourceid from SMS_CollectionMemberClientBaselineStatus where SMS_CollectionMemberClientBaselineStatus.boundarygroups like '%$($BoundaryGroup.Name)%') and SMS_R_System.Name not in ('Unknown') and SMS_R_System.Client = '1'"
        
        Add-CMDeviceCollectionQueryMembershipRule -Collection $Collection -QueryExpression $CollectionQuery -RuleName "Boundary Group"
        $Collection | Move-CMObject -FolderPath $FolderRoot
    }
}

$NoBGName = $Settings.Prefix + 'None'
if (-not ($ExistingCollection = Get-CMDeviceCollection -Name $NoBGName)) {
    $NewCMDeviceCollection = @{
        Name                   = $CollectionName
        Comment                = "All systems with no boundary group"
        LimitingCollectionName = $Settings.LimitingCollectionName
        RefreshSchedule        = $RefreshSchedule
        RefreshType            = 'Periodic'
    }

    $Collection = New-CMDeviceCollection @NewCMDeviceCollection

    $Collection = New-CMDeviceCollection @NewCMDeviceCollection
    $CollectionQuery = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.ResourceId in (select resourceid from SMS_CollectionMemberClientBaselineStatus where SMS_CollectionMemberClientBaselineStatus.boundarygroups is null) and SMS_R_System.Name not in ('Unknown') and SMS_R_System.Client = '1'"
    
    Add-CMDeviceCollectionQueryMembershipRule -Collection $Collection -QueryExpression $CollectionQuery -RuleName "No Boundary Group"
    $Collection | Move-CMObject -FolderPath $FolderRoot
}

Pop-Location
