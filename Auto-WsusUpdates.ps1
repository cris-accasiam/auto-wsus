<#
.SYNOPSIS
Approve, decline and delete updates in Windows Server Update Services (WSUS).

.DESCRIPTION
Decline several types of WSUS updates, such as preview, beta, superseded, language packs, drivers, old Windows versions. Also decline all updates for ARM and x86 architectures.
Approve all updates that are not declined.
Delete all declined updates.

.PARAMETER DeclineAll
Decline all updates on the WSUS server. Useful before a cleanup, when you want to save disk space.

.PARAMETER AutoDecline
Decline WSUS updates that are preview, beta, superseded, language packs, drivers. Decline all updates for ARM and x86 architectures.
Any update meant for a Windows 10 version before 22H2 will also be declined.

.PARAMETER AutoApprove
Approve all updates that have not been approved or declined, for the target group "All Computers".

.PARAMETER DeleteDeclined
Delete from WSUS all updates that are marked as declined.

.PARAMETER WsusSync
Start the sync process on the WSUS server. This parameter takes precedence over the others, except cleanup. If both WsusCleanup and WsusSync are selected, the script will first perform a cleanup and then sync.
The script will wait until the sync is completed, then process the other operations, if any.

.PARAMETER WsusCleanup
Start the cleanup process on the WSUS server. This parameter takes precedence over all others. If both WsusCleanup and WsusSync are selected, the script will first perform a cleanup and then sync.
The script will wait until the cleanup is completed, then process the other operations, if any.

.PARAMETER Server
The WSUS server. Deafult is localhost.

.PARAMETER UseSSL
If SSL is needed to connect to the WSUS server. Default is False.

.PARAMETER PortNumber
The port number used by WSUS. Default is 8530.

.EXAMPLE
Auto-WSUSUpdates.ps1 -WsusSync -AutoDecline -AutoApprove
Trigger a sync and wait for it to complete. Decline all updates that are not needed, and approve the remaining ones.

.EXAMPLE
Auto-WSUSUpdates.ps1 -AutoDecline -DeleteDeclined
Decline all updates that are not needed, and then delete them from the server.

.EXAMPLE
Auto-WSUSUpdates.ps1 -WsusSync -WsusCleanup
Start the WSUS cleanup, then trigger a sync and wait for it to complete.

.EXAMPLE
Auto-WSUSUpdates.ps1 -DeclineAll -DeleteDeclined
Mark all updates as declined, then delete them. This will remove all updates from the server.

#>
[cmdletbinding()]
Param(
    [Parameter(Position = 1)]
    [string]$Server = ([system.net.dns]::GetHostByName('localhost')).hostname,
    [Parameter(Position = 2)]
    [bool]$UseSSL = $False,
    [Parameter(Position = 3)]
    [int]$PortNumber = 8530,    
    [switch]$AutoDecline,
    [switch]$DeclineAll,
    [switch]$DeleteDeclined,
    [switch]$AutoApprove,
    [switch]$WsusSync,
    [switch]$WsusCleanup
)

$versions = @(
    'Windows 10 Version 1507',
    'Windows 10 Version 1607',
    'Windows 10 Version 1809',
    'Windows 10 Version 21H2'
)
try {
    [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null    
}
catch {
    Write-Host 'Could not load the Microsoft.UpdateServices.Administration assembly.'
    exit
}

try {
    $WsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($Server, $UseSSL, $PortNumber)
}
catch {
    Write-Host 'Could not connect to the WSUS server.'
    exit
}

if ($WsusCleanup) {
    Write-Host 'Starting the cleanup process. This will take a while.'
    $cleanupScope = new-object Microsoft.UpdateServices.Administration.CleanupScope
    $cleanupScope.CleanupLocalPublishedContentFiles = $true
    $cleanupScope.CleanupObsoleteComputers = $true
    $cleanupScope.CleanupObsoleteUpdates = $true
    $cleanupScope.CleanupUnneededContentFiles = $true
    $cleanupScope.CompressUpdates = $true
    $cleanupScope.DeclineExpiredUpdates = $true
    $cleanupScope.DeclineSupersededUpdates = $true
    $cleanupManager = $wsusServer.getCleanupManager()
    $res = $cleanupManager.PerformCleanup($cleanupScope)
    Write-Host 'Cleanup results'
    Write-Host ('Disk space freed: ' + ($res.DiskSpaceFreed/1MB) + ' MB')
    Write-Host ('Expired updates declined: ' + $res.ExpiredUpdatesDeclined)
    Write-Host ('Obsolete computers deleted: ' + $res.ObsoleteComputersDeleted)
    Write-Host ('Obsolete updates deleted: ' + $res.ObsoleteUpdatesDeleted) 	    
    Write-Host ('Superseded updates declined: ' + $res.SupersededUpdatesDeclined)
    Write-Host ('Updates with old revisions removed: ' + $res.UpdatesCompressed)
    Write-Host ''    
}

if ($WsusSync) {
    $subscription = $WsusServer.GetSubscription()
    Write-Host 'Starting syncronization'
    $subscription.StartSynchronization()    
    Start-Sleep -Seconds 60
    while ($subscription.GetSynchronizationProgress().Phase -ne 'NotProcessing')
    {
        Start-Sleep -Seconds 300
        Write-Host '.' -NoNewline
    }
    Write-Host '.'
    Write-Host 'Syncronization finished'
    Write-Host ''
}


if ($AutoDecline -or $DeclineAll) {
    $allUpdates = $WsusServer.GetUpdates()
    $totalCount = $allUpdates.count
    $declinedCount = 0
    foreach ($update in $allUpdates) {            
        # Skip updates that are already declined
        if ($update.IsDeclined) {
            continue
        }

        # DeclineAll option will decline all updates
        if ($DeclineAll) {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }

        # Decline all updates in Preview or Beta
        if (($update.Title -match "preview|beta|dev channel") -or ($update.IsBeta -eq $true)) {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }
    
        # Decline superseeded updates
        if ($update.IsSuperseded -eq $true) {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }
    
        # Decline updates for Arm64
        if ($update.Title -match "ARM64") {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }
    
        # Decline updates for x86
        if ($update.Title -match "x86-based") {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }
    
        # Decline updates for old versions of Windows 10    
        $declined = $false
        foreach ($v in $versions) {
            if ($update.Title -match $v) {
                $update.Decline()
                $declinedCount = $declinedCount + 1
                $declined = $true
                break
            }
        }
        if ($declined) {
            continue
        }

        # Decline Language packs
        if ($update.Title -match "LanguageFeatureOnDemand|Lang Pack (Language Feature) Feature On Demand|LanguageInterfacePack") {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }

        # Decline driver updates
        if ($update.Classification -match "Drivers") {
            $update.Decline()
            $declinedCount = $declinedCount + 1
            continue
        }   
    }
    Write-Host "Declined $declinedCount updates out of $totalCount."
}

# Delete all declined updates
if ($DeleteDeclined) {
    # get all declined updates
    $allDeclined = $WsusServer.GetUpdates([Microsoft.UpdateServices.Administration.ApprovedStates]::Declined, [DateTime]::MinValue, [DateTime]::MaxValue, $null, $null)
    Write-Host ('' + $allDeclined.count + ' declined updates found')
    $count = 0
    foreach ($update in $allDeclined) {
        $WsusServer.DeleteUpdate($update.Id.UpdateId.ToString())
        $count++
    }
    Write-Host "$count updates deleted."
}

# Auto-approve updates
if ($autoApprove) {    
    # Select only the updates that are marked as needed and are not already installed
    $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updatescope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved        
    $allUpdates = $WsusServer.GetUpdates($updatescope)    
    $approvalGroups = $WsusServer.GetComputerTargetGroups()
    $approvalGroup = $approvalGroups[0]
    foreach ($update in $allUpdates) {
        $update.Approve([Microsoft.UpdateServices.Administration.UpdateApprovalAction]::Install, $approvalGroup)
        $update.Refresh()
    }
    Write-Host ($allUpdates.count + " updates approved.")
}


