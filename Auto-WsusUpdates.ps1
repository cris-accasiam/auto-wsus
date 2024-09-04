<#
.SYNOPSIS
Approve, decline and delete updates in Windows Server Update Services (WSUS).

.DESCRIPTION
Decline several types of WSUS updates, such as preview, beta, superseded, language packs, drivers, old Windows versions. Also decline all updates for ARM and x86 architectures.
Approve all updates that are not declined.
Delete all declined updates.

.PARAMETER AutoDecline
Decline WSUS updates that are preview, beta, superseded, language packs, drivers. Decline all updates for ARM and x86 architectures.
Any update meant for a Windows 10 version before 22H2 will also be declined.

.PARAMETER AutoApprove
Approve all updates that have not been approved or declined, for the target group "All Computers".

.PARAMETER AutoDelete
Delete from WSUS all updates that are marked as declined.

.PARAMETER WsusSync
Start the sync process on the WSUS server. This parameter takes precedence over the others - the sync will be executed before any other operation.
The script will wait until the sync is completed, then process the other operations, if any.

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
Auto-WSUSUpdates.ps1 -AutoDecline -AutoDelete
Decline all updates that are not needed, and delete them from the server.

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
    [switch]$AutoDelete,
    [switch]$AutoApprove,
    [switch]$WsusSync
)

$versions = @(
    'Windows 10 Version 1507',
    'Windows 10 Version 1607',
    'Windows 10 Version 1809',
    'Windows 10 Version 21H2'
)

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($Server, $UseSSL, $PortNumber);

if ($WsusSync) {
    $subscription = $WsusServer.GetSubscription()
    Write-Host 'Starting syncronization'
    $subscription.StartSynchronization()    
    Start-Sleep -Seconds 60
    while ($subscription.GetSynchronizationProgress().Phase -ne 'NotProcessing')
    {
        Start-Sleep -Seconds 300
    }
    Write-Host 'Syncronization finished'
}

if ($AutoDecline) {
    $allUpdates = $WsusServer.GetUpdates()

    foreach ($update in $allUpdates) {    
        # skip updates that are already declined
        if ($update.IsDeclined) {
            continue
        }
        # Decline all updates in Preview or Beta
        if (($update.Title -match "preview|beta|dev channel") -or ($update.IsBeta -eq $true)) {
            $update.Decline()
            continue
        }
    
        # Decline superseeded updates
        if ($update.IsSuperseded -eq $true) {
            $update.Decline()
            continue
        }
    
        # Decline updates for Arm64
        if ($update.Title -match "ARM64") {
            $update.Decline()
            continue
        }
    
        # Decline updates for x86
        if ($update.Title -match "x86-based") {
            $update.Decline()
            continue
        }
    
        # Decline updates for old versions of Windows 10    
        $declined = $false
        foreach ($v in $versions) {
            if ($update.Title -match $v) {
                $update.Decline()
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
            continue
        }

        # Decline driver updates
        if ($update.Classification -match "Drivers") {
            $update.Decline()
            continue
        }   
    }
}

# Delete all declined updates
if ($AutoDelete) {
    # get all declined updates
    $allDeclined = $WsusServer.GetUpdates([Microsoft.UpdateServices.Administration.ApprovedStates]::Declined, [DateTime]::MinValue, [DateTime]::MaxValue, $null, $null)
    Write-Host ($allDeclined.count + ' declined updates found')
    $count = 0
    foreach ($update in $allDeclined) {
        $WsusServer.DeleteUpdate($update.Id.UpdateId.ToString())
        $count++
    }
    Write-Host "$count updates deleted"
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
}


