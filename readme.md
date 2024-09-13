# Description
This script allows you to automate several WSUS-related tasks:
- Perform a manual sync
- Decline several types of WSUS updates.
- Approve all updates that are not declined.
- Delete all declined updates.

The script will decline all updates from the following categories/types:
- preview or beta
- superseded
- language packs
- drivers
- updates for ARM
- updates for x86
- updates for all Windows 10 versions before 22H2

You will need to run the script from an account with admin rights on the WSUS server.

# Parameters
## AutoDecline
Decline WSUS updates that are preview, beta, superseded, language packs, drivers. Decline all updates for ARM and x86 architectures.
Any update meant for a Windows 10 version before 22H2 will also be declined.

## AutoApprove
Approve all updates that have not been approved or declined, for the target group "All Computers".

## DeleteDeclined
Delete from WSUS all updates that are marked as declined.

## WsusSync
Start the sync process on the WSUS server. This parameter takes precedence over the others, except cleanup. If both WsusCleanup and WsusSync are selected, the script will first perform a cleanup and then sync.
The script will wait until the sync is completed, then process the other operations, if any.

## WsusCleanup
Start the cleanup process on the WSUS server. This parameter takes precedence over all others. If both WsusCleanup and WsusSync are selected, the script will first perform a cleanup and then sync.
The script will wait until the cleanup is completed, then process the other operations, if any.

## Server
The WSUS server to connect to. Deafult is localhost.

## UseSSL
If SSL is needed to connect to the WSUS server. Default is False.

## PortNumber
The port number used by WSUS. Default is 8530.

# Examples

```powershell
Auto-WSUSUpdates.ps1 -WsusSync -AutoDecline -AutoApprove
```

Trigger a sync and wait for it to complete. Afterwards, decline all updates that are not needed, and approve the remaining ones.

```powershell
Auto-WSUSUpdates.ps1 -AutoDecline -DeleteDeclined
```

Decline all updates that match the categories listed above, and delete them from the server.

```powershell
Auto-WSUSUpdates.ps1 -DeclineAll -DeleteDeclined
```

Mark all updates as declined, then delete them. This will remove all updates from the server.
