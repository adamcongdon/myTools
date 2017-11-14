# by Adam.Congdon
# this tool should remove all backups from configuration database but skip replicas
# see here for more detail: https://helpcenter.veeam.com/docs/backup/powershell/remove-vbrrestorepoint.html?ver=95
# Use at your own risk - I tested with success in my lab.

Add-PSSnapin VeeamPSSnapin
$backup = Get-VBRBackup

$restorepoints = Get-VBRRestorePoint
foreach($point in $restorepoints)
{
    #write-host $point.backupid
    if ($backup.id -notcontains $point.backupid)
    {
        
    }
    else
    {
        Remove-VBRRestorePoint $point -Confirm:$false
    }
}

