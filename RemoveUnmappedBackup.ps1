# by Adam.Congdon
# see here for more detail: https://helpcenter.veeam.com/docs/backup/powershell/
# Use at your own risk - I tested with success in my lab.

# WHAT DO?
# Removed all un-mapped backups from configuration. You can change the flag to delete from DISK if you so choose

Add-PSSnapin VeeamPSSnapin
$backups = Get-VBRBackup
$jobs = Get-VBRJob

foreach ($backup in $backups)
{
    if($backup.jobid -notin $job.id)
    {
        try
        {
            Write-Host "Removing backup: " $backup.Name -ForegroundColor Green
            Remove-VBRBackup $backup -Confirm:$false -FromDisk:$false > $null
        }
        catch
        {
            Write-Host "Failed to remove backup, please try manually." -ForegroundColor Red
        }
    }
}