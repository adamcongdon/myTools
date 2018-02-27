## Download folder cleaner

$dlDir = "C:\Users\adam.congdon\Downloads"
$files = gci $dlDir
$date = Get-Date
$timespan = $date.AddDays(-3)

foreach($file in $files)
{
    if($file.lastaccesstime -gt $timespan)
    {
        write-host "Success" -ForegroundColor Green
        Write-Host $file -ForegroundColor Magenta
    }
    else
    {
        write-host "fail" -ForegroundColor Red
        Write-Host $file -ForegroundColor Cyan
        Remove-Item -Path $file.PSPath
    }
}