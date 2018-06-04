<#     
.NOTES 
#=========================================================================== 
# Script: 
# Author: AdamC
# Purpose: Probe Nimble storage for all VeeamAUX snapshots and how many total connections are made
#=========================================================================== 
.DESCRIPTION 

#> 

#check if Posh-SSH is installed
$isModule = Get-Module -Name Posh-SSH
if($isModule -eq $null)
    {
        Try
        {
            Install-Module Posh-SSH
        }
        catch
        {
            Write-Host("Cannot install Posh-SSH Module. Quitting")
            exit
        }
    }
else
    {
        Write-Host("Posh-SHH already installed")
    }

#check if NIMBLE PS Module is available
$mod = Get-Module -ListAvailable "HPENimblePowerShellToolkit"
if($mod -eq $null)
    {
        Write-Host("HPE Nimble PS Module not found. Quitting")
        Write-Host("See here for details: https://community.hpe.com/t5/HPE-Storage-Tech-Insiders/Introducing-the-Nimble-PowerShell-Toolkit-1-0/ba-p/6986519#.WxVT50gvyUk") -ForegroundColor Yellow
        exit
    }
else
    {
        Try
            {
                $modPath = "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\HPENimblePowerShellToolkit"
                gci -Path $modPath -Recurse | Unblock-File
                
            }
        Catch
            {
                Write-Host("Failure on Unblock. Quitting")
                exit
            }
        Import-Module -Name $mod
    }


## ignore cert auth from: https://blog.ukotic.net/2017/08/15/could-not-establish-trust-relationship-for-the-ssltls-invoke-webrequest/
if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
{
$certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
    Add-Type $certCallback
 }
[ServerCertificateValidationCallback]::Ignore()


#Set Environmental Variables
 $group = Read-host("Enter the DNS or IP of the Nimble/Group to connect: ")
 $username = Read-Host("Enter the user name needed to connect to Nimble: ")
 $cred = Read-Host("Enter Password") -AsSecureString | ConvertFrom-SecureString | out-file C:\cred.txt
 $password = Get-Content C:\cred.txt | ConvertTo-SecureString
 $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$password
 

 #$date = Get-Date -format d
 #$timestamp = $date -replace "/","."
 #Get-Date >> "$computername.$timestamp.log"
 
 #set array for final output

 $finalArray = @()

 <#
 TODO:
 1. List vol
 2. Foreach Vol > List VeeamSNaps
 3. Foreach snap > get connections
 4. Output connections to array then output average and max???

 #>

 Connect-NSGroup -Group 10.0.0.50 -Credential $credentials

#1: List Vols
$volList = Get-NSVolume

$x = 0
  while($x -ne 1000)
{
        #connect!
        Connect-NSGroup -Group 10.0.0.50 -Credential $credentials
        $ssh = New-SSHSession -ComputerName $group -Credential $credentials
        $outArray = @()
        $newArray = @()
        Write-Host("Loop number $x `n") -ForegroundColor Green
#2
        
        foreach($volume in $volList.name)
            {
                $snapList = Get-NSSnapshot -vol_name $volume
                foreach($snap in $snapList.name)
                 {
#3
                    if($snap.Contains("VeeamAUX"))
                    {
                        Write-Host("Veeam Snapshot found on volume $volume : $snap") -ForegroundColor Cyan
                        $connectionOutput = $(Invoke-SSHCommand -SSHSession $ssh -Command "snap --info $snap --vol $volume | grep -i connections" )
                        $connectionOutput.Output
                        #$connectionOutput.Output -split "\s"
                        $testArray = $connectionOutput.Output -split "\s"
                        #$testArray[-1]
                        $outArray += $testArray[-1]
                     }
                     else
                     {

                     }
                 }   
            }
        foreach($unit in $outArray)
        {
            if($unit -ne "0")
            #$newArray += $unit #remove this line and uncomment below braces
            {
                $newArray += $unit
            }
        }
        $newArray | Measure-Object -Maximum -Average -Sum
        $sample = $newArray | Measure-Object -Sum
#4
        $finalArray += $sample.Sum
        Write-host("Here is the max connection count found as of loop $x") -ForegroundColor Cyan
        $x++
        $sum = $finalArray | Measure-Object -Maximum -Sum -Average -Minimum
        $Max = $sum.Maximum
        $Sum1 = $sum.Sum
        $Avg = $sum.Average
        Write-Host("Max = $Max")
        #Write-Host("Sum = $Sum1")
        #Write-Host("Avg = $Avg")
        
        Remove-SSHSession -Index 0
        Disconnect-NSGroup
        Start-Sleep 10
    }



#Disconnect SSH

Write-host("Disconnect successful")