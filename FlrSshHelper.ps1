<#     
.NOTES 
#=========================================================================== 
# Script: 
# Author: AdamC
# Purpose: Use SSH to gather helpful info from Linux FLR Appliance
#=========================================================================== 
.DESCRIPTION 

#> 

Add-PSSnapin VeeamPSSnapin;
Add-Type -Path "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.SSH.dll";
Add-Type -Path "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Core.dll";

#$flrAppliance = Read-Host("Enter the IP assigned to the FLR appliance: ")
$flrAppliance = ""

[Veeam.Backup.SSH.CSshConnection.Protocol]::CreateCachedConnection()
# Veeam.Backup.SSH.CSshConnection.Protocol 