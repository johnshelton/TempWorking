<#
=======================================================================================
File Name: get-datastoreinfo.ps1
Created on: 
Created with VSCode
Version 1.0
Last Updated: 
Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@wegmans.com

Purpose:

Notes: 

Change Log:


=======================================================================================
#>
#
# Define Parameter(s)
#
param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [string[]] $VCenters = $(throw "-VCenter Server(s) is required")
)
#
# Check if VMWare Module is installed and if not install it
#
$VMWareModuleInstalledLoop = 0
While ($VMWareModuleInstalledLoop -lt "4" -and $VMWareModuleInstalled -ne $True) {
  $VMWareModuleInstalled = Get-InstalledModule -Name VMWare*
  If ($VMWareModuleInstalled) {Write-Host "VMWare Module Installed.  Continuing Script."; $VMWareModuleInstalled = $True} Else {Install-Module "VMWare.powercli" -Scope AllUsers -Force -AllowClobber; $VMWareModuleInstalled = $False }
  $VMWareModuleInstalledLoop ++
  If ($VMWareModuleInstalledLoop -ge "4") {Write-host "VMWare Module is not installed and the auto installation failed. Manual intervention is needed"; Exit}
}
#
# Define Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_HH-mm-ss
$DatastoreInfo = @()
#
# Define Output Paths
#
$path = "c:\temp\get-datastoreinfo\"
$FullFileName = $MyInvocation.MyCommand.Name
$FileName = $FullFilename.Substring(0, $FullFilename.LastIndexOf('.'))
$FileExt = '.xlsx'
$PathExists = Test-Path $path
$OutputFile = $path + $FileName + "_" + $ExecutionStamp + $FileExt
IF($PathExists -eq $False)
  {
  New-Item -Path $path -ItemType  Directory
  }
#
Clear-Host
ForEach ($VCenter in $VCenters){
  Connect-VIServer $VCenter
  $TableName = "TBL" + $VCenter
  $DatastoreInfo = Get-Datastore -Server $VCenter | Select Name, @{Name='RemoteHost';Expression={[string]::join("|",($_.RemoteHost))}}, RemotePath, FileSystemVersion, Datacenter, ParentFolder, FreeSpaceGB, CapacityGB, State, Type
  $DatastoreInfo | Export-Excel -Path $OutputFile -WorkSheetname $VCenter -TableName $TableName -AutoSize
}