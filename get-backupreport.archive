#Requires -RunAsAdministrator
#Requires -Modules PowershellGet
<#
=======================================================================================
File Name: get-netapprpt.ps1
Created on: 2017-05-31
Created with VSCode
Version 1.0
Last Updated: 
Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@wegmans.com

Purpose: Generate a report on configured NetApp Scheduled Backups and the VMWare Machines
         that are included and not.

Notes: 

Change Log:


=======================================================================================
#>
#
# Define Parameter(s)
#
param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [string[]] $VCenters = $(throw "-VCenter(s) are required.")
)
#
# Clear Host
#
Clear-Host
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_HH-mm-ss
$ExecutionHost = $env:COMPUTERNAME
$ExecutionDate = Get-Date
$ExecutionUser = $env:USERNAME
$ExecutionLog = "Script was run on $ExecutionHost by $ExecutionUser on $ExecutionStamp"
#
# Define Output Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_HH-mm-ss
$path = "http://rdc-shelly-01.wfm.wegmans.com/backupreports/"
$ArchivePath = "archive\"
$FileExt = '.html'
$CSSPath = $path + "css/backupreport.css"
$HTMLFileOutputPath = "c:\inetpub\wwwroot\test\backupreports\"
$HTMLFileOutputArchivePath = $HTMLFileOutputPath + "archive\"
$CSSHTMLPath = $HTMLFileOutputPath + "css/backupreport.css"
$ArchiveLink = $path + $ArchivePath + "ArchiveBackupReportTOC.html"
$WegmansLogoPath = "http://rdc-shelly-01.wfm.wegmans.com/backupreports/css/WegmansLogo.png"
#
$PathExists = Test-Path $HTMLFileOutputPath
IF($PathExists -eq $False)
  {
  New-Item -Path $HTMLFileOutputPath -ItemType  Directory
  }
$ArchivePathExists = Test-Path $HTMLFileOutputArchivePath
IF($ArchivePathExists -eq $False)
  {
  New-Item -Path $HTMLFileOutputArchivePath -ItemType  Directory
  }
$SearchIconPath = $path + "css\searchicon.png"
$BackupReportOutputPath = $HTMLFileOutputPath + "index" + $FileExt
[Hashtable]$VCenterAppServerPath = @{ "RDC-VMVC-01" = "\\rdc-vmvc-app-01\c`$\Program Files\NetApp\Virtual Storage Console\smvi\server\repository\scheduledBackups.xml"; "BDC-VMVC-01" = "\\bdc-vmvc-app-01\c`$\Program Files\NetApp\Virtual Storage Console\smvi\server\repository\scheduledBackups.xml"}
[Hashtable]$ReportNames = @{"RPT_BACKUPJOBDETAIL" = "NetApp Backup Job Detail"; "RPT_VMBACKUPJOBDETAIL" = "Backups by VM"; "RPT_VMSNAPMIRRORBACKUPJOBDETAIL" = "SnapMirror Backups by VM"; "RPT_VMSWITHNOBACKUPS" = "VMs with no backups found"; "RPT_VMWAREDATASTOREBACKUPJOBS" = "Backup Jobs by Datastore"; "RPT_VMWAREDATASTORESNAPMIRRORJOBS" = "SnapMirror Jobs by Datastore"; "RPT_VMWAREDATASTORENOJOBS" = "Datastores with NO Backup or SnapMirror Jobs"}
#
# Clear Variables
#
$BackupJobDetail = @()
$BackupJobDatastoreInfo = @()
$BackupJobDatastoresInfo = @()
$VMBackupDetail = @()
$VMSnapMirrorBackupDetail = @()
$VMsWithNoBackups = @()
$TempCurrentReportsHTML = @()
$TempArchiveReportsHTML = @()
$VMWareDataStoreBackupJobsDetail = @()
$VMWareDatastores = @()
$VMWareDataStoreBackupDetailReport = @()
$VMWareDataStoreSnapMirrorDetailReport = @()
$VMWareDataStoreNoJobReport = @()
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
# Start Processing through Servers
#
ForEach($VCenter in $VCenters){
  # Write-Host "Connecting to VCenter Server $VCenter"
  Connect-VIServer $VCenter
  $XMLPath = $VCenterAppServerPath.Item($VCenter)
  $TempXML = @()
  [XML]$TempXML = Get-Content -Path $XMLPath | select -Skip 1
  $VMWareDatastores += Get-Datastore -Server $VCenter
  ForEach ($BackupJob in $TempXML.root.backupJob) {
    ForEach ($TempBackupDatastore in $BackupJob.entities.entity){
      $TempDatastore = New-Object psobject
      $TempDatastore | Add-Member -MemberType NoteProperty -Name "JobName" -Value $BackupJob.jobname
      $TempDatastore | Add-Member -MemberType NoteProperty -Name "DatastoreName" -Value $TempBackupDatastore.name
      $TempDatastore | Add-Member -MemberType NoteProperty -Name "UUID" -Value $TempBackupDatastore.UUID
      $BackupJobDatastoreInfo += $TempDatastore
    }
    $TempBackupJobDetail = New-Object PSObject
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "JobName" -Value $BackupJob.jobname
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "DailyScheduleHour" -Value $BackupJob.DailySchedule.startHour
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "DailyScheduleMin" -Value $BackupJob.DailySchedule.startMinute
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "HourlyScheduleHour" -Value $BackupJob.HourlySchedule.startHour
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "HourlyScheduleMin" -Value $BackupJob.HourlySchedule.startMinute
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Retention" -Value $BackupJob.retention.count
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "NotificationEmailAddress" -Value $BackupJob.notification.addresses.address
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "NotificationType" -Value $BackupJob.Notification.type
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "No_SnapshotVMs" -Value $BackupJob.noVmSnaps
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "SnapMirror" -Value $BackupJob.updateMirror
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "UpdateSnapVault" -Value $BackupJob.updateSnapVault
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "JobStatus" -Value $BackupJob.jobState
    $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "IncludeIndependentDisks" -Value $BackupJob.includeIndependentDisks
    $BackupJobDetail += $TempBackupJobDetail
  }
  $VMs = Get-VM
  $CountVMs = $VMs.count
  $VMsProcessed = 0
  ForEach ($VM in $VMs){
    $VMsProcessed++
    $VMNetAppBackupJobs = @()
    $VMDatastores = @()
    $PercentVMs = ($VMsProcessed/$CountVMs*100)
    $VMName = $VM.Name
    Write-Progress -Activity "Processing through all $CountVMs VMs on $VCenter" -PercentComplete $PercentVMs -CurrentOperation "Processing $VMName"
    $VMDatastores = Get-Datastore -RelatedObject $VM
    $TempDatastoreCount = $VMDatastores.count
    # Write-Host "$VM has the following datastores $VmDatastores"
    ForEach ($VMDatastore in $VMDatastores){
      $VMNetAppBackupJobs += $BackupJobDatastoreInfo | Where-Object {($_.DatastoreName -eq $VMDatastore.Name)} | Select JobName
      # Write-Host $VMNetAppBackupJobs
    }
    IF(!($VMNetAppBackupJobs)){
      IF($VMDatastores.Name -ne "srm_cdot_placeholder"){
        $TempVMsWithNoBackups = New-Object psobject
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_VCenter" -Value $VCenter
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_Name" -Value $VM.Name
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_PowerState" -Value $VM.PowerState
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_Datastore" -Value $VMDatastores.Name
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_BackupJobName" -Value "No Backup Jobs Found"
        $VMsWithNoBackups += $TempVMsWithNoBackups
      }
    }
    Else {
      # Write-Host "$VM has a backup job"
      ForEach ($VMNetAppBackupJob in $VMNetAppBackupJobs){
        $VMNetAppBackupJobDetail = ""
        $VMNetAppBackupJobDetail = $BackupJobDetail | Where-Object {($_.JobName -eq $VMNetAppBackupJob.JobName)} | Select JobName, DailyScheduleHour, DailyScheduleMin, HourlyScheduleHour, HourlyScheduleMin, Retention, SnapMirror, JobStatus
        # Write-Host "$VM is backed up by $VMNetAppBackupJobDetail.JobName"
        # IF(($VMNetAppBackupJobDetail)){
          IF($VMNetAppBackupJobDetail.SnapMirror -eq "true" -or $VMNetAppBackupJobDetail.SnapMirror -eq $true){
            # Write-Host ("$VM has a SnapMirror")
            $TempVMDetail = New-Object psobject
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_VCenter" -Value $VCenter
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_Name" -Value $Vm.Name
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_PowerState" -Value $Vm.PowerState
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_Datastore" -Value $VMDatastores.Name
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobName" -Value $VMNetAppBackupJobDetail.JobName
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobHourlyScheduleHour" -Value $VMNetAppBackupJobDetail.HourlyScheduleHour
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobHourlyScheduleMin" -Value $VMNetAppBackupJobDetail.HourlyScheduleMin
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobRetention" -Value $VMNetAppBackupJobDetail.Retention
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobState" -Value $VMNetAppBackupJobDetail.jobStatus
            $VMSnapMirrorBackupDetail += $TempVMDetail
          }
          Else {
            # Write-host ("$VM has a Backup")
            $TempVMDetail = New-Object psobject
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_VCenter" -Value $VCenter
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_Name" -Value $Vm.Name
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_PowerState" -Value $Vm.PowerState
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_Datastore" -Value $VMDatastores.Name
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobName" -Value $VMNetAppBackupJobDetail.JobName
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobDailyScheduleHour" -Value $VMNetAppBackupJobDetail.DailyScheduleHour
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobDailyScheduleMin" -Value $VMNetAppBackupJobDetail.DailyScheduleMin
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobRetention" -Value $VMNetAppBackupJobDetail.Retention
            $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobState" -Value $VMNetAppBackupJobDetail.jobStatus
            $VMBackupDetail += $TempVMDetail       
          }
        #}
      }
    }
  }
  #
  # BackupJobByDatastore
  #
  Foreach ($TempBackupJob in $TempXML.root.backupJob){
    $TempBackupJobName = $TempBackupJob.jobname
    $TempBackupDatastores = $TempBackupJob.entities.entity
    Foreach ($TempBackupDatastore in $TempBackupDatastores) {
      IF($TempBackupJob.updateMirror -eq $True) {$TempBackupType = "SnapMirror"} Else {$TempBackupType = "Backup"}
      $TempBackupJobDatastoreInfo = New-Object psobject
      $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "JobName" -Value $TempBackupJobName
      $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $TempBackupDatastore.Name
      $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "UUID" -Value $TempBackupDatastore.uuid
      $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "BackupType" -Value $TempBackupType
      $BackupJobDatastoresInfo += $TempBackupJobDatastoreInfo
    }
  }
}
#
# Create VMWare Datastore Backup Jobs Data
#
ForEach ($VMWareDatastore in $VMWareDatastores){
  $TempVMWareDataStoreBackupJobs = $BackupJobDatastoresInfo | Where-Object {$_.Datastore -match $VMWareDatastore.Name}
  ForEach ($TempVMWareDataStoreBackupJob in $TempVMWareDataStoreBackupJobs){
    $TempBackupJobData = $BackupJobDetail | Where-Object ($_.JobName -match $TempVMWareDataStoreBackupJob.JobName)
    $TempVMWareDataStoreBackupJobDetail = New-Object psobject
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_Name" -Value $VMWareDatastore.Name
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_FreeSpaceGB" -Value $VMWareDatastore.FreeSpaceGB
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_CapacityGB" -Value $VMWareDatastore.CapacityGB
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_JobName" -Value $TempVMWareDataStoreBackupJob.JobName
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_UUID" -Value $TempVMWareDataStoreBackupJob.UUID
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "BackupType" -Value $TempVMWareDataStoreBackupJob.BackupType
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "BackupRetention" -Value $TempBackupJobData.Retention
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_NoSnapshotVMs" -Value $TempBackupJobData.No_SnapshotVMs
    $VMWareDataStoreBackupJobsDetail += $TempVMWareDataStoreBackupJobDetail
  }
  IF(!($TempVMWareDataStoreBackupJobs)){
    $TempVMWareDataStoreBackupJobDetail = New-Object psobject
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_Name" -Value $VMWareDatastore.Name
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_FreeSpaceGB" -Value "NA"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "VMWare_Datastore_CapacityGB" -Value "NA"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_JobName" -Value "*** No BackupJob Found ****"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_UUID" -Value "NA"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "BackupType" -Value "NA"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "BackupRetention" -Value "NA"
    $TempVMWareDataStoreBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Backup_NoSnapshotVMs" -Value "NA"
    $VMWareDataStoreBackupJobsDetail += $TempVMWareDataStoreBackupJobDetail    
  }
}
#
# Create Separate Datastore Reports for Backup vs SnapMirror
#
ForEach ($VMWareDataStoreBackupJobDetail in $VMWareDataStoreBackupJobsDetail){
  IF($VMWareDataStoreBackupJobDetail.VMWare_Datastore_Name -ne "srm_cdot_placeholder"){
    IF($VMWareDataStoreBackupJobDetail.BackupType -eq "Backup"){
      $VMWareDataStoreBackupDetailReport += $VMWareDataStoreBackupJobDetail
    }
    Else {
      IF($VMWareDataStoreBackupJobDetail.BackupType -eq "SnapMirror"){
        $VMWareDataStoreSnapMirrorDetailReport += $VMWareDataStoreBackupJobDetail
      }
      ELSE{
        $VMWareDataStoreNoJobReport += $VMWareDataStoreBackupJobDetail
      }
    }
  }
}
#
# Archive the Current Reports
#
$PreviousReportFileNameVariable = $HTMLFileOutputPath + "*.html"
$PreviousReports = Get-ChildItem $PreviousReportFileNameVariable
$TempCurrentRptLogFile = Get-ChildItem $HTMLFileOutputPath | Where-Object {$_.Name -like "*.log"}
$TempCurrentRptDate = $TempCurrentRptLogFile.Name.Split(".log")[0]
$TempNewArchivePath = $HTMLFileOutputArchivePath + $TempCurrentRptDate + "\"
New-Item -Path $TempNewArchivePath -ItemType Directory
ForEach ($PreviousReport in $PreviousReports){
  # $TempArchiveFileName = $TempCurrentRptDate + "_" + $PreviousReport.name
  $TempCurrentReportFilePath = $HTMLFileOutputPath + $PreviousReport.Name
  # $TempArchiveFilePath = $HTMLFileOutputArchivePath + $TempArchiveFileName
  Write-Host "Archiving $TempCurrentReportFilePath to $TempNewArchivePath"
  Move-Item -Path $TempCurrentReportFilePath -Destination $TempNewArchivePath
}
#
# Archive the Current Log File
#
$TempCurrentRptLogFilePath = $HTMLFileOutputPath + $TempCurrentRptLogFile
$TempDestRptLogFilePath = $HTMLFileOutputArchivePath + $TempCurrentRptLogFile
Move-Item -Path $TempCurrentRptLogFilePath -Destination $TempNewArchivePath
#
# Generate Reports
#
IF($VCenters.count -gt 1) {$VCentersString = [system.String]::Join(", ",$VCenters)} Else {$VCentersString = $VCenters}
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmbackupjobdetail" + $FileExt
$VMBackupDetailHTML = $VMBackupDetail | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -CssUri $CSSHTMLPath -Title "VM Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2><p><input type=`"text`" id=`"myInput`" onkeyup=`"myFunction()`" placeholder=`"Search for VM names..`" title=`"Type in a VM name`"></center> <script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById(`"myInput`");  filter = input.value.toUpperCase();  table = document.getElementById(`"myTable`");  tr = table.getElementsByTagName(`"tr`");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName(`"td`")[1];    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = `"`";      } else {        tr[i].style.display = `"none`";      }    }        }}</script><center>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMBackupDetailHTML = $VMBackupDetailHTML -replace "<table>", "<table id=`"myTable`">"
$VMBackupDetailHTML = $VMBackupDetailHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMBackupDetailHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmsnapmirrorbackupjobdetail" + $FileExt
$VMSnapMirrorBackupDetailHTML = $VMSnapMirrorBackupDetail | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "VM SnapMirror Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>VM SnapMirror Backup Job Report for $VCentersString as of $ExecutionStamp</H2><p><input type=`"text`" id=`"myInput`" onkeyup=`"myFunction()`" placeholder=`"Search for VM names..`" title=`"Type in a VM name`"></center> <script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById(`"myInput`");  filter = input.value.toUpperCase();  table = document.getElementById(`"myTable`");  tr = table.getElementsByTagName(`"tr`");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName(`"td`")[1];    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = `"`";      } else {        tr[i].style.display = `"none`";      }    }        }}</script><center>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMSnapMirrorBackupDetailHTML = $VMSnapMirrorBackupDetailHTML -replace "<table>", "<table id=`"myTable`">"
$VMSnapMirrorBackupDetailHTML = $VMSnapMirrorBackupDetailHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMSnapMirrorBackupDetailHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_backupjobdetail" + $FileExt
$BackupJobDetailHTML = $BackupJobDetail | Sort-Object JobName | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "Backup Job Report as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$BackupJobDetailHTML = $BackupJobDetailHTML -replace "<table>", "<table id=`"myTable`">"
$BackupJobDetailHTML = $BackupJobDetailHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$BackupJobDetailHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmswithnobackups" + $FileExt
$VMsWithNoBackupsHTML = $VMsWithNoBackups | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "VMs With No Backup Jobs as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>VMs With No Backup Jobs as of $ExecutionStamp</H2>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMsWithNoBackupsHTML = $VMsWithNoBackupsHTML -replace "<table>", "<table id=`"myTable`">"
$VMsWithNoBackupsHTML = $VMsWithNoBackupsHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMsWithNoBackupsHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmwaredatastorebackupjobs" + $FileExt
$VMWareDatastoreBackupJobsHTML = $VMWareDataStoreBackupDetailReport | Sort-Object VMWare_Datastore_Name | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "Datastore Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>Datastore Backup Job Report for $VCentersString as of $ExecutionStamp</H2><p><input type=`"text`" id=`"myInput`" onkeyup=`"myFunction()`" placeholder=`"Search for Datastore names..`" title=`"Type in a Datastore name`"></center> <script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById(`"myInput`");  filter = input.value.toUpperCase();  table = document.getElementById(`"myTable`");  tr = table.getElementsByTagName(`"tr`");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName(`"td`")[0];    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = `"`";      } else {        tr[i].style.display = `"none`";      }    }        }}</script><center>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "<table>", "<table id=`"myTable`">"
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMWareDatastoreBackupJobsHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmwaredatastoresnapmirrorjobs" + $FileExt
$VMWareDatastoreBackupJobsHTML = $VMWareDataStoreSnapMirrorDetailReport | Sort-Object VMWare_Datastore_Name | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "Datastore SnapMirror Job Report for $VCentersString as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>Datastore SnapMirror Job Report for $VCentersString as of $ExecutionStamp</H2><p><input type=`"text`" id=`"myInput`" onkeyup=`"myFunction()`" placeholder=`"Search for Datastore names..`" title=`"Type in a Datastore name`"></center> <script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById(`"myInput`");  filter = input.value.toUpperCase();  table = document.getElementById(`"myTable`");  tr = table.getElementsByTagName(`"tr`");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName(`"td`")[0];    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = `"`";      } else {        tr[i].style.display = `"none`";      }    }        }}</script><center>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "<table>", "<table id=`"myTable`">"
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMWareDatastoreBackupJobsHTML | Out-File -FilePath $OutputPath
#
$OutputPath = $HTMLFileOutputPath + "rpt_vmwaredatastorenojobs" + $FileExt
$VMWareDatastoreBackupJobsHTML = $VMWareDataStoreNoJobReport | Sort-Object VMWare_Datastore_Name | ConvertTo-Html -CssUri $CSSHTMLPath  -Title "Datastore With No Backup or SnapMirror Job Report for $VCentersString as of $ExecutionStamp" -Body "<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H2>Datastore With No Backup or SnapMirror Job Report for $VCentersString as of $ExecutionStamp</H2><p><input type=`"text`" id=`"myInput`" onkeyup=`"myFunction()`" placeholder=`"Search for Datastore names..`" title=`"Type in a Datastore name`"></center> <script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById(`"myInput`");  filter = input.value.toUpperCase();  table = document.getElementById(`"myTable`");  tr = table.getElementsByTagName(`"tr`");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName(`"td`")[0];    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = `"`";      } else {        tr[i].style.display = `"none`";      }    }        }}</script><center>" -PostContent "<p><a href=`"$path`">Current Reports</a><p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "<table>", "<table id=`"myTable`">"
$VMWareDatastoreBackupJobsHTML = $VMWareDatastoreBackupJobsHTML -replace "file:///c:/inetpub/wwwroot/test/backupreports/css/backupreport.css", $CSSPath
$VMWareDatastoreBackupJobsHTML | Out-File -FilePath $OutputPath
#
# Create Main HTML Table Of Contents
#
$CurrentReports = Get-ChildItem $PreviousReportFileNameVariable
ForEach ($CurrentReport in $CurrentReports){
  $TempCurrentReportName = ($CurrentReport.Name.Substring(0, $CurrentReport.Name.IndexOf('.'))).ToUpper()
  $TempCurrentReportName = $ReportNames.Item($TempCurrentReportName)
  $TempCurrentReportFileName = $CurrentReport.Name
  $TempCurrentReportLink = "`n<tr><td><a href=`"$TempCurrentReportFileName`">$TempCurrentReportName</a></td></tr>"
  $TempReports = New-Object psobject
  $TempReports | Add-Member -MemberType NoteProperty -Name "Report" -Value $TempCurrentReportLink
  $TempCurrentReportsHTML += $TempReports
}
#
$TOCHTML = "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
$TOCHTML += "`n<head>"
$TOCHTML += "`n<link rel = `"stylesheet`" type = `"text/css`" href = `"$CSSPath`" />"
$TOCHTML += "`n</head><body><center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p>"
$TOCHTML += "`n<H1>Backup Report Table of Contents<br></H1><H3>As of $ExecutionDate</H3>"
$TOCHTML += "`n<table id=`"myTable`">"
$TOCHTML += "`n<colgroup><col/></colgroup>"
$TOCHTML += "`n<tr><th>Backup Reports</th></tr>"
ForEach ($TempCurrentReportHTML in $TempCurrentReportsHTML) {
  $TOCHTML += $TempCurrentReportHTML.Report
}
$TOCHTML += "`n</table>"
#
# Generate Archive TOC
#
$ArchiveReports = Get-ChildItem $HTMLFileOutputArchivePath -Directory
$ArchiveReports = $ArchiveReports | Sort-Object $_.name -Descending
ForEach ($ArchiveReport in $ArchiveReports){
  $TempArchiveReportDate = $ArchiveReport.Name.Substring(0,8).Insert(4,'-').Insert(7,'-')
  $TempArchiveReportTime = $ArchiveReport.Name.Substring(9)
  $TempArchiveReportFilePath = $ArchiveReport.Name + "/index.html"
  $TempArchiveReportLink = "`n<tr><td><a href=`"$TempArchiveReportFilePath`">Backup Reports</a></td><td>$TempArchiveReportDate</td><td>$TempArchiveReportTime</td></tr>"
  $TempReports = New-Object psobject
  $TempReports | Add-Member -MemberType NoteProperty -Name "Report" -Value $TempArchiveReportLink
  $TempArchiveReportsHTML += $TempReports
}
$ArchiveOutputPath = $HTMLFileOutputArchivePath + "ArchiveBackupReportTOC" + $FileExt
$ArchiveTOCHTML = "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
$ArchiveTOCHTML += "`n<head>"
$ArchiveTOCHTML += "`n<link rel=`"stylesheet`" type=`"text/css`" href=`"$CSSPath`" />"
$ArchiveTOCHTML += "`n</head>`n<body>`n<center>"
$ArchiveTOCHTML += "`n<center><img src=`"$WegmansLogoPath`" alt=`"Wegmans Logo`"><p><H1>Archive Backup Report Table of Contents<br></H1><H3>Last Updated $ExecutionDate</H3>"
$ArchiveTOCHTML += "`n<table id=`"myTable`">"
$ArchiveTOCHTML += "`n<colgroup><col/></colgroup>"
$ArchiveTOCHTML += "`n<tr><th>Archived Backup Report TOCs</th><th>Report Date</th><th>Report Time</th></tr>"
ForEach ($TempArchiveReportHTML in $TempArchiveReportsHTML) {
  $ArchiveTOCHTML += $TempArchiveReportHTML.Report
}
$ArchiveTOCHTML += "`n</table>"
$ArchiveTOCHTML += "`n<p><a href=`"$path`">Current Reports</a>"
$ArchiveTOCHTML += "`n<p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser .`n<\HTML>"
$ArchiveTOCHTML | Out-File -FilePath $ArchiveOutputPath
#
$TOCHTML += "`n<p><a href=`"$ArchiveLink`">Archived Reports</a>"
$TOCHTML += "`n<p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports was ran on $ExecutionHost by $ExecutionUser ."
$TOCHTML += "`n</HTML>"
$TOCHTML | Out-File -FilePath $BackupReportOutputPath
#
# Execution Log
#
$ExecutionLogFilePath = $HTMLFileOutputPath + $ExecutionStamp + ".log"
$ExecutionLog | Out-File -FilePath $ExecutionLogFilePath
#