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
  [string[]] $VCenters = $(throw "-VCenters is required.")
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
#
# Define Output Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_HH-mm-ss
$path = "c:\temp\netapp\"
$ArchivePath = "C:\Temp\netapp\archive\"
$FileExt = '.html'
#
$PathExists = Test-Path $path
IF($PathExists -eq $False)
  {
  New-Item -Path $path -ItemType  Directory
  }
$ArchivePathExists = Test-Path $path
IF($ArchivePathExists -eq $False)
  {
  New-Item -Path $ArchivePath -ItemType  Directory
  }
#
# HTML Format Settings
#
$HTMLFormat = "<Style>"
$HTMLFormat = $HTMLFormat + "BODY{background-color:White;}"
$HTMLFormat = $HTMLFormat + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$HTMLFormat = $HTMLFormat + "TH{border-width: 1px;padding: 5px 10px 5px 10px;border-style: solid;border-color: black;background-color:royalblue}"
$HTMLFormat = $HTMLFormat + "TD{border-width: 1px;padding: 5px 10px 5px 10px;border-style: solid;border-color: black;background-color:gainsboro}"
$HTMLFormat = $HTMLFormat + "</style>"
#
# Load VMWare PSSnapin
#
Add-PSSnapin VMWare.VimAutomation.Core
#
[Hashtable]$VCenterAppServerPath = @{ "RDC-VMVC-01" = "\\rdc-vmvc-app-01\c`$\Program Files\NetApp\Virtual Storage Console\smvi\server\repository\#scheduledBackups.xml"; "BDC-VMVC-01" = "\\bdc-vmvc-app-01\c`$\Program Files\NetApp\Virtual Storage Console\smvi\server\repository\#scheduledBackups.xml"}
#
# Clear Variables
#
$AllDatastores = @()
$BackupJobDetail = @()
$BackupJobDatastoreInfo = @()
$BackupJobDatastoresInfo = @()
$VMBackupDetail = @()
$VMSnapMirrorBackupDetail = @()
$VMsWithNoBackups = @()
$TempCurrentReportsHTML = @()
$TempArchiveReportsHTML = @()
#
ForEach($VCenter in $VCenters){
  Write-Host "Connecting to VCenter Server $VCenter"
  Connect-VIServer $VCenter
  $XMLPath = $VCenterAppServerPath.Item($VCenter)
  $TempXML = @()
  [XML]$TempXML = Get-Content -Path $XMLPath | select -Skip 1
  # $TempExcel = Import-Excel 'C:\users\techadminjxs\OneDrive - Wegmans Food Markets, Inc\Working\NetApp_SMVI\ScheduledBackups.xlsx'
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
    $PercentVMs = ($VMsProcessed/$CountVMs*100)
    $VMName = $VM.Name
    Write-Progress -Activity "Processing through all $CountVMs VMs on $VCenter" -PercentComplete $PercentVMs -CurrentOperation "Processing $VMName"
    $VMDatastores = Get-Datastore -RelatedObject $VM
    $VMNetAppBackupJobs = $BackupJobDatastoreInfo | Where-Object {($_.DatastoreName -eq $VMDatastores.Name)} | Select JobName
    ForEach ($VMNetAppBackupJob in $VMNetAppBackupJobs){
      $VMNetAppBackupJobDetail = ""
      $VMNetAppBackupJobDetail = $BackupJobDetail | Where-Object {($_.JobName -eq $VMNetAppBackupJob.JobName)} | Select JobName, DailyScheduleHour, DailyScheduleMin, HourlyScheduleHour, HourlyScheduleMin, Retention, SnapMirror, JobStatus
      IF(($VMNetAppBackupJobDetail)){
        IF($VMNetAppBackupJobDetail.SnapMirror -eq "true"){
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
      }
      Else {
        $TempVMsWithNoBackups = New-Object psobject
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_VCenter" -Value $VCenter
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_Name" -Value $VM.Name
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_PowerState" -Value $VM.PowerState
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_Datastore" -Value $VMDatastores.Name
        $TempVMsWithNoBackups | Add-Member -MemberType NoteProperty -Name "VM_BackupJobName" -Value "No Backup Jobs Found"
        $VMsWithNoBackups += $TempVMsWithNoBackups
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
    $TempBackupJobDatastoreInfo = New-Object psobject
    $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "JobName" -Value $TempBackupJobName
    $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $TempBackupDatastore.Name
    $TempBackupJobDatastoreInfo | Add-Member -MemberType NoteProperty -Name "UUID" -Value $TempBackupDatastore.uuid
    $BackupJobDatastoresInfo += $TempBackupJobDatastoreInfo
  }
}
}
#
# Archive the Current Reports
#
$PreviousReportFileNameVariable = $path + "*.html"
$PreviousReports = Get-ChildItem $PreviousReportFileNameVariable
$TempCurrentRptLogFile = Get-ChildItem $path | Where-Object {$_.Name -like "*.log"}
$TempCurrentRptDate = $TempCurrentRptLogFile.Name.Split(".log")[0]
$TempNewArchivePath = $ArchivePath + $TempCurrentRptDate + "\"
New-Item -Path $TempNewArchivePath -ItemType Directory
ForEach ($PreviousReport in $PreviousReports){
  # $TempArchiveFileName = $TempCurrentRptDate + "_" + $PreviousReport.name
  $TempCurrentReportFilePath = $path + $PreviousReport.Name
  # $TempArchiveFilePath = $ArchivePath + $TempArchiveFileName
  Write-Host "Archiving $TempCurrentReportFilePath to $TempNewArchivePath"
  Move-Item -Path $TempCurrentReportFilePath -Destination $TempNewArchivePath
}
#
# Archive the Current Log File
#
$TempCurrentRptLogFilePath = $path + $TempCurrentRptLogFile
$TempDestRptLogFilePath = $archivepath + $TempCurrentRptLogFile
Move-Item -Path $TempCurrentRptLogFilePath -Destination $TempNewArchivePath
#
# Generate Reports
#
IF($VCenters.count -gt 1) {$VCentersString = [system.String]::Join(", ",$VCenters)} Else {$VCentersString = $VCenters}
#
$OutputPath = $path + "rpt_vmbackupjobdetail" + $FileExt
$VMBackupDetail | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -Head $HTMLFormat -Title "VM Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2>" | Out-File -FilePath $OutputPath
#
#$VMBackupDetail | ConvertTo-Html -Head $HTMLFormat -Title "VM Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2>" | Out-String
#
$OutputPath = $path + "rpt_vmsnapmirrorbackupjobdetail" + $FileExt
$VMSnapMirrorBackupDetail | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -Head $HTMLFormat -Title "VM Backup Job Report for $VCentersString as of $ExecutionStamp" -Body "<H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2>" | Out-File -FilePath $OutputPath
#
$OutputPath = $path + "rpt_backupjobdetail" + $FileExt
$BackupJobDetail | Sort-Object JobName | ConvertTo-Html -Head $HTMLFormat -Title "Backup Job Report as of $ExecutionStamp" -Body "<H2>Backup Job Report for $VCentersString as of $ExecutionStamp</H2>" | Out-File -FilePath $OutputPath
#
$OutputPath = $path + "rpt_vmswithnobackups" + $FileExt
$VMsWithNoBackups | Sort-Object VM_VCenter, VM_Name | ConvertTo-Html -Head $HTMLFormat -Title "VMs With No Backup Jobs as of $ExecutionStamp" -Body "<H2>VMs With No Backup Jobs as of $ExecutionStamp</H2>" | Out-File -FilePath $OutputPath
#
# Create Main HTML Table Of Contents
#
$CurrentReports = Get-ChildItem $PreviousReportFileNameVariable
ForEach ($CurrentReport in $CurrentReports){
  $TempCurrentReportName = ($CurrentReport.Name.Substring(0, $CurrentReport.Name.IndexOf('.'))).ToUpper()
  $TempCurrentReportFileName = $CurrentReport.Name
  $TempCurrentReportLink = "<tr><td><a href=`"$TempCurrentReportFileName`">$TempCurrentReportName</a></td></tr>"
  $TempReports = New-Object psobject
  $TempReports | Add-Member -MemberType NoteProperty -Name "Report" -Value $TempCurrentReportLink
  $TempCurrentReportsHTML += $TempReports
}
#
$BackupReportOutputPath = $path + "BackupReportTOC" + $FileExt
$TOCHTML = "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
$TOCHTML += "<head>"
$TOCHTML += $HTMLFormat
$TOCHTML += "</head><body><center>"
$TOCHTML += "<H1>Backup Report Table of Contents<br></H1><H3>As of $ExecutionDate</H3>"
$TOCHTML += "<table>"
$TOCHTML += "<colgroup><col/></colgroup>"
$TOCHTML += "<tr><th>Backup Reports</th></tr>"
ForEach ($TempCurrentReportHTML in $TempCurrentReportsHTML) {
  $TOCHTML += $TempCurrentReportHTML.Report
}
$TOCHTML += "</table>"
#
# Generate Archive TOC
#
$ArchiveReports = Get-ChildItem $ArchivePath -Directory
ForEach ($ArchiveReport in $ArchiveReports){
  $TempArchiveReportName = $ArchiveReport.Name
  $TempArchiveReportFileName = $ArchiveReport.Name
  $TempArchiveReportDate = $ArchiveReport.Name.Substring(0,8).Insert(4,'-').Insert(7,'-')
  $TempArchiveReportTime = $ArchiveReport.Name.Substring(9)
  $TempArchiveReportFilePath = $ArchivePath + $ArchiveReport.Name + "\BackupReportTOC.html"
  $TempArchiveReportLink = "<tr><td><a href=`"$TempArchiveReportFilePath`">Backup Reports</a></td><td>$TempArchiveReportDate</td><td>$TempArchiveReportTime</td></tr>"
  $TempReports = New-Object psobject
  $TempReports | Add-Member -MemberType NoteProperty -Name "Report" -Value $TempArchiveReportLink
  $TempArchiveReportsHTML += $TempReports
}
$ArchiveOutputPath = $ArchivePath + "ArchiveBackupReportTOC" + $FileExt
$ArchiveTOCHTML = "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
$ArchiveTOCHTML += "<head>"
$ArchiveTOCHTML += $HTMLFormat
$ArchiveTOCHTML += "</head><body><center>"
$ArchiveTOCHTML += "<H1>Archive Backup Report Table of Contents<br></H1><H3>Last Updated $ExecutionDate</H3>"
$ArchiveTOCHTML += "<table>"
$ArchiveTOCHTML += "<colgroup><col/></colgroup>"
$ArchiveTOCHTML += "<tr><th>Archived Backup Report TOCs</th><th>Report Date</th><th>Report Time</th></tr>"
ForEach ($TempArchiveReportHTML in $TempArchiveReportsHTML) {
  $ArchiveTOCHTML += $TempArchiveReportHTML.Report
}
$ArchiveTOCHTML += "</table>"
$ArchiveTOCHTML += "<p><a href=`"$BackupReportOutputPath`">Current Reports</a>"
$ArchiveTOCHTML += "<p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports last ran on $ExecutionHost by $ExecutionUser ."
$ArchiveTOCHTML | Out-File -FilePath $ArchiveOutputPath
#
$TOCHTML += "<p><a href=`"$ArchiveOutputPath`">Archived Reports</a>"
$TOCHTML += "<p><hr>This site is maintained by TechWintel.  Created by John Shelton.  The script that created this page and the reports last ran on $ExecutionHost by $ExecutionUser ."
$TOCHTML | Out-File -FilePath $BackupReportOutputPath
#
# Execution Log
#
$ExecutionLogFilePath = $Path + $ExecutionStamp + ".log"
$ExecutionLog | Out-File -FilePath $ExecutionLogFilePath
#