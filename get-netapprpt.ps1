#
# Load VMWare PSSnapin
#
Add-PSSnapin VMWare.VimAutomation.Core
Connect-VIServer "RDC-VMVC-01"
#
$TempXML = @()
[XML]$TempXML = Get-Content -Path C:\temp\netapp\scheduledBackups.xml | select -Skip 1
$BackupJobDetail = @()
$BackupJobDatastoreInfo = @()
$VMBackupDetail = @()
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
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "NotificationEmailAddress" -Value $BackupJob.Notication.addresses.address
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "NotificationType" -Value $BackupJob.Notication.type
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "No_SnapshotVMs" -Value $BackupJob.noVmSnaps
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "SnapMirror" -Value $BackupJob.updateMirror
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "UpdateSnapVault" -Value $BackupJob.updateSnapVault
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "JobStatus" -Value $BackupJob.jobState
  $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "IncludeIndependentDisks" -Value $BackupJob.includeIndependentDisks
  # $TempBackupJobDetail | Add-Member -MemberType NoteProperty -Name "Datastores" -Value $BackupJobDatastoreInfo
  $BackupJobDetail += $TempBackupJobDetail
}

$VMs = Get-VM "RDC-EPS-WEB-01"
ForEach ($VM in $VMs){
  $VMDatastores = Get-Datastore -RelatedObject $VM
  $VMNetAppBackupJobs = $BackupJobDatastoreInfo | Where-Object {($_.DatastoreName -eq $VMDatastores.Name)} | Select JobName
  ForEach ($VMNetAppBackupJob in $VMNetAppBackupJobs){
    $VMNetAppBackupJobDetail = $BackupJobDetail | Where-Object {($_.JobName -eq $VMNetAppBackupJob.JobName)} | Select JobName, DailyScheduleHour, DailyScheduleMin, HourlyScheduleHour, HourlyScheduleMin, Retention, SnapMirror, JobStatus
    $TempVMDetail = New-Object psobject
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_Name" -Value $Vm.Name
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_PowerState" -Value $Vm.PowerState
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobName" -Value $VMNetAppBackupJobDetail.JobName
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobDailyScheduleHour" -Value $VMNetAppBackupJobDetail.DailyScheduleHour
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobDailyScheduleMin" -Value $VMNetAppBackupJobDetail.DailyScheduleMin
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobHourlyScheduleHour" -Value $VMNetAppBackupJobDetail.HourlyScheduleHour
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobHourlyScheduleMin" -Value $VMNetAppBackupJobDetail.HourlyScheduleMin
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobRetention" -Value $VMNetAppBackupJobDetail.Retention
    $TempVMDetail | Add-Member -MemberType NoteProperty -Name "VM_BackupJobState" -Value $VMNetAppBackupJobDetail.jobState
    $VMBackupDetail += $TempVMDetail
  }
}




