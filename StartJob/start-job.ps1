$vms = "RDC-Shelly-01"
#
# Hash Table for Drive Type
#
$hash = @{
   "2" = "Removable disk"
   "3" = "Fixed local disk"
   "4" = "Network disk"
   "5" = "Compact disk"
   }
#
#
# Declare Empty Arrays
#
$ServerDisksInfo = @()
$ServerDiskDetailInfo = @()
$JobServerDiskDetailInfo = @()
$JobVMDiskInfo
#
ForEach -parralel ($VM in $VMs) {
    $ServerName = $args[0]
    Write-Host $ServerName
    $AllServerDiskInfo = Get-WMIObject Win32_Volume -ComputerName $ServerName
    ForEach ($TempDiskInfo in $AllServerDiskInfo) {
      $Temp = New-Object PSObject
      $TempDriveType = $TempDiskInfo.DriveType.ToString()
      $Temp | Add-Member -MemberType NoteProperty -Name Date -Value (Get-Date)
      $Temp | Add-Member -MemberType NoteProperty -Name Server -Value $ServerName
      $Temp | Add-Member -MemberType NoteProperty -Name DriveLetter -Value $TempDiskInfo.Name
      $Temp | Add-Member -MemberType NoteProperty -Name VolumeName -Value $TempDiskInfo.Label
      $Temp | Add-Member -MemberType NoteProperty -Name DriveType -Value ($Hash.Item($TempDriveType))
      $Temp | Add-Member -MemberType NoteProperty -Name FileSystem -Value $TempDiskInfo.FileSystem
      $Temp | Add-Member -MemberType NoteProperty -Name "Size(GB)" -Value ([Math]::Round($TempDiskInfo.Capacity / 1GB,2))
      $Temp | Add-Member -MemberType NoteProperty -Name "FreeSpace(GB)" -Value ([Math]::Round($TempDiskInfo.FreeSpace / 1GB,2))
      IF($TempDiskInfo.Capacity -gt 0) {$Temp | Add-Member -MemberType NoteProperty -Name "%Free" -Value ([Math]::Round(($TempDiskInfo.FreeSpace/$TempDiskInfo.Capacity)*100,2))} 
      Else {$Temp | Add-Member -MemberType NoteProperty -Name "%Free" -Value 0}
      $Temp
  }
}
#

#$ServerDiskDetailInfo | ConvertTo-HTML Date,Server,DriveLetter,VolumeName,DriveType,FileSystem,"Size(GB)","FreeSpace(GB)" -Title "Disk Info for $ServerName" -body "$HTMLHead<H2> Disk Info for $ServerName </H2> </P>" | Set-Content $OutputHTMLFile
#Invoke-Item $OutputHTMLFile