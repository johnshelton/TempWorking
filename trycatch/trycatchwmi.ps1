$servers = get-adcomputer -Filter {OperatingSystem -like "*Windows*Server*" -AND Enabled -eq $True}
ForEach ($srv in $servers){
  $srvname = $srv.Name
  # Write-Progress -Activity "Searching Servers for Logged On Users" -Status "Progress -- Checking $srvname" -PercentComplete (($Progress/$ServerCount)*100)
  if (test-connection -computername $srv.name -Count 1 -Quiet -ErrorAction SilentlyContinue){
    try{
      $TempLogonSessions = Get-WmiObject Win32_LogonSession -ComputerName $srv.name -ErrorAction Stop | Where-Object {($_.LogonType -eq '10') -or ($_.LogonType -eq '2')}
      $Srvname | Out-File c:\temp\20170608_serverswithsessions.txt -Append
    }
    catch{
      Write-Host "$Srvname is unreachable by wmi"
      $srvname | out-file c:\temp\20170608_serverswminotworking.txt -Append
    }
  }
  else {
    Write-Host "$Srvname is offline"
    $srvname | out-file c:\temp\20170608_serversoffline.txt -Append
  }
}