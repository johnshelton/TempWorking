param(
    [string[]] $OU = @(
                        'OU=Information Workers,OU=Employees,DC=wfm,DC=wegmans,DC=com',
                        'OU=Non-Employees,DC=wfm,DC=wegmans,DC=com'
                      ),
                           
    [ValidateScript({ $_ -le [TimeSpan]::FromDays(90) })]
    [TimeSpan] $TimeSpan = [TimeSpan]::FromDays(90),  
    [Switch] $WhatIf     = $false
)


begin {
    $ErrorActionPreference     = 'Stop'
    $Env:ADPS_LoadDefaultDrive = 0
    
    Import-Module ActiveDirectory
    
    trap {
        Send-MailMessage -SmtpServer 'smtp.wegmans.com' -From 'fim@wegmans.com' -To 'windows.support@wegmans.com' -Subject "Job Failure: $($MyInvocation.MyCommand.Name)" -Body ($_ | Out-String -Width 16384) -Priority High
    }
}

process {
    try {
        $users = $OU |% { Write-Host $_; Search-ADAccount -AccountInactive -TimeSpan $TimeSpan -UsersOnly -SearchBase $_  } |? { $_.LastLogonDate -and $_.Enabled }
    } catch {
        Send-MailMessage -SmtpServer 'smtp.wegmans.com' -From 'fim@wegmans.com' -To 'windows.support@wegmans.com' -Subject "Job Failure: $($MyInvocation.MyCommand.Name)" -Body ($_ | Out-String -Width 16384) -Priority High
        throw
    }
    $timestamp = get-date -Format yyyyMMdd_HHmmss
    $rootpath = "c:\temp\disable-inactiveaccounts\"
    $csvfilename = "disable-inactiveadaccounts_log.csv"
    $path = $rootpath + $csvfilename
    $excelfilename = "_disable-inactiveadaccounts_log.xlsx"
    $excelpath = $rootpath + $timestamp + $excelfilename
    $PathExists = Test-Path $rootpath
    IF($PathExists -eq $False){
        New-Item -Path $rootpath -ItemType  Directory
    }
    $users | Measure-Object | Format-List -Property Count
    Write-Host "Users that need to be disabled"
    $users | Format-List -Property *
    $UsersToBeDisabled = @()
    # $UsersToBeDisabledwithDateDisabled = @()
    $UsersThatHaveAlreadyBeenDisabled = @()
    $UsersToBeDisabledWithNonExpiredAccount = @()
    $UsersDisabled = @()
    $UsersPreviouslyDisabled = Import-csv $path
    $CountUsers = $Users.count
    <#
    ForEach ($User in $Users) {
        $TempUserswithDateDisabled = New-Object PSObject
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "AccountExpirationDate" -Value $User.AccountExpirationDate
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $User.DistinguishedName
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "LastLogonDate" -Value $User.LastLogonDate
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "Name" -Value $User.Name
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $User.SamAccountName
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "SID" -Value $User.SID.Value       
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $User.UserPrincipalName      
        $TempUserswithDateDisabled | Add-Member -MemberType NoteProperty -Name "Disabledate" -value (Get-Date -Format d)
        $UsersToBeDisabledwithDateDisabled += $TempUserswithDateDisabled
    }
    #>
    Clear-Host
    Write-Host "Total users to be disabled" $CountUsers
    Write-Host "Count of users that have previously been disabled" $UsersPreviouslyDisabled.Count
    ForEach ($User in $Users){
        Write-Host "Begining processing" $User.Name        
        IF ($UsersPreviouslyDisabled.Sid -contains $User.SID) {
            $TempUser = New-Object PSObject
            Write-Host "User was previously disabled" $User.Name -ForegroundColor RED -BackgroundColor White
            # $UsersPreviouslyDisabled | Where-Object {$_.Sid -contains $UsersToBeDisabled.Sid}
            $PreviouslyDisabledDate = $UsersPreviouslyDisabled | Where-Object {$User.SID -eq $_.SID} | Select DisableDate
            $UserInfo = Get-ADuser $User.Sid -Properties *
            $TempUser = $User
            $TempUser | Add-Member -MemberType NoteProperty -Name "PreviouslyDisabledDate" -Value $PreviouslyDisabledDate.Disabledate -Force
            $TempUser | Add-Member -MemberType NoteProperty -Name "Manager" -Value $UserInfo.Manager -Force
            $UsersThatHaveAlreadyBeenDisabled += $TempUser
        }
        Else {
            $Tempuser = New-Object PSObject
            Write-Host "User was NOT previously disabled" $User.Name
            $UserInfo = Get-ADuser $User.Sid -Properties *
            $TempUser = $User
            $TempUser | Add-Member -MemberType NoteProperty -Name "Manager" -Value $UserInfo.Manager -Force
            $TempUser | Add-Member -MemberType NoteProperty -Name "AccountExpirationDate" -Value $UserInfo.AccountExpirationDate -Force
            $TempUser | Add-Member -MemberType NoteProperty -Name "DateDisabled" -Value $timestamp -Force
            $UsersToBeDisabled += $TempUser
        }
    }
    # $UsersToBeDisabled | export-csv $path -Append -NoTypeInformation
    $CountUsersThatHaveAlreadyBeenDisabled = $UsersThatHaveAlreadyBeenDisabled.count
    $CountUsersToBeDisabled = $UsersToBeDisabled.count
    ForEach ($UserToBeDisabled in $UsersToBeDisabled){
        $Date = Get-Date
        IF($UserToBeDisabled.AccountExpirationDate -lt $Date -or (!($UserToBeDisabled.AccountExpirationDate))){
            # Disable-ADAccount -Identity $UserToBeDisabled.SamAccountName -ErrorAction 'Continue' -WhatIf:$WhatIf
            $UsersDisabled += $UserToBeDisabled
        }
        Else{
            $UsersToBeDisabledWithNonExpiredAccount += $UserToBeDisabled
        }
    }
    $CountUsersDisabled = $UsersDisabled.Count
    $CountUsersToBeDisabledWithNonExpiredAccount = $UsersToBeDisabledWithNonExpiredAccount.Count
    IF($CountUsersDisabled -gt 0) {
        Write-Host "Count of users disabled is" $CountUsersToBeDisabled
        $UsersDisabled | Export-Excel -path $excelpath -WorkSheetname "UsersDisabled" -TableName "TBLUsersDisabled" -AutoSize
    }
    IF($CountUsersToBeDisabledWithNonExpiredAccount -gt 0){
        Write-Host "Count of users that should be disabled but account has not expired" $CountUsersToBeDisabledWithNonExpiredAccount
        $UsersToBeDisabledWithNonExpiredAccount | Export-Excel -path $excelpath -WorkSheetname "UsersToBeDisabledWithNonExpiredAccount" -TableName "TBLUsersToBeDisabledWithNonExpiredAccounts" -AutoSize
        Send-MailMessage -SmtpServer 'smtp.wegmans.com' -From 'fim@wegmans.com' -To 'john.shelton@wegmans.com' -Subject "$($MyInvocation.MyCommand.Name) - $CountUsersToBeDisabledWithNonExpiredAccounts accounts should be disabled but their account is not expired yet." -Body ("Please see report attached. `n There were $CountUsersToBeDisabledWithNonExpiredAccountCountUsers users that should be disabled however their account has not expired yet.") -Attachments $excelpath -Priority High
    }
    IF($CountUsersThatHaveAlreadyBeenDisabled -gt 0){
        Write-Host "Count of users previously disabled is" $CountUsersThatHaveAlreadyBeenDisabled
        $UsersThatHaveAlreadyBeenDisabled | Export-Excel -path $excelpath -WorkSheetname "UsersThatHaveAlreadyBeenDisabled" -TableName "TBLUsersThatHaveAlreadyBeenDisabled" -AutoSize
        Send-MailMessage -SmtpServer 'smtp.wegmans.com' -From 'fim@wegmans.com' -To 'john.shelton@wegmans.com' -Subject "$($MyInvocation.MyCommand.Name) - $CountUsersThatHaveAlreadyBeenDisabled were previosly disabled." -Body ("Please see report attached. There were $CountUsers users that should be disabled however $CountUsersThatHaveAlreadyBeenDisabled users had previously been disabled.`n There were $CountUsersToBeDisabled accounts that were newly disabled.") -Attachments $excelpath -Priority High
    }
    $UsersDisabled | Export-CSV $path -Append -NoTypeInformation
}