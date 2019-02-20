$users = Get-ADUser -filter * -SearchBase "OU=Departments,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
$hostname = "ptdom1"
ForEach ($u in $users){ 
Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $aUser -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }
Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $aUser | Get-ADObject -Properties lastLogon -server $hostname | select-object -ExpandProperty lastLogon
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

Function Get-DaysRemaining{
        param ($aUser)
        #Receive the user's remaining days before password expires and email's the user with how many days are left.
        $newUser = Get-ExpireDate $aUser
        $newUserWhitespace = $newUser.Substring(0,9)
        $newUserFinal = $newUserWhitespace.trim()
        $todayDate = Get-Date -Format g
        $todayWhitespace = $todayDate.Substring(0,9)
        $todayFinal = $todayWhitespace.trim()
        $daysLeft = New-TimeSpan -Start $newUserFinal -End $todayFinal | Select-Object Days -ExpandProperty days
        $passwordExpiring = $daysLeft.ToString()
        $getResult = $passwordExpiring.substring(1,2) -le "15"
        return $getResult
        }
Function Get-DaysLeft{
        param ($aUser)
        #Receive the user's remaining days before password expires and email's the user with how many days are left.
        $newUser = Get-ExpireDate $aUser
        $newUserWhitespace = $newUser.Substring(0,9)
        $newUserFinal = $newUserWhitespace.trim()
        $todayDate = Get-Date -Format g
        $todayWhitespace = $todayDate.Substring(0,9)
        $todayFinal = $todayWhitespace.trim()
        $daysLeft = New-TimeSpan -Start $newUserFinal -End $todayFinal | Select-Object Days -ExpandProperty days
        $passwordExpiring = $daysLeft.ToString()
        $getNewResult = $passwordExpiring.substring(1,2)
        return $getNewResult
        }
Function Send-MyEmails{
        param ($aUser)
        $password = Get-Content "C:\Temp\SCCM365.txt" | ConvertTo-SecureString 
        $Creds = new-object -typename System.Management.Automation.PSCredential("srv-sccm@ptholding.onmicrosoft.com",$password)
        $mailUser = Get-ADUser $aUser -Properties mail | Select-Object mail -ExpandProperty mail
        $userName = Get-ADUser $aUser -Properties name | Select-Object -ExpandProperty name
        $MyEmail = "rjefferson@partstown.com"
        $SMTP = "smtp.office365.com"
        $To = $mailUser
        $Bcc = "rjefferson@partstown.com"
        $Subject = "$userName PASSWORD EXPIRES IN $(Get-DaysLeft $aUser) DAYS"
        $Body = "This is the updated report that notifies you to change your Windows/Network password as soon as possible! The Result of not changing your password will lock you out your computer and kick you out of Finesse, Syspro, and Outlook. If you need help changing your password feel free to contact your local helpdesk. Thank you and have a great rest of your day :)."

        Start-Sleep 2
        Send-MailMessage -To $To -Bcc $Bcc -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
        
        }

$daysLeftPW =  Get-DaysLeft $u
$theLoginTime = Get-LogonTime $u
$theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
$DoesPasswordExpire = Get-ADUser $u -properties passwordneverexpires | Select-Object passwordneverexpires -ExpandProperty passwordneverexpires
$userPasswordExpiring = Get-DaysRemaining $u
#Get-DaysRemaining "$theUser" 
if ($DoesPasswordExpire -eq "$TRUE"){
     write-host "$theUsersName Password will not expire."
} else {
    if ($theLoginTime -eq "12/31/1600 6:00 PM"){
            Write-Host "$theUsersName has not logged in yet." 
    }
    Elseif ($theLoginTime -ne "12/31/1600 6:00 PM" -and $userPasswordExpiring -eq $TRUE){
         
                $myEmail = Send-MyEmails $u
                $myemail
                Add-Content C:\Users\srv-sccm\Documents\code\AllAccounts\passwordReport.txt "$theUsersName password will be expiring in $daysLeftPW days."
    }  
    Elseif ($theLoginTime -ne "12/31/1600 6:00 PM" -and $userPasswordExpiring -eq $FALSE){
                Add-Content C:\Users\srv-sccm\Documents\code\AllAccounts\PasswordsNormal.txt "$theUsersName password has over $daysLeftPW days left before it expires" 
    }  

    }
}