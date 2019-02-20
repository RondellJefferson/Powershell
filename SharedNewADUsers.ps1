#Set-Location "\\prdfile01\CompanyData\IT Files\NewUsers"
$users = Import-Csv UserAD.lnk

#1 Gets my credentials and uses it for sending all emails Username = Full Email Address <ALIAS>@partstown.com .
#$password = get-content C:\cred.txt | convertto-securestring
#$Creds = new-object -typename System.Management.Automation.PSCredential -argumentlist "rjefferson@partstown.com"

#2 Gets my credentials and uses to login Exchange Server Username = Full Admin Name <USERNAME>@ptown.local .
#$password = get-content C:\cred.txt | convertto-securestring
#$MyCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist "rd22@ptown.local"

#Setup the connection to On-Premise Exchange Server.
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://PRDEXCHANGE.ptown.local/PowerShell/ -Authentication Kerberos -Credential $MyCreds

#Makes the connection to the On-Premise Exchange Server
#Import-PSSession $Session

Function Get-ADPath {
#Function will copy a user's OU Path and Return only the OU and DC Path
    param ($aUser)
    $PersonOu = Get-ADUser -Identity $aUser -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
    $mySplit = [System.Collections.Generic.List[System.Object]]$PersonOu.split(",")
    $mySplit.RemoveAt(0)
    $newPath = $mySplit -join ","
    return $newPath
    }

Function Send-EmailHelpdesk {
#Function will email the results to the Helpdesk Email Account
    $MyEmail = "helpdesk@partstown.com"
    $SMTP = "smtp.office365.com"
    $To = "helpdesk@partstown.com"
    $Subject = "New AD User Added / Powershell"
    $Body = "Finished Creating $display in Active Directory"

    Start-Sleep 2

    Send-MailMessage -To $To -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
        }


ForEach ($u in $users){
    $display = $u.First + " " + $u.Last
    

    Function Add-ADUsers{
        $myFirst = $u.CopyFirst
        $myLast = $u.CopyLast
        $Filter = "givenName -like ""*$myFirst*"" -and sn -like ""$myLast"""
        $copiedUser = Get-ADUser -Filter $Filter | Select-Object samAccountName | select -expandproperty samAccountName

        $newPwd = ConvertTo-SecureString -String "Partstown1!" -AsPlainText -Force
        $path = Get-ADPath $copiedUser
        $upn = $u.UserID + "@partstown.com"
        Add-Content .\newUserReport.txt "Started Creating $display in $path $(Get-Date)"
        New-ADUser -GivenName $u.First -Surname $u.Last -Name $display -DisplayName $display -SamAccountName $u.UserID.ToLower() -UserPrincipalName $upn.ToLower() -Path $path -Department $u.Department -Description $u.Department -ScriptPath partstown.bat -HomeDrive "H:" -HomeDirectory "\\filer02\users\$($u.UserID.ToLower())"
        $newPwd = ConvertTo-SecureString -String "Partstown1!" -AsPlainText -Force
        Set-ADAccountPassword $u.userID -NewPassword $newPwd -Reset -PassThru 
        Set-ADUser $u.userID -ChangePasswordAtLogon $true
        Enable-ADAccount -Identity $u.UserID
        Enable-Mailbox -Identity $u.UserID -Alias $u.UserID -Database DB01

#Copy an AD User GroupMembership
#Grab the User and select MemberOf Tab
        Get-ADUser -Identity $copiedUser -Properties memberof |
#Grab the object of MemberOf
        Select-Object -ExpandProperty memberof |
#Use the object and place it with the user below
        Add-ADGroupMember -Members $u.UserID
#Setup The Users Jabber Gadget function
        $u.UserID | Set-ADUser -Replace @{jabberName=($display) + " " + "($($u.Department))"}

    }
    Function Send-EmailGlobalRelay {
    #Function will email the results to the Helpdesk Email Account
        $MyEmail = "rjefferson@partstown.com"
        $SMTP = "smtp.office365.com"
        $To = "support@globalrelay.net"
        $Cc = "helpdesk@partstown.com"
        $Subject = "New User Setup ( $display )"
                        $Body = "Hello, GlobalRelay
    
    Good Morning. Can you please add the current user email mailbox $display to our achieve, enable the continuity, and filtering services? Thank you for your help. 

    Name: $display
    Email Address: $($u.UserID)@partstown.com "

        Start-Sleep 2

        Send-MailMessage -To $To -From $MyEmail -Cc $Cc -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
    }

    $todayDate = Get-Date -Format d
    if ($u.HireDate -eq $todayDate){
        Add-ADUsers
        #Send-EmailHelpdesk
        #Send-EmailGlobalRelay
        Add-Content .\newUserReport.txt "Finished Creating $display Active Directory Account $(Get-Date)"
    } else {
        Add-Content .\newUserReport.txt "$display start date has not arrived yet  $(get-date)"
        
    }
    
    }
#Disconnects the connection to the On-Premise Exchange Server.
#Remove-PSSession $Session