Set-Location C:\Users\rjefferson\Documents\code
#Creates Excel application
$excel = New-Object -Com Excel.Application

#Makes Excel Visable
$excel.Visible = $True

#$excel.DisplayAlerts = $false

#Add Workbook
$book = $excel.Workbooks.Add()

#Open Excel workBook
$Workbook = $excel.Workbooks.Item(1)

#Adds worksheets
$sheet = $Workbook.Worksheets.add()
$Bastian = $Workbook.Worksheets.add()
$BTPUser = $Workbook.Worksheets.add() 
$Mindsight = $Workbook.Worksheets.add()
$AllPartstown = $Workbook.Worksheets.add()
$RDSTest = $Workbook.Worksheets.add()
$servAccounts = $Workbook.Worksheets.add()
$Remoteusers = $Workbook.Worksheets.add()

#gets the work sheet and Names it
$sheet.Name = "AdminUsers"
$Bastian.Name = "BastianUsers"
$BTPUser.Name = "BTPUsers"
$Mindsight.Name = "MindSightUsers"
$AllPartstown.Name = "PartsTownUsers"
$RDSTest.Name = "RDSTest"
$servAccounts.Name = "ServiceAccounts"
$Remoteusers.Name = "RemoteUsers"

#Select a worksheet to view, Does Not select worksheet
$sheet.Activate() | Out-Null

#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#----------------AdminUsers------------------------------#

#Add the word Information and change the Font of the word
$sheet.Cells.Item($row,$column) = "AdminUsers"
$sheet.Cells.Item($row,$column).Font.Name = "Calibri Light"
$sheet.Cells.Item($row,$column).Font.Size = 21
$sheet.Cells.Item($row,$column).Font.ColorIndex = 3
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 2
$sheet.Cells.Item($row,$column).HorizontalAlignment = -4108
$sheet.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $sheet.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$sheet.Cells.Item($row,$column) = "Full Name"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "Username"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "LastLogin"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "PasswordExpired"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Admins,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $sheet.Cells.Item($row,$column) = $theUsersName
        $column++
        $sheet.Cells.Item($row,$column) = $u
        $column++
        $sheet.Cells.Item($row,$column) = $theLoginTime
        $column++
        $sheet.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $sheet.Cells.Item($row,$column) = $theUsersName
        $column++
        $sheet.Cells.Item($row,$column) = $u
        $column++
        $sheet.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $sheet.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}

#Fits cells to size
$UsedRange = $sheet.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null

#----------------BastianUsers------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$Bastian.Cells.Item($row,$column) = "BastianUsers"
$Bastian.Cells.Item($row,$column).Font.Name = "Calibri Light"
$Bastian.Cells.Item($row,$column).Font.Size = 21
$Bastian.Cells.Item($row,$column).Font.ColorIndex = 3
$Bastian.Cells.Item($row,$column).Interior.ColorIndex = 2
$Bastian.Cells.Item($row,$column).HorizontalAlignment = -4108
$Bastian.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $Bastian.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$Bastian.Cells.Item($row,$column) = "Full Name"
$Bastian.Cells.Item($row,$column).Font.Size = 16
$Bastian.Cells.Item($row,$column).Font.ColorIndex = 1
$Bastian.Cells.Item($row,$column).Interior.ColorIndex = 48
$Bastian.Cells.Item($row,$column).Font.Bold = $true
$column++
$Bastian.Cells.Item($row,$column) = "Username"
$Bastian.Cells.Item($row,$column).Font.Size = 16
$Bastian.Cells.Item($row,$column).Font.ColorIndex = 1
$Bastian.Cells.Item($row,$column).Interior.ColorIndex = 48
$Bastian.Cells.Item($row,$column).Font.Bold = $true
$column++
$Bastian.Cells.Item($row,$column) = "LastLogin"
$Bastian.Cells.Item($row,$column).Font.Size = 16
$Bastian.Cells.Item($row,$column).Font.ColorIndex = 1
$Bastian.Cells.Item($row,$column).Interior.ColorIndex = 48
$Bastian.Cells.Item($row,$column).Font.Bold = $true
$column++
$Bastian.Cells.Item($row,$column) = "PasswordExpired"
$Bastian.Cells.Item($row,$column).Font.Size = 16
$Bastian.Cells.Item($row,$column).Font.ColorIndex = 1
$Bastian.Cells.Item($row,$column).Interior.ColorIndex = 48
$Bastian.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Bastian AutoStore,OU=Service Accounts,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $Bastian.Cells.Item($row,$column) = $theUsersName
        $column++
        $Bastian.Cells.Item($row,$column) = $u
        $column++
        $Bastian.Cells.Item($row,$column) = $theLoginTime
        $column++
        $Bastian.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $Bastian.Cells.Item($row,$column) = $theUsersName
        $column++
        $Bastian.Cells.Item($row,$column) = $u
        $column++
        $Bastian.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $Bastian.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $Bastian.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


#----------------BTPUsers------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$BTPUser.Cells.Item($row,$column) = "BTPUsers"
$BTPUser.Cells.Item($row,$column).Font.Name = "Calibri Light"
$BTPUser.Cells.Item($row,$column).Font.Size = 21
$BTPUser.Cells.Item($row,$column).Font.ColorIndex = 3
$BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 2
$BTPUser.Cells.Item($row,$column).HorizontalAlignment = -4108
$BTPUser.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $BTPUser.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$BTPUser.Cells.Item($row,$column) = "Full Name"
$BTPUser.Cells.Item($row,$column).Font.Size = 16
$BTPUser.Cells.Item($row,$column).Font.ColorIndex = 1
$BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 48
$BTPUser.Cells.Item($row,$column).Font.Bold = $true
$column++
$BTPUser.Cells.Item($row,$column) = "Username"
$BTPUser.Cells.Item($row,$column).Font.Size = 16
$BTPUser.Cells.Item($row,$column).Font.ColorIndex = 1
$BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 48
$BTPUser.Cells.Item($row,$column).Font.Bold = $true
$column++
$BTPUser.Cells.Item($row,$column) = "LastLogin"
$BTPUser.Cells.Item($row,$column).Font.Size = 16
$BTPUser.Cells.Item($row,$column).Font.ColorIndex = 1
$BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 48
$BTPUser.Cells.Item($row,$column).Font.Bold = $true
$column++
$BTPUser.Cells.Item($row,$column) = "PasswordExpired"
$BTPUser.Cells.Item($row,$column).Font.Size = 16
$BTPUser.Cells.Item($row,$column).Font.ColorIndex = 1
$BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 48
$BTPUser.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase " OU=BTP,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $BTPUser.Cells.Item($row,$column) = $theUsersName
        $column++
        $BTPUser.Cells.Item($row,$column) = $u
        $column++
        $BTPUser.Cells.Item($row,$column) = $theLoginTime
        $column++
        $BTPUser.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $BTPUser.Cells.Item($row,$column) = $theUsersName
        $column++
        $BTPUser.Cells.Item($row,$column) = $u
        $column++
        $BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $BTPUser.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $BTPUser.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


#----------------RDSTestUsers------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$RDSTest.Cells.Item($row,$column) = "RDSTestUsers"
$RDSTest.Cells.Item($row,$column).Font.Name = "Calibri Light"
$RDSTest.Cells.Item($row,$column).Font.Size = 21
$RDSTest.Cells.Item($row,$column).Font.ColorIndex = 3
$RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 2
$RDSTest.Cells.Item($row,$column).HorizontalAlignment = -4108
$RDSTest.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $RDSTest.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$RDSTest.Cells.Item($row,$column) = "Full Name"
$RDSTest.Cells.Item($row,$column).Font.Size = 16
$RDSTest.Cells.Item($row,$column).Font.ColorIndex = 1
$RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 48
$RDSTest.Cells.Item($row,$column).Font.Bold = $true
$column++
$RDSTest.Cells.Item($row,$column) = "Username"
$RDSTest.Cells.Item($row,$column).Font.Size = 16
$RDSTest.Cells.Item($row,$column).Font.ColorIndex = 1
$RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 48
$RDSTest.Cells.Item($row,$column).Font.Bold = $true
$column++
$RDSTest.Cells.Item($row,$column) = "LastLogin"
$RDSTest.Cells.Item($row,$column).Font.Size = 16
$RDSTest.Cells.Item($row,$column).Font.ColorIndex = 1
$RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 48
$RDSTest.Cells.Item($row,$column).Font.Bold = $true
$column++
$RDSTest.Cells.Item($row,$column) = "PasswordExpired"
$RDSTest.Cells.Item($row,$column).Font.Size = 16
$RDSTest.Cells.Item($row,$column).Font.ColorIndex = 1
$RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 48
$RDSTest.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=RDS Test Users,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $RDSTest.Cells.Item($row,$column) = $theUsersName
        $column++
        $RDSTest.Cells.Item($row,$column) = $u
        $column++
        $RDSTest.Cells.Item($row,$column) = $theLoginTime
        $column++
        $RDSTest.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $RDSTest.Cells.Item($row,$column) = $theUsersName
        $column++
        $RDSTest.Cells.Item($row,$column) = $u
        $column++
        $RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $RDSTest.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $RDSTest.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null

#----------------RemoteUsers------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$Remoteusers.Cells.Item($row,$column) = "RemoteUsers"
$Remoteusers.Cells.Item($row,$column).Font.Name = "Calibri Light"
$Remoteusers.Cells.Item($row,$column).Font.Size = 21
$Remoteusers.Cells.Item($row,$column).Font.ColorIndex = 3
$Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 2
$Remoteusers.Cells.Item($row,$column).HorizontalAlignment = -4108
$Remoteusers.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $Remoteusers.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$Remoteusers.Cells.Item($row,$column) = "Full Name"
$Remoteusers.Cells.Item($row,$column).Font.Size = 16
$Remoteusers.Cells.Item($row,$column).Font.ColorIndex = 1
$Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 48
$Remoteusers.Cells.Item($row,$column).Font.Bold = $true
$column++
$Remoteusers.Cells.Item($row,$column) = "Username"
$Remoteusers.Cells.Item($row,$column).Font.Size = 16
$Remoteusers.Cells.Item($row,$column).Font.ColorIndex = 1
$Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 48
$Remoteusers.Cells.Item($row,$column).Font.Bold = $true
$column++
$Remoteusers.Cells.Item($row,$column) = "LastLogin"
$Remoteusers.Cells.Item($row,$column).Font.Size = 16
$Remoteusers.Cells.Item($row,$column).Font.ColorIndex = 1
$Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 48
$Remoteusers.Cells.Item($row,$column).Font.Bold = $true
$column++
$Remoteusers.Cells.Item($row,$column) = "PasswordExpired"
$Remoteusers.Cells.Item($row,$column).Font.Size = 16
$Remoteusers.Cells.Item($row,$column).Font.ColorIndex = 1
$Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 48
$Remoteusers.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Remote Users,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $Remoteusers.Cells.Item($row,$column) = $theUsersName
        $column++
        $Remoteusers.Cells.Item($row,$column) = $u
        $column++
        $Remoteusers.Cells.Item($row,$column) = $theLoginTime
        $column++
        $Remoteusers.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $Remoteusers.Cells.Item($row,$column) = $theUsersName
        $column++
        $Remoteusers.Cells.Item($row,$column) = $u
        $column++
        $Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $Remoteusers.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $Remoteusers.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


#----------------ServiceAccounts------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$servAccounts.Cells.Item($row,$column) = "ServiceAccounts"
$servAccounts.Cells.Item($row,$column).Font.Name = "Calibri Light"
$servAccounts.Cells.Item($row,$column).Font.Size = 21
$servAccounts.Cells.Item($row,$column).Font.ColorIndex = 3
$servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 2
$servAccounts.Cells.Item($row,$column).HorizontalAlignment = -4108
$servAccounts.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $servAccounts.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$servAccounts.Cells.Item($row,$column) = "Full Name"
$servAccounts.Cells.Item($row,$column).Font.Size = 16
$servAccounts.Cells.Item($row,$column).Font.ColorIndex = 1
$servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 48
$servAccounts.Cells.Item($row,$column).Font.Bold = $true
$column++
$servAccounts.Cells.Item($row,$column) = "Username"
$servAccounts.Cells.Item($row,$column).Font.Size = 16
$servAccounts.Cells.Item($row,$column).Font.ColorIndex = 1
$servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 48
$servAccounts.Cells.Item($row,$column).Font.Bold = $true
$column++
$servAccounts.Cells.Item($row,$column) = "LastLogin"
$servAccounts.Cells.Item($row,$column).Font.Size = 16
$servAccounts.Cells.Item($row,$column).Font.ColorIndex = 1
$servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 48
$servAccounts.Cells.Item($row,$column).Font.Bold = $true
$column++
$servAccounts.Cells.Item($row,$column) = "PasswordExpired"
$servAccounts.Cells.Item($row,$column).Font.Size = 16
$servAccounts.Cells.Item($row,$column).Font.ColorIndex = 1
$servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 48
$servAccounts.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Service Accounts,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $servAccounts.Cells.Item($row,$column) = $theUsersName
        $column++
        $servAccounts.Cells.Item($row,$column) = $u
        $column++
        $servAccounts.Cells.Item($row,$column) = $theLoginTime
        $column++
        $servAccounts.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $servAccounts.Cells.Item($row,$column) = $theUsersName
        $column++
        $servAccounts.Cells.Item($row,$column) = $u
        $column++
        $servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $servAccounts.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $servAccounts.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


#----------------MindSightUsers------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$Mindsight.Cells.Item($row,$column) = "MindSightUsers"
$Mindsight.Cells.Item($row,$column).Font.Name = "Calibri Light"
$Mindsight.Cells.Item($row,$column).Font.Size = 21
$Mindsight.Cells.Item($row,$column).Font.ColorIndex = 3
$Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 2
$Mindsight.Cells.Item($row,$column).HorizontalAlignment = -4108
$Mindsight.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $Mindsight.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$Mindsight.Cells.Item($row,$column) = "Full Name"
$Mindsight.Cells.Item($row,$column).Font.Size = 16
$Mindsight.Cells.Item($row,$column).Font.ColorIndex = 1
$Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 48
$Mindsight.Cells.Item($row,$column).Font.Bold = $true
$column++
$Mindsight.Cells.Item($row,$column) = "Username"
$Mindsight.Cells.Item($row,$column).Font.Size = 16
$Mindsight.Cells.Item($row,$column).Font.ColorIndex = 1
$Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 48
$Mindsight.Cells.Item($row,$column).Font.Bold = $true
$column++
$Mindsight.Cells.Item($row,$column) = "LastLogin"
$Mindsight.Cells.Item($row,$column).Font.Size = 16
$Mindsight.Cells.Item($row,$column).Font.ColorIndex = 1
$Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 48
$Mindsight.Cells.Item($row,$column).Font.Bold = $true
$column++
$Mindsight.Cells.Item($row,$column) = "PasswordExpired"
$Mindsight.Cells.Item($row,$column).Font.Size = 16
$Mindsight.Cells.Item($row,$column).Font.ColorIndex = 1
$Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 48
$Mindsight.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Mindsight,OU=Service Accounts,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $Mindsight.Cells.Item($row,$column) = $theUsersName
        $column++
        $Mindsight.Cells.Item($row,$column) = $u
        $column++
        $Mindsight.Cells.Item($row,$column) = $theLoginTime
        $column++
        $Mindsight.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $Mindsight.Cells.Item($row,$column) = $theUsersName
        $column++
        $Mindsight.Cells.Item($row,$column) = $u
        $column++
        $Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $Mindsight.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $Mindsight.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null



#----------------AllPartstown------------------------------#
#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$AllPartstown.Cells.Item($row,$column) = "AllPartsTownUsers"
$AllPartstown.Cells.Item($row,$column).Font.Name = "Calibri Light"
$AllPartstown.Cells.Item($row,$column).Font.Size = 21
$AllPartstown.Cells.Item($row,$column).Font.ColorIndex = 3
$AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 2
$AllPartstown.Cells.Item($row,$column).HorizontalAlignment = -4108
$AllPartstown.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $AllPartstown.Range("a1:d1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$AllPartstown.Cells.Item($row,$column) = "Full Name"
$AllPartstown.Cells.Item($row,$column).Font.Size = 16
$AllPartstown.Cells.Item($row,$column).Font.ColorIndex = 1
$AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 48
$AllPartstown.Cells.Item($row,$column).Font.Bold = $true
$column++
$AllPartstown.Cells.Item($row,$column) = "Username"
$AllPartstown.Cells.Item($row,$column).Font.Size = 16
$AllPartstown.Cells.Item($row,$column).Font.ColorIndex = 1
$AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 48
$AllPartstown.Cells.Item($row,$column).Font.Bold = $true
$column++
$AllPartstown.Cells.Item($row,$column) = "LastLogin"
$AllPartstown.Cells.Item($row,$column).Font.Size = 16
$AllPartstown.Cells.Item($row,$column).Font.ColorIndex = 1
$AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 48
$AllPartstown.Cells.Item($row,$column).Font.Bold = $true
$column++
$AllPartstown.Cells.Item($row,$column) = "PasswordExpired"
$AllPartstown.Cells.Item($row,$column).Font.Size = 16
$AllPartstown.Cells.Item($row,$column).Font.ColorIndex = 1
$AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 48
$AllPartstown.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Get-ADUser -filter * -SearchBase "OU=Departments,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u
        $theExpirationDate = Get-ExpireDate $u
        $AllPartstown.Cells.Item($row,$column) = $theUsersName
        $column++
        $AllPartstown.Cells.Item($row,$column) = $u
        $column++
        $AllPartstown.Cells.Item($row,$column) = $theLoginTime
        $column++
        $AllPartstown.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $AllPartstown.Cells.Item($row,$column) = $theUsersName
        $column++
        $AllPartstown.Cells.Item($row,$column) = $u
        $column++
        $AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 3
        $column++
        $AllPartstown.Cells.Item($row,$column).Interior.ColorIndex = 3
    } finally {
        $row++
        $column = 1
    }

}
#Fits cells to size
$UsedRange = $AllPartstown.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


$book.Save()
$excel.Quit()  

 Function Send-MyEmail {
#Function will email the results to the Helpdesk Email Account
    $password = get-content C:\cred.txt | convertto-securestring
    $Creds = new-object -typename System.Management.Automation.PSCredential -argumentlist "rjefferson@partstown.com",$password

    $MyEmail = "rjefferson@partstown.com"
    $SMTP = "smtp.office365.com"
    $To = "rjefferson@partstown.com"
    $Subject = "Updated Bastin User Login"
    $Body = "The new updated Bastin Account Information"

    Start-Sleep 2

    Send-MailMessage -To $To -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Attachments "C:\Users\rjefferson\My Documents\code\ScheduledTask\BastinAccountManagement.xlsx" -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
        }

Send-MyEmail 