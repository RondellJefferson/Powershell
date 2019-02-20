Set-Location C:\Users\rjefferson\Documents\code
#Creates Excel application
$excel = New-Object -ComObject excel.application

#Makes Excel Visable
$excel.Application.Visible = $true

#$excel.DisplayAlerts = $false

#Open Excel workBook
$book = $excel.Workbooks.open("C:\Users\rjefferson\My Documents\code\ScheduledTask\BastinAccountManagement.xlsx")

#Adds worksheets

#gets the work sheet and Names it
$sheet = $book.Worksheets.Item(1)
$sheet.name = 'BastianUserManagement'

#Select a worksheet
$sheet.Activate() | Out-Null

#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$sheet.Cells.Item($row,$column) = "Bastian Users"
$sheet.Cells.Item($row,$column).Font.Name = "Calibri Light"
$sheet.Cells.Item($row,$column).Font.Size = 21
$sheet.Cells.Item($row,$column).Font.ColorIndex = 3
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 2
$sheet.Cells.Item($row,$column).HorizontalAlignment = -4108
$sheet.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $sheet.Range("A1:f1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$sheet.Cells.Item($row,$column) = "First Name"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "Last Name"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "Email"
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

$users = Import-Csv BastianUsers.csv
ForEach ($u in $users){ 
    Function Get-LogonTime{
        param ($aUser)
        #Receive the user's LastDomainLogin Time.
        $whenLogonTime = Get-ADUser $u.VPNUsername -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
        $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
        return $LoginTime
        }

    Function Get-ExpireDate{
        param ($aUser)
        #Receive the password expire date and convert it to normal time.
        $expirePass = Get-ADUser $u.VPNUsername -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
        $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
        return $giveMeExpireDate
        }

    Try{
        #Function to receive the Users Full Name
        $theUsersName = Get-ADUser $u.VPNUsername -Properties name | Select-Object -ExpandProperty name
        $theLoginTime = Get-LogonTime $u.VPNUsername
        $theExpirationDate = Get-ExpireDate $u.VPNUsername
        $sheet.Cells.Item($row,$column) = $u.FirstName
        $column++
        $sheet.Cells.Item($row,$column) = $u.LastName
        $column++
        $sheet.Cells.Item($row,$column) = $u.email
        $column++
        $sheet.Cells.Item($row,$column) = $u.VPNUsername
        $column++
        $sheet.Cells.Item($row,$column) = $theLoginTime
        $column++
        $sheet.Cells.Item($row,$column) = $theExpirationDate
        $column++

    } catch {
        $sheet.Cells.Item($row,$column) = $u.FirstName
        $column++
        $sheet.Cells.Item($row,$column) = $u.LastName
        $column++
        $sheet.Cells.Item($row,$column) = $u.email
        $column++
        $sheet.Cells.Item($row,$column) = $u.VPNUsername
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

$book.Save()
$excel.Quit()

 Function Send-MyEmail {
#Function will email the results to the Helpdesk Email Account
    $myUsername = "rjefferson@partstown.com"
    $password = Get-Content "C:\Users\rjefferson\Documents\code\pass.txt" | ConvertTo-SecureString 
    $Creds = new-object -typename System.Management.Automation.PSCredential("rjefferson@partstown.com",$password)

    $MyEmail = "rjefferson@partstown.com"
    $SMTP = "smtp.office365.com"
    $To = "rjefferson@partstown.com"
    $Cc = "phoward@partstown.com"
    $Subject = "Updated Bastian User Login"
    $Body = "The new updated Bastian Account Information"

    Start-Sleep 2

    Send-MailMessage -To $To -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Attachments "C:\Users\rjefferson\My Documents\code\ScheduledTask\BastinAccountManagement.xlsx" -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
        }

Send-MyEmail 