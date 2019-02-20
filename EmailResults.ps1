$MyEmail = "rjefferson@partstown.com"
$SMTP = "smtp.office365.com"
$To = "Helpdesk@partstown.com"
$Subject = "New AD Users Added"
$Body = "This is a PowerShell test"
$Creds = (Get-Credential -Credential "$MyEmail")

Start-Sleep 2

Send-MailMessage -To $To -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
