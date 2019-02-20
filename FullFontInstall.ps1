foreach($computer in (get-content "C:\Users\rjefferson\Documents\code\list.txt"))
{
$ErrorActionPreference="Stop"
TRY{
    robocopy "C:\CompanyFonts" "\\$computer\c$\Windows\Fonts" #/e /log:"C:\Users\rjefferson\Documents\Logging\$($computer)FontLogs.log"
    robocopy "C:\CompanyFonts\Fonts" "\\$computer\c$\CompanyFonts\Fonts" #/e /log:"C:\Users\rjefferson\Documents\Logging\$($computer)FontLogsRegedit.log"
    psexec.exe \\$computer reg import c:\CompanyFonts\Fonts\Fonts.reg
    
}
Catch{
Write-Warning "$error[0] on $computer"
}
}