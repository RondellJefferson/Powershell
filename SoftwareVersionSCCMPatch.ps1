#Creates Excel application
$excel = New-Object -ComObject excel.application

#Makes Excel Visable
$excel.Application.Visible = $true

#$excel.DisplayAlerts = $false

#Creates Excel workBook
$book = $excel.Workbooks.Add()

#Adds worksheets

#gets the work sheet and Names it
$sheet = $book.Worksheets.Item(1)
$sheet.name = 'ComputerInfo'

#Select a worksheet
$sheet.Activate() | Out-Null

#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$sheet.Cells.Item($row,$column) = "Partstown PC Info"
$sheet.Cells.Item($row,$column).Font.Name = "Calibri Light"
$sheet.Cells.Item($row,$column).Font.Size = 21
$sheet.Cells.Item($row,$column).Font.ColorIndex = 3
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 2
$sheet.Cells.Item($row,$column).HorizontalAlignment = -4108
$sheet.Cells.Item($row,$column).Font.Bold = $true
#Merge the cells
$range = $sheet.Range("A1:c1").Merge() | Out-Null

#Move to the next row
$row++
#Create Intial row so you can add borders later
$initalRow = $row
#create Headers for your sheet
$sheet.Cells.Item($row,$column) = "ComputerName"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "ESET Endpoint Version"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "Traps Version"
$sheet.Cells.Item($row,$column).Font.Size = 16
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1


$computers = Get-ADComputer -SearchBase "OU=Parts Town Computers,DC=ptown,DC=local" -Filter * | Select-Object name -ExpandProperty name
foreach ($cpu in $computers){
    Function Get-EsetInfo{
            param ($aComputer)
            #Receive the password expire date and convert it to normal time.
            $esetFind = Get-CimInstance -ComputerName $aComputer -ClassName win32_product | Select-Object name,Version | Where-Object {$_.name -eq "Eset Endpoint Antivirus" } | Select-Object version
            $esetVersion = $esetFind.version.toString()
            return $esetVersion
            }
    Function Get-trapsFind{
            param ($aComputer)
            #Receive the password expire date and convert it to normal time.
            $trapsFind = Get-CimInstance -ComputerName $aComputer -ClassName win32_product | Select-Object name,Version | Where-Object {$_.name -eq "Traps 4.1.3.33176" } | Select-Object version
            $trapsVersion = $trapsFind.version.toString()
            return $trapsVersion
            }
    Function Test-PCConnection{
            param ($aComputer)
            $test = Test-Connection -ComputerName $aComputer -BufferSize 16 -Count 1 -ea 0 -Quiet
            return $test
            }

    $sheet.Cells.Item($row,$column) = $cpu
    $runCodeOrNaw = Test-PCConnection $cpu

    if ($runCodeOrNaw -eq "True"){
        try {
            $column++
            $esetInfomation = Get-EsetInfo $cpu
            $trapsInformation = Get-trapsFind $cpu
            $sheet.Cells.Item($row,$column) = $esetInfomation
            $column++
            $sheet.Cells.Item($row,$column) = $trapsInformation
            $column++
            } catch {
                Enable-PSRemoting –force
                $sheet.Cells.Item($row,$column) = $esetInfomation
                $column++
                $sheet.Cells.Item($row,$column) = $trapsInformation
                $column++
                $sheet.Cells.Item($row,$column) = $cpu
                $column++
                    } 
        } else {
        
        $sheet.Cells.Item($row,$column).Interior.ColorIndex = 3
        
        }
        
        $row++
        $column = 1
}
#$pcUpdated = Get-HotFix | Select-Object HotFixID,Description,InstalledOn | Where-Object {$_.InstalledOn -eq "6/16/2018"}
#$aString = $pcUpdated.installedon.Equals("6/16/2018")
#$isPCUpdated = $aString.ToString()