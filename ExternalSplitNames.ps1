﻿#Creates Excel application
$excel = New-Object -ComObject excel.application

#Makes Excel Visable
$excel.Application.Visible = $true

#$excel.DisplayAlerts = $false

#Open Excel workBook
$book = $excel.Workbooks.add()

#Adds worksheets

#gets the work sheet and Names it
$sheet = $book.Worksheets.Item(1)
$sheet.name = 'SplitNames'

#Select a worksheet
$sheet.Activate() | Out-Null

#Create a row and set it to Row 1
$row = 1

#Create a Column Variable and set it to column 1
$column = 1

#Add the word Information and change the Font of the word
$sheet.Cells.Item($row,$column) = "ExternalContactlist"
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
$sheet.Cells.Item($row,$column) = "Alias"
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

#Now that the headers are done we go down a row and back to column 1
$row++
$column = 1

$users = Import-Csv ExternalContacts.csv
ForEach ($u in $users){ 
    
#Function to receive the Users Full Name
        $fullname = $u.name
        $SplitName = $fullname.Split()
        $sheet.Cells.Item($row,$column) = $SplitName[0]
        $column++
        $sheet.Cells.Item($row,$column) = $splitName[1]
        $column++
        $sheet.Cells.Item($row,$column) = $SplitName[0] + "." + $SplitName[1]
        $column++
        $sheet.Cells.Item($row,$column) = $u.EmailAddress
        $column++
        $row++
        $column = 1
        
}
#Fits cells to size
$UsedRange = $sheet.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null  