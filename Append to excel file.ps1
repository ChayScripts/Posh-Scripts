<#
.SYNOPSIS
Append to excel file

.DESCRIPTION
Execute the given command or task and add the output to the same excel file

.PARAMETER
.\Append_excel_file_with_output.ps1 
No Parameters needed. 

.PREREQUISITE
Excel 2007 or higher. if you dont have this excel version, change the xlsx to xls in the script.

.EXAMPLE
Edit the directory path and execute the script as mentioned in parameter part or double click it. 
I used the directory at multiple places in the script. So, I took it into a variable at the top of the script so that if I have to change the directory path, I just edit the first line in the script.

.NOTES
This script pings hostnames from file and writes result to .xlsx file. 

.FUNCTIONALITY
Basic idea of this script is to append to excel file. you can also use it for other commands as per your requirement.
what ever is the command, output will be written to excel file and will be appended from next run. 
To perform the task repeatedly, use scheduled task. It would append to the same excel file.

#>


#get hostnames file
$directory = "c:\general"
$input=Get-Content $directory\hostnames.txt

#get only date and time
$a = Get-Date
$Date = $a.ToShortDateString()
$Time = $a.ToShortTimeString()

#checks if file already exists. if there is no file, create new one with the headers given below.

if ((Test-Path $directory\output.xlsx) -eq $false) {

#create new excel object, and add headers
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$excel.DisplayAlerts = $False
$Excel.Rows.Item(1).Font.Bold = $true 
$value1= $workbook.Worksheets.Item(1)
$value1.Name = 'Sessions Count'
$value1.Cells.Item(1,1) = 'Date'
$value1.Cells.Item(1,2) = 'Time'
$value1.Cells.Item(1,3) = 'Server01'
$value1.Cells.Item(1,4) = 'Server02'
$value1.Cells.Item(1,5) = 'Server03'

$row = 2
$column = 1

#even if the server is online or not, date and time needs to be written to the file. So these two values are out of for loop.

  #Print Date 
    $value1.Cells.Item($row,$column) = $Date
    $column++
    #Print Time
    $value1.Cells.Item($row,$column) = $Time
    $column++

foreach ($_ in $input) {

if ((Test-Connection -computername $_ -Quiet) -eq $true) {
$var =  "$_ is online"
} else {
$var = "$_ not online" 
}
  
    #write the result if server is online or not
    $value1.Cells.Item($row,$column) = $var
    $column++
 }  
 
#write data to excel file and save it 
$usedRange = $value1.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null
$workbook.SaveAs("$directory\output.xlsx") 
$excel.Quit()

#release the excel object created and remove variable
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

} else {

#if file already exists, count the lines and take the count of number of lines into a variable so that we can print data to next line.

$strExcelFile = "$directory\output.xlsx"
$objExcel = New-Object -comobject Excel.Application
$objWorkbook = $objExcel.Workbooks.Open($strExcelFile)
$ExcelWorkSheet = $ObjExcel.WorkSheets.item("Sessions Count")
$ExcelWorkSheet.activate()
$lastRow = $ExcelWorkSheet.UsedRange.rows.count + 1
$objexcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel)
Remove-Variable objexcel

$row = $lastRow
$column = 1

#now the cell is determined where the output needs to be written. So, open that file and write data.
$ExcelFile = "$directory\output.xlsx"
$excelnew = New-Object -ComObject excel.application
$excelWorkbook = $excelnew.Workbooks.Open($ExcelFile)
$value1= $excelworkbook.Worksheets.Item(1)
$value1.activate()

  #Print Date 
    $value1.Cells.Item($row,$column) = $Date
    $column++
    #Print Time
    $value1.Cells.Item($row,$column) = $Time
    $column++

foreach ($_ in $input) {

if ((Test-Connection -computername $_ -Quiet) -eq $true) {
$var =  "$_ is online"
} else {
$var = "$_ not online" 
}
  
    #write the result if server is online or not
    $value1.Cells.Item($row,$column) = $var
    $column++
 }  

#save the excel file, release it and remove variable.

$usedRange = $value1.UsedRange
$excelnew.DisplayAlerts = $False
$usedRange.EntireColumn.AutoFit() | Out-Null
$excelworkbook.SaveAs("$directory\output.xlsx")
$excelnew.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelnew)
Remove-Variable excelnew
}
