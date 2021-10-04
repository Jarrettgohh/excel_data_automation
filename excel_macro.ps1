# start Excel
$excel = New-Object -comobject Excel.Application

# open file
# $FilePath = 'C:\Users\gohja\Desktop\excel_data_automation\data_calculations.xlsm'
$FilePath = 'C:\Users\gohja\Desktop\CV test\Test_1\data_calculations.xlsm'
$excel.Workbooks.Open($FilePath)

$excel.Visible = $true
$excel.Run('Macro_1')