# Parsing param 
param($cell_ranges_with_row)
# $target_file_directory = '', 
#, $cell_ranges_with_row = @())


# Write-Output $dict.cell_range
# $cell_ranges_with_row = (Get-Content $cell_ranges_with_row | Out-String | ConvertFrom-Json)


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open("C:\Users\gohja\Desktop\excel_data_automation\data_calculations.xlsm")

$excel.Visible = $true

foreach ($cell_object in $cell_ranges_with_row) {

 $cell_select = $cell_object.cell_select
 $cell_range = $cell_object.cell_range

 # $excel_module = $workbook.VBProject.VBComponents.Add(1)

 # $code = @"
 # Sub average_capacitance(cell_select, cell_range)
 #     Range(cell_select).Select
 #     ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
 #     Range(cell_select).Select
 #     Selection.AutoFill Destination:=Range(cell_range), Type:=xlFillDefault
 #     Range(cell_range).Select
 # End Sub
 # "@


 # # To add VBA sript into Excel macro
 # $excel_module.CodeModule.AddFromString($code)

 # To run the VBA script (Somehow wouldn't work if the .Run() method is called in the same execution as the .AddFromString() method)
 $excel.Run("average_capacitance", $cell_select, $cell_range)
}

$workbook.Save()
