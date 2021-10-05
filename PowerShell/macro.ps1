# Parsing param 
param([string]$target_file_directory = '')

Write-Output $target_file_directory




$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open("C:\Users\gohja\Desktop\excel_data_automation\data_calculations.xlsm")


foreach ($cell_range in $cell_ranges) {
 $regex_right_case_1 = $cell_range -match ':\D{1}' # A:F -> :F
 $regex_left_case_1 = $cell_range -match '\D{1}:' # A:F -> A:
 $regex_right_case_2 = $cell_range -match ':\D{2}' # A:FA OR AA:FA -> :FA
 $regex_left_case_2 = $cell_range -match '\D{2}:' # AA:FA -> AA:


 if ($regex_right_case_1) {


  $cell_select = $cell_range -replace $matches[0], ''

 }

 if ($regex_right_case_1 -or $regex_right_case_2) {



  $cell_select = $cell_range -replace $matches[0], ''
 }

 Write-Output $cell_select
 Write-Output $cell_range


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

 # $excel.Visible = $true

 # To add VBA sript into Excel macro
 # $excel_module.CodeModule.AddFromString($code)

 # To run the VBA script (Somehow wouldn't work if the .Run() method is called in the same execution as the .AddFromString() method)
 # $excel.Run("average_capacitance", -join( $cell_select, '13'), $cell_range)
 # $workbook.Save()
}