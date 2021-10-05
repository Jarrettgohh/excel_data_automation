# # start Excel
# $excel = New-Object -comobject Excel.Application

# # open file
# # $FilePath = 'C:\Users\gohja\Desktop\excel_data_automation\data_calculations.xlsm'
# $FilePath = 'C:\Users\gohja\Desktop\CV test\Test_1\data_calculations.xlsm'
# $excel.Workbooks.Open($FilePath)

# $excel.Visible = $true
# $excel.Run("average_capacitance", "P12", "P12:T12")


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open("C:\Users\gohja\Desktop\excel_data_automation\data_calculations.xlsm")
# $excel_module = $workbook.VBProject.VBComponents.Add(1)

# $code = @"
# sub average_capacitance(cell_select, cell_range)
# range(cell_select).Select
# ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
# range(cell_select).Select
# Selection.AutoFill Destination:=range(cell_range), Type:=xlFillDefault
# range(cell_range).Select
# end Sub
# "@

# $code = @"

# Sub average_capacitance(cell_select, cell_range)
#     Range(cell_select).Select
#     ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
#     Range(cell_select).Select
#     Selection.AutoFill Destination:=Range(cell_range), Type:=xlFillDefault
#     Range(cell_range).Select
# End Sub
# "@

$excel.Visible = $true

# To add VBA sript into Excel macro
# $excel_module.CodeModule.AddFromString($code)

# To run the VBA script (Somehow wouldn't work if the .Run() method is called in the same execution as the .AddFromString() method)
$excel.Run("average_capacitance", "I13", "I13:M13")
$workbook.Save()