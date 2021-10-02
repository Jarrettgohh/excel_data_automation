
# Defining paths

$path_to_file_directory = 'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\'
$name_of_file = 'R200_D5'

$name_of_file_xls = -join ($name_of_file, '.xls')

$full_path_to_file_xls = -join ($path_to_file_directory, $name_of_file_xls)

$name_of_file_xlsx = -join ($name_of_file, '.xlsx')

$full_path_to_file_xlsx = -join ($path_to_file_directory, $name_of_file_xlsx)


# .xls can't be read -> need to be converted to .xlsx format
# .SaveAs() function will cause the new .xls file to be encrypted

# Rename-Item -Path $full_path_to_file_xls -NewName $name_of_file_xlsx


# Open excel
$excel = Open-ExcelPackage -Path $full_path_to_file_xlsx

$worksheet = $excel.Workbook.Worksheets['Single-Point Filter Task']


# Looping through each data from C28-C38 (Relevant datas are stored here)

$cell_indices = 28..37
$cell_data = @()

foreach ($i in $cell_indices) {
 
 $cell_index = -join ('C', $i)
 
 $cell_value = $worksheet.Cells[$cell_index].Value
 $cell_data = $cell_data + $cell_value
 
}

# Creating a new .xlsx file

$full_path_to_new_excel = -join ($path_to_file_directory, 'test_data_2.xlsx')

$new_excel = New-Object -ComObject excel.application
$new_excel.visible = $False

$new_excel_workbook = $new_excel.Workbooks.Add()
$new_excel_workbook.SaveAs($full_path_to_new_excel) 
$new_excel.Quit()


$new_excel_package = Open-ExcelPackage -Path $full_path_to_new_excel

$new_excel_worksheet = $new_excel_package.Workbook.Worksheets['Sheet1']


# Writing to a new excel file
$new_excel_worksheet.Cells['B6'].Value = .6645645

# $index = 0
# foreach ($i in $cell_data) {

#  $cell_index = -join ('B', $index)

#  $new_excel_worksheet.Cells[$cell_index].Value = $i

#  $index = $index + 1
# }

Close-ExcelPackage $new_excel_package


# $worksheet.Cells['A1'].Value = 'differentvalue'
# $worksheet.Cells['C37'].Value = 0.6
# # $worksheet.Cells['A60'].Value = 'Jarrett'





# %
# %

# $excel_xls.SaveAs('C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\R50_D3.xlsx', 51)


# $csv_file = Import-Csv 'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\R50_D2.csv'


# Write-Output $csv_file

#
#
#

# $excel_xlsx = Open-ExcelPackage -Path 'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\R50_D1.xlsx'
# $worksheet = $excel_xlsx.Workbook.Worksheets['Single-Point Filter Task']

# Write-Output $worksheet.Cells['C28'].Value

#
#
#

# $cell_value = $worksheet.Cells['C28'].Value
# Write-Output $cell_value

# $cell_datas = 28

# foreach ($i in $cell_datas) {
 
#  $cell_index = -join ('C', $i)
 
#  $cell_value = $worksheet.Cells['C28'].Value
#  Write-Output $cell_value
 
# }



# Close-ExcelPackage $excel