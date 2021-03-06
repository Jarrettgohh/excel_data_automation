
# Path: to get from user input
$path_to_file_directory = #'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm'
# 'C:\Users\gohja\Desktop\CV test\PPP_69%_20nm'
'C:\Users\gohja\Desktop\CV test\Test_2'

# Get all the files present in the folder
$files = Get-ChildItem -Path $path_to_file_directory


# Initialize empty array to store file names
$file_names = @()

foreach ($file in $files) {
 # Rename all .xls files to .xlsx extension
 $path_to_file = -join ($path_to_file_directory, -join ("\", $file.Name))

 $file_name_without_extension = [System.IO.Path]::GetFileNameWithoutExtension($file)
 $file_name_xlsx = -join ( $file_name_without_extension, '.xlsx')

 Rename-Item -Path $path_to_file -NewName  $file_name_xlsx

 # Add file names to array
 $file_names = $file_names + $file_name_without_extension
}


$current_path = Get-Location

# Get name of folder containing data
Set-Location $path_to_file_directory

$target_folder = Split-Path -Path (Get-Location )-Leaf

# Creating a new .xlsx file
$name_of_new_excel_xlsx = -join ($target_folder, '_data_calculations.xlsx')

$full_path_to_new_excel = -join ($path_to_file_directory, -join ('\', $name_of_new_excel_xlsx ))

$new_excel = New-Object -ComObject excel.application
$new_excel.visible = $False

$new_excel_workbook = $new_excel.Workbooks.Add()

$xlsm = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled

# get destination path
$xlsm_extension = [System.Io.Path]::ChangeExtension($full_path_to_new_excel, 'xlsm')

$new_excel_workbook.SaveAs($xlsm_extension, $xlsm)


# Add custom macros
$excel_macro = $new_excel_workbook.VBProject.VBComponents.Add(1)

$code = @"
Sub average_capacitance(cell_select, cell_range)
    Range(cell_select).Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
    Range(cell_select).Select
    Selection.AutoFill Destination:=Range(cell_range), Type:=xlFillDefault
    Range(cell_range).Select
End Sub
"@

# To add VBA sript into Excel macro
$excel_macro.CodeModule.AddFromString($code)

# Save and close workbook
$new_excel_workbook.Save()
$new_excel_workbook.Close()

# Quit excel
$new_excel.Quit()
   
# release COM objects to prevent memory leaks:
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($new_excel_workbook)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($new_excel)
$new_excel = $new_excel_workbook = $null
# clean up:
[GC]::Collect()
[GC]::WaitForPendingFinalizers()


# CMD command to stop all excel.exe tasks: taskkill /f /im excel.exe


# Set location back to the current script path
Set-Location $current_path


# Pipe array consisting of file names and directory path to Python
# To remove the newly added .xlsx file first
[System.Collections.ArrayList]$file_names_without_extra = $file_names

$name_of_new_excel_without_ext = [System.IO.Path]::GetFileNameWithoutExtension($name_of_new_excel_xlsx)

$file_names_without_extra.Remove($name_of_new_excel_without_ext)


# Pipe data to Python
@($file_names_without_extra, $path_to_file_directory, $name_of_new_excel_xlsx) | python main.py
