
# Path: to get from user input
$path_to_file_directory = #'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\'
'C:\Users\gohja\Desktop\CV test\PPP_69%_20nm'

# Get all the files present in the folder
$files = Get-ChildItem -Path $path_to_file_directory


# Initialize empty array to store file names
$file_names = @()

foreach ($file in $files) {
 # # Rename all .xls files to .xlsx extension
 # $path_to_file = -join ($path_to_file_directory, -join ("\", $file.Name))

 # $file_name_without_extension = [System.IO.Path]::GetFileNameWithoutExtension($file)

 # $file_name_xlsx = -join ( $file_name_without_extension, '.xlsx')


 # Rename-Item -Path $path_to_file -NewName  $file_name_xlsx


 # Add file names to array
 $file_name = $file.Name
 $file_names = $file_names + $file_name 
}


$current_path = Get-Location

# Get name of folder containing data
Set-Location $path_to_file_directory

$current_folder = Split-Path -Path (Get-Location )-Leaf


# Creating a new .xlsx file
$name_of_new_excel_xlsx = -join ($current_folder, '_data_calculations.xlsx')

$full_path_to_new_excel = -join ($path_to_file_directory, -join ('\', $name_of_new_excel_xlsx ))

$new_excel = New-Object -ComObject excel.application
$new_excel.visible = $False

$new_excel_workbook = $new_excel.Workbooks.Add()
$new_excel_workbook.SaveAs($full_path_to_new_excel) 
$new_excel.Quit()


# Set location back to the current script path
Set-Location $current_path

# Pipe array consisting of file names and directory path to Python
@($file_names , $path_to_file_directory) | python main.py
