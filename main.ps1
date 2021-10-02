
# Path: to get from user input

$path_to_file_directory = #'C:\Users\gohja\Desktop\CV test\PLD_30%_20nm\'
'C:\Users\gohja\Desktop\CV test\PPP_69%_20nm'

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

# Get name of folder
Set-Location $path_to_file_directory

$current_folder = Split-Path -Path (Get-Location )-Leaf

Set-Location $current_path

# Pipe array consisting of file names and directory path to Python
@($file_names , $current_folder ) | python main.py
