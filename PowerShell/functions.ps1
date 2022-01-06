Function connvert_xls_to_xlsx($path_to_file_directory){

# Get all the files present in the folder
$files = Get-ChildItem -Path $path_to_file_directory

foreach ($file in $files) {
 # Rename all .xls files to .xlsx extension
 $path_to_file = -join ($path_to_file_directory, -join ("\", $file.Name))

 $file_name_without_extension = [System.IO.Path]::GetFileNameWithoutExtension($file)
 $file_name_xlsx = -join ( $file_name_without_extension, '.xlsx')

 Rename-Item -Path $path_to_file -NewName  $file_name_xlsx
}
}