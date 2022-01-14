Function convert_xls_to_xlsx($file_dir,$file_name) {

 Write-Output $file_dir
 Write-Output $file_name



  # # Rename all .xls files to .xlsx extension
  # $path_to_file = -join ($file_dir, -join ("\", $file_name))

  # $file_name_without_extension = [System.IO.Path]::GetFileNameWithoutExtension($file_name)
  # $file_name_xlsx = -join ( $file_name_without_extension, '.xlsx')

  # Rename-Item -Path $path_to_file -NewName  $file_name_xlsx
  }