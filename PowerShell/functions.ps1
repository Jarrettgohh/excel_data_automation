Function convert_xls_to_xlsx($file_dir,$file_name) {

  $path_to_file = -join ($file_dir, -join ("\", $file_name))

  $file_name_without_extension = [System.IO.Path]::GetFileNameWithoutExtension($file_name)
  $file_name_xlsx = -join ( $file_name_without_extension, '.xlsx')

  Rename-Item -Path $path_to_file -NewName  $file_name_xlsx
  }