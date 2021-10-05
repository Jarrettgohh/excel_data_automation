# excel_data_automation
- Repository to store the automations I have made for my internship
- Data analysis automation with Python and Powershell, to be used with Excel `.xls` files.
- The calculations includes k value; `dielectric constant` or `relative permittivity`, calculations from `capacitance` and `thickness` values of a `ferroelectric` sample wafer.

# Usage
1. Store the script anywhere in the file directory
2. Execute the script passing in the path to the folder that contains the data as the argument Eg: `C:/Users/<user_name>/Desktop/CV_test`
3. Let the *magic* begin!

# Working principles

**PowerShell**
1. Browse through a folder to get the names of all the `.xls` files  
2. To support working with Python `pandas`, convert all the `.xls` file formats to `.xlsx`
3. Create a new `.xlsx` with name of the folder name and "data_calculations" prefix (Eg: folder name: "PVD_20%_40nm" -> "PVD_20%_40nm_data_calculations")
4. Save the newly created `.xlsx` file as `.xlsm` to support macros and insert the VBA script for each individual calculations ("average_capacitance", "average_k", etc.) into the `.xlsm` file
5. *Pipe* the information containing array of file names, path to target folder directory and the name of newly created Excel file to **Python**

**Python**
1. Use `pandas` to copy the data from the individual files and format it in the new `.xlsm` file  
2. Generate the column index range to run the `macro` on. The format can be found from `config.json` (Eg.[{"cell_select": "B13", "cell_range": "B13:F13"]])
3. Run through a method to convert the column index range in Python to be a PowerShell object type 
4. *Pipe* information containing path to target file directory and the `list` of column index range (`cell_select` & `cell_range` with the proper format to pass as args into Excel macro VBA script) back to PowerShell (To call using `subprocess` to execute powershell  with path (Eg. "../macro.ps1"))

**PowerShell**
1. Run the macro according in the target file directory with the column index range information received as params

# Ideas
1. Have an `.xlsm` file that have pre-recorded macros (Allows taking in parameter to see number of device sizes to run for)
2. Use **PowerShell** to run the macro and transfer the output to the newly created `.xlsx` file (Perhaps **Python** would be better for the output transfer part) 

# config.json
1. The configuration of which rows and columns to read from the individual Excel sheets can be changed in the `config.json` file

-- UNCOMPLETED --

# References
1. https://stackoverflow.com/questions/38074678/append-existing-excel-sheet-with-new-dataframe-using-python-pandas/38075046#38075046

# Required Python packages
1. `pandas`
2. `openpyxl` (To work with `.xlsx` files)
3. `sys`; comes with Python
4. `json`; comes with Python
