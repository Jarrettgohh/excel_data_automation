# excel_data_automation
- Repository to store the automations I have made for my internship
- Data analysis automation with Python and Powershell, to be used with Excel `.xls` files.
- The calculations includes k value; `dielectric constant` or `relative permittivity`, calculations from `capacitance` and `thickness` values of a `ferroelectric` sample wafer.

# Usage
1. Store the script anywhere in the file directory
2. Execute the script passing in the path to the folder that contains the data as the argument Eg: `C:/Users/<user_name>/Desktop/CV_test`
3. Let the *magic* begin!

# Working principles

`PowerShell`
1. **Powershell** would be used to browse through a folder to get the names of all the `.xls` files  
2. To support working with Python `pandas`, **Powershell** would be used to convert all the `.xls` file formats to `.xlsx`
3. **Powershell** would then *pipe* informations (array of file names and file directory to the data) to **Python**

`Python`
1. Use `regex` to split the data between the different sizes 
2. **Python** would be used to calculate various parameters with the data (Parameters mentioned above); `thickness` can be taken from the folder name
3. **Python** would *pipe* data back to **Powershell** to open the newly created `.xlsx` file along with all the calculations
4. The calculated data would then be formatted apprioprately into the new `.xlsx` file created

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
