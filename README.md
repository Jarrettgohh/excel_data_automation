# excel_data_automation
- Repository to store the automations I have made for my internship
- Data analysis automation with Python and Powershell, to be used with Excel `.xls` files.
- The calculations includes k value; `dielectric constant` or `relative permittivity`, calculations from `capacitance` and `thickness` values of a `ferroelectric` sample wafer.

# Usage
1. Store the script anywhere in the file directory
2. Execute the script passing in the path to the folder that contains the data as the argument Eg: `C:/Users/<user_name>/Desktop/CV_test`
3. Let the *magic* begin!

# Working principles
1. **Powershell** would be used to browse through a folder to get the names of all the `.xls` files  
2. To support working with Python `pandas`, **Powershell** would be used to convert all the `.xls` file formats to `.xlsx`
3. **Powershell** would then *pipe* informations (array of file names and file directory to the data)t o **Python**
4. **Python** would be used to calculate various parameters with the data (Parameters mentioned above); `thickness` can be taken from the folder name
5. **Python** would *pipe* data back to **Powershell** to open the newly created `.xlsx` file along with all the calculations
6. The calculated data would then be formatted apprioprately into the new `.xlsx` file created

-- UNCOMPLETED --

# Required Python packages
1. `pandas`
2. `openpyxl` (To work with `.xlsx` files)
3. `sys`; comes with Python
4. `json`; comes with Python
