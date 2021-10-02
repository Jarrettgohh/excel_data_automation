# excel_data_automation
- Repository to store the automations I have made for my internship
- Data analysis automation with Python and Powershell, to be used with Excel `.xls` files.
- The calculations includes k value; `dielectric constant` or `relative permittivity`, calculations from `capacitance` and `thickness` values of a `ferroelectric` sample wafer.


# Working principles
1. **Powershell** would be used to browse through a folder to get the names of all the `.xls` files  
2. To support working with Python `pandas`, **Powershell** would be used to convert the `.xls` file formats to `.xlsx`
3. **Powershell** would then *pipe* informations such as the array of name of the files and folder directory (information about sample wafer to be found from folder name) to **Python**
4. **Python** would be used to calculate various parameters with the data as mentioned above; `thickness` can be taken from the folder name
5. **Python** would *pipe* data back to **Powershell** with folder name for it to create a new file with `data_calulations` extended to its name, with extension of `.xlsx`
6. The calculated data would then be formatted apprioprately into the new `.xlsx` file created

-- UNCOMPLETED --
