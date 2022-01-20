# excel_data_automation

# Version 2
- Support for matching files to read according to a certain regex/string pattern

# Options
1. *`Option 1`*
- This option can be used to extract data from any supported file formats (listed below under `supported file formats`) file to excel (`.xlsx`) file.
- Select this option if you wish to update the config.json file
- For example, the row and column values of the data to extract may be unclear when extracting from a text (`.txt`) file. Thus, this option could be used to temporarily extract and transfer the text file to an excel file to be able to decide the row and column values to read
- The configurations can be set in the `config.json` file

2. *`Option 2`*
- This option can be used to extract data from any supported file formats (listed below under `supported file formats`), and transfer directly into a `.xlsx` file. 
- The configurations can be set in the `config.json` file

# Configuration file
- The configuration for this automation could be configured in the `config.json`
- The following are the configuration options

1. 


# Supported file formats
1. Text (`.txt`) file
2. Comma-separated values (`.csv`) file
3. Microsoft Excel Binary File format (`.xls`) file
-> To allow support of older Excel versions


# Required Python packages; outside of the basic ones
1. `pandas`
2. `openpyxl` (To work with `.xlsx` files)

