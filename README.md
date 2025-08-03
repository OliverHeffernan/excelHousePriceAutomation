# About:
This is a python script that takes an excel spreadsheet with house addresses, finds the current valuation of that house according to homes.co.nz, and then stores that in the spreadsheet.

The first time running the script will be slow, as it needs to find the home.co.nz links for each address.
These links will be stored in the excel file, so running the script in the future will be much faster.

# Usage:
## Setting up the spreadsheet
You should have addresses in column C. Leave column D blank; the URLs will be stored there.
Leave column E blank; the prices will be stored there.
Run the main.py file using the following command.
## Running the script
```shell
python3 main.py {path to spreadsheet}
```
Replacing the {path to spreadsheet} with the path to the spreadsheet with the addresses, e.g. test.xlsx
