# Instructions for IT

## Installation

This automation process needs Python3 installed on the host Mac. 

The easiest way to do this is to use homebrew. From the commandline (logged in as admin), follow the instructions at [homebrew (brew.sh)](https://brew.sh).
Once this completes, still logged in as admin

`brew install python3`

David will also need the OneDrive app, and it will need to sync the required folders to the local machine

Finally, run `pip3 install openpyxl`

# Instructions for David

This version of the script only works for RP4 spreadsheet templates.
Specifically this requires the number of examiners to be in cell C2;
the main assessment grid (grey area) to run from C13:Q23, and;
the comments to be A27. 

Create a folder for every student, named whatever you want. 

For example, it might be `Documents/OneDrive/Assessments/2021/S1/RP4/StudentName`

In each of those folders, the resulting spreadsheets need to be included in a subfolder called `data`, which means that you'll have `RP4/StudentName/data/spreadsheet1.xlsx`, `RP4/StudentName/data/spreadsheet2.xlsx`, `RP4/StudentName/data/spreadsheet3.xlsx`.
The excel files can be called whatever you want. The current assessment folder structure should work just fine.

Download `excel-wrangle-recursive.py` into the RP4 folder. In that folder you should therefore have the .py file, and then as many folders called StudentName as necessary, within each of which will be a data folder with three excel spreadsheets.

Right click on the RP4 folder, and click 'services', then 'open terminal at folder'. This will open the command line.

Type `python3 excel-wrangle-recursive.py` and press enter, then watch the magic.

