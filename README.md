# excel2latex
Simple script which creates a latex table from an Excel table while keeping the background color.

## How to use
### Install environment
You will need python, pandas and openpyxl.
```
conda env create -f environment.yml 
```
### Generate LaTeX script
Replace TABLE and SHEET with your filename and sheet name.
```
python excel2latex.py TABLE.xlsx SHEET
```
This command will generate 2 files "color_definitions.txt" and "latex_table.txt". 

### Include in .tex
You need the package:
```
\usepackage{colortbl}
```
After that include the color definitions and the table. 

### Notes
- If no background color is set in excel, openpyxl will translate that to black.
