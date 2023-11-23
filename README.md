# nessus-input-parser
Simple parser that trasnforms nessus compliance rules input to an **excel** file.

## Dependencies
Some packages need to be installed, so you should run this commands:
- pip install pandas
- pip install df
- pip install openpyxl
- pip install fsspec

## Setup
To make it work, the input must be in **xml** format (there may be some format errors such as "&", "&&", or "<" and ">" characters in the input file, so these should be manually removed) and the xml tree must be **<custom_item>** tags inside a **\<word\>** tag.

## Usage
python nessuspar.py -h ------> Shows help message for usage information.

python nessuspar.py -f <input_file> -------> Takes <input_file> (xml format) and creates an "output.xlsx" file in the current folder.

python nessuspar.py -f <input_file> -o <output_file> --------> Takes <input_file> (xml format) and creates an output file in <output_file> (xlsx format).


## Issues
The results of the excel may contain unnecesary columns and some columns may not contain any data. To solve this last issue, edit the 38th line to add the name of the column with missing values.

