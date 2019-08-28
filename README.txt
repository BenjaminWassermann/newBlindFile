newBlindFile.py
Ben Wassermann
8/28/2019

The purpose of this script is to generate a randomized blind sheet for
mouse studies which utilize 0 or more surgeries and 0 or more dose types.

Dependencies:

	Requires openpyxl, pathlib, pandas, and xlwings

	openpyxl, pandas, and xlwings are each used to interact with .xlsx files

	openpyxl: 
		general excel file management and placement of data in cells

	pandas:
		allows for sorting of a dataframe and saving to an excel sheet

	xlwings:
		openpyxl does not allow for running formulas. xlwings allows
		me to side-step this by openning the sheet after modification
		and calculating the formulas, saving to the same sheet.

How-To:

	1. Install python3 and perform pip install openpyxl, pandas, pathlib,
	   xlwings, and pypiwin32.
	2. Run script 'newBlindFile.py'
	3. Provide inputs to the script:
		a. total number of animals
		b. zero or more surgery types and total desired for each
		c. zero or more dose types
		d. project code
	4. Script will save randomized blind-file with all required groups
	   and group counts to [project code].xlsx

