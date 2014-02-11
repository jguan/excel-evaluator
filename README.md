JExcelEvaluator
===============

This java command-line program to to the following:

1. load a spreadsheet
2. update values in the spreadsheet (defined by *input.txt*)
3. evaluate the spreadsheet calculations
4. print out the values of specified cells in the spreadsheet (defined by *output.txt*)

The program takes three command line parameters:

-model = fully specified path to an XLSX file (the model)  
-input = fully specified path to an input file (*input.txt*)  
-output = fully specified path to an output file (*output.txt*)  

*input.txt* is a space-delimited text file. Each line contains:

column0: spreadsheet cell coordinate (eg D10)  
column1: a value to put in that cell (eg 0.15)  

Example:  
D10 5  
D11 0.15  
D12 1200.00  

*output.txt* contains a single column of cell coordinates. 

The java program will read the values from the cells and print the output as follows:

<cell coordinate> <space> <value>

Example:  
D37 1992.606  
D38 0.1  
D39 5  
