# DoSomethingFromExcel
This is for automating small dumb tasks with PowerShell when the source data is in CSV or XLSX format.

The formatting and coding practices in this template are not good. Please learn scripting from someone with the time to do things better.

Essentially, "DoSomethingFromExcel" is a template script designed to open a CSV or XLSX file and do some piece of work leveraging input from each row in the first column of a spreadsheet. It might be user names, computer names, IPs, or any other list of data you may be given to perform a task.

Example 1 (DoSomethingFromExcel.ps1) does the following things:
  Declares a variable to store the path to the input file.
  Created the Excel application object with the option to keep it visible.
  Opens the spreadhseet and sets the starting row to 2 (to accomodate column headers).
  Begins a Do loop that continues until it hits a row with a null value in column 1.
  This exmaple uses server names or IPs to run a Test-Connection PowerShell cmdlet with only 1 ping, assuming a reliable network.
  If the connection test is successful, it changes the cell color to Green, else set the cell color to Red.
  It writes the result of the connection test into column 2 of the working row.


This concept can be expanded by adding mini PowerShell scripts as functions. This is intended as an entry-level demonstration on how to interact with Excel as an application object in PowerShell. There are many more functions than reading a cell value into a variable, changing the color, and writing values back to Excel.
