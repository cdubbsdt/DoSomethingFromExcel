#### Spreadsheet Location (full path)
$strInputFile = "C:\Temp\SERVERINFO\serverinfo.xlsx"

#### Create the application object for Excel
$objExcel = New-Object -ComObject Excel.Application

#### Make the running copy of excel visible on the desktop while running.
$objExcel.Visible = $True

#### Set the spreadsheet to the input file,
#### activate it, and open it to the first page.
$objSpread = $objExcel.Workbooks.Open($strInputFile)
$objSpread.Activate
$objWorksheet = $objExcel.Worksheets.Item(1)

#### Set the starting row number to 2, accounting for column headers.
$intRow = 2

#### Main script loop - continues until the first blank row in Column 1.
Do{
  #### Change the working cell to Yellow in case your script hangs.
  #### It will be overwritten in this loop, but remember to clear it
  ####   at the end if you aren't highlighting cells in your main loop.
    $objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 6


  #### Read current loop row and column 1 into a variable.
  #### If the file has more than 1 worksheet (tab/page) and
  ####  it is saved on a different worksheet, this command
  ####  will read data from the wrong sheet. Use objWorksheet!!!
  $server = $objExcel.Cells.Item($intRow,1).Value()

  #### Perform the work
  If (Test-Connection $server -Count 1 -Quiet){
    $pingStatus = "True"

    #### Write status back to cell by coloring the cell green.
    $objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 4
    }
  else
    {
    $pingStatus = "false"

    #### Write status back to cell by coloring the cell red
    $objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 3
    }

  #### Write ouptut of the IF statement to column 2 to support pivot tables.
  $objWorksheet.Cells.Item($intRow,2).Value() = $pingStatus

  #### Clear the cell color
  $objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 0


#### Increment the loop   
$intRow ++
}
#### Script ends at first blank entry in column 1
Until ($objWorksheet.Cells.Item($intRow,1).Value() -eq $null)

#### Optionally, you could save and close the workbook and exit excel here.
#### This is usually ad-hoc work, so I leave excel running.

#### The Quit function will hang if working from a csv you added colors.
#$objExcel.Save()
#$objExcel.Quit()

#### Release references in case other scrips are run in the same session
#### where the new script doesn't set the sheet.
# $a = Release-Ref($objWorksheet)
# $a = Release-Ref($objWorkbook)
# $a = Release-Ref($objExcel)
