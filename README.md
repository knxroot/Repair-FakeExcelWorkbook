# Automagically fix the Excel warning: The file you are trying to open is in a different format #

Open a fake Excel file with MS Excel in background (silence mode) and save again as XLSX format (Real Excel File).

## What is this? ##
In some case you can't open a Microsoft Office Excel file, excel show the message: "The file you are trying to open, 'example.XLS' is in a different format than specified by the file extension. Verify that the file is not corrupted and is from a trusted source before opening the file. Do you want to open the file now?".

While working with automatic process, developers need repair the file in execution time for read the Excel file from the next step (Pentaho Data Integration Process, Bash, etc). Ofcourse if you check 'yes' you can open the file, but in a automatic process you can't are click in 'yes' for force the opened.

With this script you can open the bad file and save again in XLSX format using Microsoft Excel in background (silence mode), ideal for use in automatic process !.

## Requeriments ##

- MS Windows 7 + MS Office 2007 or later

## Use ##

`powershell.exe -ExecutionPolicy Bypass -file "c:\Repair-FakeExcelWorkbook\Repair-FakeExcelWorkbook.ps1" "PATH_TO_FAKE_EXCEL_FILE" "PATH_TO_FIXED_EXCEL_FILE"`

## How to do it... ##

Carry out the following steps:

1. Download, unzip the software in `C:\Repair-FakeExcelWorkbook` folder
2. Try open  `C:\Repair-FakeExcelWorkbook\examples\example1.XLS`, can you see the warning?
3. Open cmd and execute a example: `powershell.exe -ExecutionPolicy Bypass -file "c:\Repair-FakeExcelWorkbook\Repair-FakeExcelWorkbook.ps1" "c:\Repair-FakeExcelWorkbook\examples\example1.XLS" "c:\Repair-FakeExcelWorkbook\examples\example1-fixed.xlsx"`
4. Try Open  `C:\Repair-FakeExcelWorkbook\examples\example1-fixed.xlsx`, you should not see the warning, this is a real Excel File :).

## How it works... ##

The file `C:\Repair-FakeExcelWorkbook\examples\example1.XLS` is a plain text file with the extension .XLS, is a fake Excel file. Excel detect this before open. The script use PowerShell and [the Excel ComObject call](http://msdn.microsoft.com/en-us/library/wss56bz7.aspx "COM call") for open the fake excel and save again in a real xlsx format.

## See also ##
- [https://gist.github.com/gabceb/954418](https://gist.github.com/gabceb/954418 "Powershell script to convert all xls documents to xlsx in a folder recursively")
- [http://msdn.microsoft.com/en-us/library/wss56bz7.aspx](http://msdn.microsoft.com/en-us/library/wss56bz7.aspx "Excel Object Model Overview")