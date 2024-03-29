# CODING         : UTF-8
# NAME           : Repair-FakeExcelWorkbook
# COMMENT        : Open a fake Excel file with MS Excel in background (silence mode) and save again as XLSX format (Real Excel File).
#
# AUTHOR         : Gustavo Lacoste (gustavo@lacosox.org)
# SUPPORTED      : W7 Professional.

<#
.SYNOPSIS
	Open a fake Excel file with MS Excel in background (silence mode) and save again as XLSX format (Real Excel File).
.EXAMPLE
	powershell.exe -ExecutionPolicy Bypass -file "c:\Repair-FakeExcelWorkbook\Repair-FakeExcelWorkbook.ps1" "c:\Repair-FakeExcelWorkbook\examples\example1.XLS" "c:\Repair-FakeExcelWorkbook\examples\example1-fixed.xlsx"
#>

param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$in_fpath,

		[Parameter(Position=1, Mandatory=$false)]
		[System.String]
		$out_fpath
	)

# Take a input file open this with Excel and save again in xlsx format
function Repair-FakeExcelWorkbook {
	param(
		[string]$fi,
		[string]$fo
	)
  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
  $objExcel = New-Object -ComObject excel.application
  $objExcel.visible = $false

  $cuInfo = [System.Globalization.CultureInfo]'en-US'
  $WorkBook=$objExcel.Workbooks.PSBase.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $objExcel.Workbooks, $fi, $cuInfo)
  if ($?) {
  	[void]$WorkBook.PSBase.GetType().InvokeMember('SaveAs', [Reflection.BindingFlags]::InvokeMethod, $null, $WorkBook, ($fo, $xlFixedFormat,$null,$null,$false,$false,"xlNoChange","xlLocalSessionChanges"), $cuInfo)
  }
  [void]$WorkBook.PSBase.GetType().InvokeMember('Close', [Reflection.BindingFlags]::InvokeMethod, $null, $WorkBook, 0, $cuInfo)

  $objExcel.Quit()
  $objExcel = $null
  [gc]::collect()
  [gc]::WaitForPendingFinalizers()
}

trap
{

	if (Test-Path $in_fpath -include *.xls, *.xlsx) {
	
		if ([string]::IsNullOrEmpty($out_fpath)) {
			$out_fpath = (Join-Path (Get-ChildItem $in_fpath ).DirectoryName (Get-ChildItem $in_fpath ).Basename) + "-fixed.xlsx"
		}
		
		if ($out_fpath.ToLower().EndsWith("xlsx")) {
			if (Test-Path $out_fpath) {
				Remove-Item -Path $out_fpath -Force
			}
			
			Repair-FakeExcelWorkbook $in_fpath $out_fpath
			
			if (Test-Path $out_fpath) {
				exit 0
			}
			
		} else {
			$path_error = New-Object System.FormatException "The output file path should by .xlsx file"
		}
		
	} else {
		$path_error = New-Object System.FormatException "The input file path is not a valid file"
	}
	
	Throw $path_error 
} throw 'Unknown Error, sorry :('