<#
.SYNOPSIS
Ping workstations and report results in Excel

.PARAMETER
No parameters required

.DESCRIPTION
Loops through array of computers and tests network connectivity via WMI ping. Results are created in a new Excel worksheet in real time.
#>

$erroractionpreference = "SilentlyContinue"

# Create a new Excel Workbook and make it visible
$ExcelObject = New-Object -comobject Excel.Application
$ExcelObject.visible = $True
$ExcelWorkbook = $ExcelObject.Workbooks.Add()
$ExcelWorksheet = $ExcelWorkbook.Worksheets.Item(1)
# This will be the name of the Excel sheet. Goes down to seconds to stop multiple instances of script from replacing old reports
$Filename = "d:\cn\PingList\PingList{0:ddMMyyy-HHmmss}.xls" -f (get-date)

$InputFile = get-content D:\CN\Scripts\servers.txt

# Create column headings and format them nicely
$ExcelWorksheet.Cells.Item(1,1) = "Machine Name"
$ExcelWorksheet.Cells.Item(1,2) = "Ping Status"
$ExcelWorksheet.Cells.Item(1,3) = "Status Code"
$ExcelWorksheet.cells.item(1,4) = "IP"
$ExcelHeadings = $ExcelWorksheet.UsedRange
$ExcelHeadings.Interior.ColorIndex = 15
$ExcelHeadings.Font.ColorIndex = 11
$ExcelHeadings.Font.Bold = $True

# $ExcelHeadings.EntireColumn.AutoFit()

# This will set our Excel cursor to the second row in the worksheet so we don't overwrite headings
$intRow = 2

# Get list of computers to ping from source and start working
$colComputers = $InputFile

# Loop through source array 

foreach ($strComputer in $colComputers)
{

	# Write hostname from text file into first cell in column
	$ExcelWorksheet.Cells.Item($intRow, 1) = $strComputer.ToUpper()

	# Use WMI Ping on hostname and select the properties we are interested in
	$ping = get-wmiobject win32_pingstatus -filter "address='$strcomputer'" | select-object Statuscode,protocoladdress,PrimaryAddressResolutionStatus

	# If this attribute doesn't equal 1, DNS lookup has failed
	if ($ping.PrimaryAddressResolutionStatus -ne 0)
	{
		$ExcelWorksheet.Cells.Item($intRow, 2) = "Offline"
		$ExcelWorksheet.cells.item($introw, 3) = 'DNS Lookup Failed'
		$ExcelWorksheet.cells.item($introw, 3).interior.ColorIndex = 3

	}

	# If the statuscode is 0, ping has succeeded
	if ($ping.statuscode –eq 0)
	{
		$ExcelWorksheet.Cells.Item($intRow, 2) = "Online"
		$ExcelWorksheet.cells.item($intRow, 3) = "Request Successful"
		$ExcelWorksheet.cells.item($introw, 4) = $Ping.ProtocolAddress
		$ExcelWorksheet.cells.item($introw, 3).interior.ColorIndex = 4

		# Otherwise the ping has failed but why?
	}
	else
	{
		
		$ExcelWorksheet.Cells.Item($intRow, 2) = "Offline"

		# This code means it has timed out
		if ($ping.statuscode -eq 11010)
		{
			$ExcelWorksheet.cells.item($introw, 3) = 'Request Timed Out'
			$ExcelWorksheet.cells.item($introw, 3).interior.ColorIndex = 6
		}
		# And this one means TTL has expired
		if ($ping.statuscode -eq 11013)
		{
			$ExcelWorksheet.cells.item($introw, 3) = 'TTL Expired in Transit'
			$ExcelWorksheet.cells.item($introw, 3).interior.ColorIndex = 7
		}

		# Last step is to write the IP address into the fourth column. This will only be retrieved if the ping has succeeded
		$ExcelWorksheet.cells.item($introw, 4) = $Ping.ProtocolAddress
		# $ExcelHeadings.EntireColumn.AutoFit()
	}
	#$Reply = ""

	#Move to the next row in worksheet
	$intRow = $intRow + 1
}
#Auto fit columns and save workbook using $filename
$ExcelHeadings.EntireColumn.AutoFit()
$ExcelWorkbook.SaveAs("$FileName")