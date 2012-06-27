<#
.SYNOPSIS
Ping workstations and report results in a csv.

.PARAMETER
-targets The file to be used which lists hosts to ping.
File should be in text format with single FQDN server name on each line
-outputDest Specify a destination folder where the report will be saved
Do not include a trailing slash

.DESCRIPTION
Loops through array of computers and tests network connectivity via WMI ping. 
Results are created in a new csv file but can easily be changed to anything else (host window, out-gridview, text file etc.)

.EXAMPLE
.\get-ping-status-noexcel.ps1 -targets "D:\CN\target_servers.txt" -outputDest "D:\CN\Output"
.\get-ping-status-noexcel.ps1 "D:\JB\target_servers.txt" "D:\JB\Output"

This is essentially the same as the get-ping-status-excel script but without the requirement to have excel installed and the ability to be run non interactively
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
	[string]$TargetFile,
	[Parameter(Mandatory=$true)]
	[string]$OutputFolder
)


# Setup variables for input and output arrays
$Filename = "$OutputFolder\Ping_Results{0:yyyddMM-HHmmss}.csv" -f (get-date)
$InputFile = get-content $TargetFile
$Output = @()
#$colComputers will be our list of computers to ping from the input file.
$colComputers = $InputFile

# Loop through each computer in colComputers to perform the ping
Foreach ($strComputer in $colComputers) 
{
# Blank out arrays
	$Status = @()
	$Result = @()
	$IPAddress = @()

# Perform wmi ping using gwmi cmdlet with win32_pingstatus class
	$ping = get-wmiobject win32_pingstatus -filter "address='$strcomputer'" | select-object Statuscode,protocoladdress,PrimaryAddressResolutionStatus

# If the statuscode is 0 ping was successful and we can report on that along with the IP address
	if ($ping.statuscode -eq 0)
	{
		$Status = "Online"
		$Result = "Request Successful"
		$IPAddress = $Ping.ProtocolAddress
	
# What if the ping failed?		
	} else 
	{
# Set our status to Offline and our IP address to blank
# the WMI ping only returns an IP if the ping is successfull
		$Status = "Offline"
		$IPAddress = ""

# The switch table maps status codes to friendly text
# There are a lot more status codes, but most of them will never be seen
# http://msdn.microsoft.com/en-us/library/windows/desktop/aa394350(v=vs.85).aspx
		$result = switch ($Ping.statuscode)		
		{
			11010 {"Request Timed Out"}
			11013 {"TTL Expired in transit"}
			11003 {"Destination Host Unreachable"}
			default {"Unknown code"}
		}
	}

# An outlying ping failure is if name resolution fails
# In this case the PrimaryAddressResolutionStatus property will be set to anything but 0
# So we check for that, and if it happened, report that DNS failed
	if ($ping.PrimaryAddressResolutionStatus -ne 0)
	{
		$Status = "Offline"
		$Result = "DNS Lookup Failed"
		$IPAddress = ""
	}

# Write our results to $Output
# One day I'll learn to do this properly...
	$Results = {} | select HostName,HostStatus,PingResult,IP
	$Results.Hostname = $strComputer
	$Results.HostStatus = $Status
	$Results.PingResult = $Result
	$Results.IP = $IPAddress
	$Output += $Results
}

# Export our results to a CSV file using the $filename created at the start of the script
# At this point if you don't want to use a CSV you can change $output to do something else 
# E.g. write to console, out-gridview, out-file etc
# If you do this, you may want to ditch the output parameter at the start of the script
$Output | export-Csv $filename -NoTypeInformation