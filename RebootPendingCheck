#### This script is intended to:
#### 1. Read a server list from an excel spreadsheet or populate it from SCCM.
#### 2. For each server, ping check.  If bad, color cell red.  If good, green.
#### 3. Query the registry to determine if SCCM is in a pending reboot state after
#### 	  applying updates.
#### 4. Apply True (pending a reboot) or False to column B with a timestamp.
####     If a server is red, you should check the console as it may have hung
####       on a scheduled reboot.
####     If Column B states "True" and the maintenance window has elapsed, it may 
####       have active user sessions and requires a reboot.

#-----------------Function Section Begins--------------------------------------------------
Function Get-SCCMCollectionMembers([string]$CollectionName)
{
$SiteServer = 'YourSCCMServer.YourDomain.local'
$SiteCode = 'YourSiteCode'
#Retrieve SCCM collection by name
$Collection = Get-WmiObject -ComputerName $siteServer -NameSpace "ROOT\SMS\site_$SiteCode"
     -Class SMS_Collection | where {$_.Name -eq "$CollectionName"}
#Retrieve members of collection
$SMSClients = Get-WmiObject -ComputerName $SiteServer -Namespace  "ROOT\SMS\site_$SiteCode"
     -Query "SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID='
     $($Collection.CollectionID)' order by name" | select Name
return $SMSClients
}
#-----------------Function Section Ends-----------------------------------------------------

#-----------------Excel Setup Begins--------------------------------------------------------
#### Spreadsheet Location
$strInputFile = 'C:\Temp\SERVERINFO\serverlist.xlsx'

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $True
$objSpread = $objExcel.Workbooks.Open($strInputFile)
$objSpread.Activate
$objWorksheet = $objExcel.Worksheets.Item(1)

# Set this variable to "Production" or to "DevQA" to populate a blank spreadsheet from SCCM.
#   This is normally used on the first run of a patching cyle.
# Set this variable to "None" to check status from a spreadsheet. This is generally used
#   for subsequent runs within a patching cycle.
$PopulateSheet = "None"

# Clear all first column highlighting
$objWorksheet.Cells.Item(1,1).EntireColumn.Interior.ColorIndex = 0
#-----------------Excel Setup Ends-----------------------------------------------------------


If ($PopulateSheet = "Production")
{
  # Get all members of the relevent collections and overwrite sheet values for Production.
  $intRow = 2

  $Targets1 = Get-SCCMCollectionMembers -CollectionName "ProductionServerPatching"
  $Targets2 = Get-SCCMCollectionMembers -CollectionName "ProductionPatchingAllUnidentifiedServers"
 
  foreach ($Target in $Targets1)
  {
  $Target = $Target -Replace '@{Name=',""
  $Target = $Target -Replace '}',""
  $objWorksheet.Cells.Item($intRow,1).value() = "$Target"
  $intRow ++
  }

  foreach ($Target in $Targets2)
  {
  $Target = $Target -Replace '@{Name=',""
  $Target = $Target -Replace '}',""
  $objWorksheet.Cells.Item($intRow,1).value() = "$Target"
  $intRow ++
  }
}
ElseIf ($PopulateSheet = "DevQA")
{
  # Get all members of the relevent collections and overwrite sheet values for Dev and QA.
  $intRow = 2

  $Targets1 = Get-SCCMCollectionMembers -CollectionName "non-prod Patching Collection"

  foreach ($Target in $Targets1)
  {
	$Target = $Target -Replace '@{Name=',""
    $Target = $Target -Replace '}',""
    $objWorksheet.Cells.Item($intRow,1).value() = "$Target"
    $intRow ++
  }
}
Else
{
# Skip sheet population
}

# Begin the primary loop to check status
$intRow = 2
Do{
#Reads in the hostname from the spreadsheet
$server = $objExcel.Cells.Item($intRow,1).Value()

#Gets the FQDN and writes domain to the spreadsheet
$PingResponse = ping -a $server
$SplitPingArray = $PingResponse.Split(" ")
$fqdn = $SplitPingArray[2]
$fqdn
Switch -wildcard ($fqdn)
	{
	*dmz.YourDomain1.local {$objWorksheet.Cells.Item($intRow,3).value() = "dmz.YourDomain1.local";break}
	*YourDomain2.local {$objWorksheet.Cells.Item($intRow,3).value() = "YourDomain2.local";break}
	*YourDomain1.local {$objWorksheet.Cells.Item($intRow,3).value() = "YourDomain1.local";break}
	default {$objWorksheet.Cells.Item($intRow,3).value() = "Not Resolveable";break}
	}

If (Test-Connection $server -Count 1 -Quiet){
	$pingStatus = "True"
	$objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 4
	$reboot = [wmiclass]"\\$server\root\ccm\ClientSDK:CCM_ClientUtilities"
	$result = $reboot.DetermineIfRebootPending() | select RebootPending
	$resultString = $result
	$resultString = $resultString -replace "@{RebootPending=", ""
	$resultString = $resultString -replace "}", ""
	Write-Host $resultString

 	$objWorksheet.Cells.Item($intRow,2).value() = "$resultString Date: $(Get-Date -format 'u')"
	}

else

    {
    $pingStatus = "false"
    $objWorksheet.Cells.Item($intRow,1).Interior.ColorIndex = 3
    }
    #$objWorksheet.Cells.Item($intRow,2).Value() = $pingStatus
   
$intRow ++
}
Until ($objWorksheet.Cells.Item($intRow,1).Value() -eq $null)
