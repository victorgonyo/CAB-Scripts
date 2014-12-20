
$ServerName=$env:ComputerName
$line="--------------------------------------------------"
$Stars="******************************************"

function Log 
{
	param([string]$text)
	#Output to logfile
	#Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
	#Output to screen
	Write-Output $text 
}
$starttime=(get-date)

$computer = "LocalHost" 
$namespace = "root\cimv2" 

#region	Connect to WSUS

#create session object
$ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
$Session = New-Object -ComObject "Microsoft.Update.Session" 
$Searcher = $Session.CreateUpdateSearcher()

$Service=$ServiceManager.Services|Where {$_.IsDefaultAUService -eq $true}

If($Service)
{
	$serviceName = $Service.Name
}
Else
{
	Log ("{1} - {0} - Can't Find Update Service" -f (Get-Date),$ServerName)
	Return
}
Log ($line)
Log ("{2} - {0} - Connecting to {1}, please wait..." -f (Get-Date),$serviceName,$ServerName)
Log ($line)
#endregion

#region Check for updates
Try
{
	$Results = $Searcher.Search("IsInstalled=0")
}
Catch
{
	If($_ -match "HRESULT: 0x80072EE2")
	{
		Log ($line)
		Log ("{1} - 0} - WARNING: Unable to connect to Windows Update server" -f (Get-Date),$ServerName)
		Log ($line)
	}
	Return
}
Log ($line)
Log ("{2} - {0} - Found {1} updates for installation" -f (Get-Date),$Results.Updates.count,$ServerName)
Log ($line)
#endregion