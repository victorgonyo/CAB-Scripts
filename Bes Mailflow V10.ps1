#Bes 10,5,4 Mailflow Validation
$LogDir="C:\Avanade"
$ServerName=$env:COMPUTERNAME
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ServerName)
# Check for BES Installations
$BesV10 = "SOFTWARE\\Wow6432Node\\Research In Motion\\BlackBerry Enterprise Service\\Logging Info"
$BesV5 = "SOFTWARE\\Wow6432Node\\Research in Motion\\BlackBerry Enterprise Server\\Logging Info"
$BesV4 = "SOFTWARE\\Research in Motion\\BlackBerry Enterprise Server\\Logging Info"
$BesV10=$reg.OpenSubKey($BesV10)
$BesV5=$reg.OpenSubKey($BesV5)
$BesV4=$reg.OpenSubKey($BesV4)

If ((Test-Path $LogDir) -eq $false)
	{mkdir $LogDir | out-null}
$filename=("{0}\{1}_BES.txt" -f $LogDir,$env:ComputerName)
$line="--------------------------------------------------"
$BesVersion=$null
Function Get-BESLogFolder(){
	If ($BesV4)
		{
		Return $BesV4.GetValue("LogRoot")
		}
	If ($BesV5)
		{
		Return $BesV5.GetValue("LogRoot")
		}
	If ($BesV10)
		{
		Return $BesV10.GetValue("LogRoot")
		}
	}
Function Get-BESVersion(){
	If ($BesV4)
		{
		Return 4
		}
	If ($BesV5)
		{
		Return 5
		}
	If ($BesV10)
		{
		Return 10
		}
	}
function Log {
	param([string]$text)
	#Output to logfile
	if($LogToFile -eq $true) {Out-File $filename -append -noclobber -inputobject $text -encoding ASCII}
	#Output to screen
	Write-Output $text 
}
$BesVersion=Get-BESVersion
Write-Output "BES v$BesVersion detected."
$Baselogfolder=Get-BESLogFolder
$server=$env:computername  
$Folder=Join-Path -Path $Baselogfolder -ChildPath ("{0:yyyyMMdd}" -f (Get-Date))  
if ((Test-Path $Folder) -eq $false)
	{
	Log ("{0} - Could not find Log Folder" -f (Get-Date))
	Write-Output
	exit
	}
If ($BesVersion -eq 4 -or $BesVersion -eq 5)
	{
	$files=Get-ChildItem -Path $Folder -Filter "*MAGT*"
	$agents=@()
	foreach ($file in $files) {
		#Agent Number
		[string]$filename=$file.name
		$cnt=$filename.IndexOf("MAGT_")+5
		$agents+=$filename.Substring($cnt,$filename.Substring($cnt).IndexOf("_"))
	}
	#Remove Duplicate values
	$agents=$agents|sort -Unique  

	Log ("{0} - Found [{1}] Agent(s)" -f (Get-Date),$agents.count)

	foreach ($agent in $agents) 
		{
		$AgentFile=$files| Where {$_.Name -match ("{0}_MAGT_{2}_{1:yyyyMMdd}" -f $server, (Get-Date),$agent)}|Sort CreationTime -Descending |Select-Object -First 1
		Log ("{3}`n{0} - Checking Agent: {1}`nUsing file - {2}`n{3}" -f (get-date),$agent,$AgentFile.name,$line)
		#Check for user count - Return whole line
		$search="[30362]"
		Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile |%{$_.Line}|out-string
		Log ("{1}`n{0} - Checking for delivered messages`n{1}" -f (Get-Date),$line)
		#check for Delivered Messages - Last 5 Instances
		$search="[30097]"
		$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
		If ($TotalMessages)
			{
			Write-Output ("Found [{0}] delivered messages" -f $TotalMessages.count)|out-string
			$TotalMessages|Select-Object -Last 5|%{$_.Line}|out-string
			}else{
			Write-Output "***Found [0] delivered messages***"
			}
		#check for Failed to Resolve Name in ScanGAL - Last 10 Instances
		$search="[40210]"
		$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
		If ($TotalMessages)
			{
			Write-Output ("Found [{0}] Failed name resolution" -f $TotalMessages.count)|out-string
			$TotalMessages|Select-Object -Last 10|%{$_.Line}|out-string
			}else{
			Write-Output "Found [0] Failed name resolution"
			}
		$search="[20482]"
		$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
		If ($TotalMessages)
			{
			Write-Output ("***Found [{0}] 20482 events. Please failover BES to standby instance and restart BES services***" -f $TotalMessages.count)|out-string
			$search="[30447]"
			Write-Output ("These are the last 10 SMTP addresses that are erroring -")
			$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
			$TotalMessages | Select-Object -Last 10 | %{$_.Line} | Out-String
			}
		}
	}
If ($BesVersion -eq 10)
	{
	$files=Get-ChildItem -Path $Folder -Filter "*DISP*"
	$DispFile=$files | Sort CreationTime -Descending | Select-Object -First 1
	Log ("{2}`n{0} - Checking DISP Log using file - {1}`n{2}" -f (Get-Date),$DispFile.name,$line)
	#check for user count
	$search="[30558]"
	$Users=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $DispFile |%{$_.Line} |Select-Object -Last 1 | out-string
	If ($Users)
	{
	Log ($Users)
	}else{
	Log ("No Users connected log found in file")
	}
	#check for delivered packets
	$search="[30368]"
	$SentPackets=@(Select-String -AllMatches -SimpleMatch -Pattern $search -Path $DispFile | %{$_.Line}).count
	Log ("Found [{0}] delivered packets." -f $SentPackets)
	}
