Param([switch]$LogToFile=$false)
$LogDir="C:\Avanade"
if ((Test-Path $LogDir ) -eq $false) {mkdir $LogDir|out-null}
$filename=("{0}\{1}_BES.txt" -f $LogDir,$env:ComputerName)
$line="--------------------------------------------------"

function Get-BESLogFolder(){
	# Check for BES Installations
	$v5 = "HKLM:\SOFTWARE\Wow6432Node\Research in Motion\BlackBerry Enterprise Server\Logging Info"
	$v4 = "HKLM:\SOFTWARE\Research in Motion\BlackBerry Enterprise Server\Logging Info"
	#BES 4
	if (Test-Path $v4){	Return (Get-ItemProperty -path $v4).LogRoot	}
	#BES 5
	if (Test-Path $v5){ 	Return (Get-ItemProperty -path $v5).LogRoot }
	
	#For Testing
	#return "C:\Test"
} 

function Log {
	param([string]$text)
	#Output to logfile
	if($LogToFile -eq $true) {Out-File $filename -append -noclobber -inputobject $text -encoding ASCII}
	#Output to screen
	Write-Output $text 
}

$Baselogfolder=Get-BESLogFolder  
$server=$env:computername  
$Folder=Join-Path -Path $Baselogfolder -ChildPath ("{0:yyyyMMdd}" -f (Get-Date))  
if ((Test-Path $Folder) -eq $false) {Log ("{0} - Could not find Log Folder" -f (Get-Date));Write-Output ;return}
#Get All MAGT Logs
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

foreach ($agent in $agents) {
	$AgentFile=$files| Where {$_.Name -match ("{0}_MAGT_{2}_{1:yyyyMMdd}" -f $server, (Get-Date),$agent)}|Sort CreationTime -Descending |Select-Object -First 1
	Log ("{3}`n{0} - Checking Agent: {1}`nUsing file - {2}`n{3}" -f (get-date),$agent,$AgentFile.name,$line)
	#Check for user count - Return whole line
	$search="[30362]"
	Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile |%{$_.Line}|out-string
	Log ("{1}`n{0} - Checking for delivered messages`n{1}" -f (Get-Date),$line)
	#check for Delivered Messages - Last 5 Instances
	$search="[30097]"
	$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
	Write-Output ("Found [{0}] delivered messages" -f $TotalMessages.count)|out-string
	$TotalMessages|Select-Object -Last 5|%{$_.Line}|out-string
	#check for Failed to Resolve Name in ScanGAL - Last 10 Instances
	$search="[40210]"
	$TotalMessages=Select-String -AllMatches -SimpleMatch -Pattern $search -Path $AgentFile
	Write-Output ("Found [{0}] Failed name resolution" -f $TotalMessages.count)|out-string
	$TotalMessages|Select-Object -Last 10|%{$_.Line}|out-string

}