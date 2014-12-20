Param([string]$database)
get-pssnapin -registered | add-pssnapin -ErrorAction SilentlyContinue
$stars="***************************"
$date=get-date -format g
Write-Output $stars,$date,$stars
$Server = $env:COMPUTERNAME
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Server)
$Exch2007="SOFTWARE\\Microsoft\\Exchange\\V8.0"
$Exch2010="SOFTWARE\\Microsoft\\ExchangeServer\\v14"
$Exch2013="SOFTWARE\\Microsoft\\ExchangeServer\\v15"
$reg2007=$reg.OpenSubKey($Exch2007)
$reg2010=$reg.OpenSubKey($Exch2010)
$reg2013=$reg.OpenSubKey($Exch2013)
If ($reg2007)
	{
	Write-Output $stars,"Exchange 2007 detected.",$stars,""
	$Clusters = Get-MailboxServer | ? {$_.ClusteredStorageType -ne "Disabled"}
	ForEach ($Cluster in $Clusters) 
		{
		[String]$NodeName=$Cluster.Name
		$ClusterStatus = Get-ClusteredMailboxServerStatus -Identity $Cluster.Name | Select -Expand OperationalMachines | ForEach {If($_ -like "*Active*") {$_}}
		[String]$ActiveNode=$ClusterStatus.Split(" ")[0]
		$QuorumStatus=$ClusterStatus = Get-ClusteredMailboxServerStatus -Identity $Cluster.Name | Select -Expand OperationalMachines | ForEach {If($_ -like "*Quorum Owner*") {$_}}
		[string]$QuorumOwner=$QuorumStatus.split(" ")[0]
		Write-Output "Clustered Mailbox Server $NodeName is active on $ActiveNode."
		Write-Output "Cluster Group is active on $QuorumOwner."
		}
	}
elseif ($reg2010)
	{
	If (!$database)
		{
		Write-Output "Please input a database name."
		exit
		}
	Write-Output $stars,"Exchange 2010 detected.",$stars,""
	$check=Get-MailboxDatabase $database -ErrorAction silentlycontinue
	If (!$check)
		{
		Write-Output "No database exists with the name ""$database""."
		exit
		}
	$databasecopystatus=Get-MailboxDatabaseCopyStatus $database | ?{$_.mailboxserver -eq $env:computername}
	$circularlogging=Get-MailboxDatabase $database -Status | select -expandproperty CircularLoggingEnabled
	[array]$databasecopies=Get-MailboxDatabase $database -Status | select -expand databasecopies | select -expand identity | select -expand name
	[string]$servernames=""
	$countmembers=0
	While ($countmembers -lt $databasecopies.count)
		{
		$databasecopy=$databasecopies[$countmembers]
		if ($countmembers -lt ($databasecopies.count - 1))
			{
			$servernames=$servernames+"$databasecopy, "
			}
			else
			{
			$servernames=$servernames+"$databasecopy"
			}
		$countmembers++
		}
	$storageusagestats=Get-StoreUsageStatistics -Database $database | ? { $_.DigestCategory -like "Log*" -and $_.LogRecordBytes -gt 1048576 } | Sort LogRecordBytes -Descending | select DisplayName, LogRecordBytes, LogRecordCount
	$movesinprogress=Get-MoveRequest -MoveStatus InProgress -TargetDatabase $database
	$mailboximports=Get-MailboxImportRequest -Status InProgress -Database $database
	Write-Output $stars,"$Database Copy Status",$stars
	Write-Output ("`nName: {0}`nStatus: {1}`nCopy Queue Length: {2}`nReplay Queue Length: {3}`nLast Inspected Log Time: {4}`nContent Index State: {5}`n" -f $databasecopystatus.name,$databasecopystatus.status,$databasecopystatus.copyqueuelength,$databasecopystatus.replayqueuelength,$databasecopystatus.lastinspectedlogtime,$databasecopystatus.contentindexstate)
	Write-Output $stars,"Checking Circular Logging for $Database",$stars,""
	Write-Output "Circular Logging Enabled: $circularlogging"
	Write-Output "Database Copies: $servernames",""
	Write-Output $stars,"$database Storage Usage Statistice",$stars,"",$storageusagestats,""
	Write-Output $stars,"Checking Mailbox moves to $database",$stars,""
	If ($movesinprogress)
		{
		Write-Output $movesinprogress,""
		}else{
		Write-Output "No moves in progress for database ""$database"".",""
		}
	Write-Output $stars,"Checking Mailbox imports to $database",$stars,""
	If ($mailboximports)
		{
		Write-Output $mailboximports,""
		}else{
		Write-Output "No mailbox imports in progress for database ""$database"".",""
		}
	}
elseif ($reg2013)
	{
	Write-Output "Not implemented yet for 2013"
	}
else
	{
	Write-Output "Exchange is not installed or a valid version of Exchange is not installed."
	}