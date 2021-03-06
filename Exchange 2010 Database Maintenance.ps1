Param([switch]$MaintenanceStart,[switch]$MaintenanceStop,[switch]$BalanceDBsByActivationPreference,[switch]$MoveClusterGroups)
get-pssnapin -registered | add-pssnapin -ErrorAction SilentlyContinue
#region Variables
##Define variables that will be used
$check=$null
$switch=$null
[array]$Primary = @()
[array]$Secondary = @()
[array]$Tertiary = @()
[array]$MXServers = @()
[array]$all = @()
[array]$DBs = @()
[string]$AP = ""
[string]$DB = ""
[int]$countdb=0
[int]$countap=0
[int]$errors=0
[string]$filename=""
[string]$line="--------------------------------------------------"
[String[]]$ActiveNode=""
[String[]]$PassiveNode=""
[String[]]$NodeNames=""
$filename=("C:\Avanade\CAB\{0}-Maintenance.txt" -f $ServerName)
$ServerName=$env:COMPUTERNAME
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ServerName)
$Exch2007="SOFTWARE\\Microsoft\\Exchange\\V8.0"
$Exch2010="SOFTWARE\\Microsoft\\ExchangeServer\\v14"
$Exch2013="SOFTWARE\\Microsoft\\ExchangeServer\\v15"
$reg2007=$reg.OpenSubKey($Exch2007)
$reg2010=$reg.OpenSubKey($Exch2010)
$reg2013=$reg.OpenSubKey($Exch2013)
#endregion

#region Functions
Function GetClusterGroup
{
	$outputs = Invoke-Expression "cluster group"
	$outputs = $outputs[7..($outputs.length)]
	foreach ($output in $Outputs) {
		if ($output)
			{
			$output = $output -replace 'le St','le-St'
			$output = $output -replace 'ter Gr','ter-Gr'
			$output = $output -replace '\s+',' '
			$parts = $output -split ' '
			New-Object -Type PSObject -Property @{
				"Group" = [string]$parts[0]
				"Node" = [string]$parts[1]
				"Status" = [string]$parts[2]
			}
		}
	}
}
Function GetClusterNode
{
	$outputs = Invoke-Expression "cluster node"
	$outputs = $outputs[7..($outputs.length)]
	foreach ($output in $Outputs) {
		if ($output)
			{
			$output = $output -replace '\s+',' '
			$parts = $output -split ' '
			New-Object -Type PSObject -Property @{
				"Node" = [string]$parts[0]
				"Node ID" = [string]$parts[1]
				"Status" = [string]$parts[2]
			}
		}
	}
}
function MakeDBsHealthy()
{
$check = Get-MailboxServer | where {$_.DatabaseAvailabilityGroup -match $DagName} | Get-MailboxDatabaseCopyStatus | ?{$_.Status -ne "Mounted" -and $_.Status -ne "Healthy"}
	If ($check)
		{
		ForEach ($status in $check)
			{
			If ($status.name -match "RDB")
				{
				continue
				}
			If ($status.status -eq "Dismounted")
				{
				$name=$status.name
				Log ("{0} {1} - $name is Dismounted. Attempting to mount database..." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
				Log ($line)
				$databasename=Get-MailboxDatabaseCopyStatus $status.name | select -ExpandProperty DatabaseName
				SilentLog ("Mount-Database $databasename")
				Mount-Database $databasename
				Start-Sleep 5
				$name+=$all
				continue
				}
			If ($status.status -eq "Suspended" -or $status.status -eq "FailedandSuspended" -or $status.status -eq "Failed" -or $status.status -eq "Unicorn")
				{
				$name=$status.name
				$dbstatus=$status.status
				Log ("{0} {1} - $name is $dbstatus. Attempting to resume copy..." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
				Log ($line)
				SilentLog ("Get-MailboxDatabaseCopyStatus $name | Resume-MailboxDatabaseCopy")
				Get-MailboxDatabaseCopyStatus $status.name | Resume-MailboxDatabaseCopy				
				Start-Sleep 5
				$name+=$all
				continue
				}
			}
		Start-Sleep 30
		$check=$null
		ForEach ($name in $all)
			{
			$status = Get-MailboxDatabaseCopyStatus $name | ?{$_.Status -ne "Mounted" -and $_.Status -ne "Healthy"}
			If ($status)
				{
				If ($status.name -match "RDB")
				{
				continue
				}
				$dbstatus=$status.status
				Log ("{0} {1} - $name is still $dbstatus" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
				$check=1
				}
			}
		If ($check)
			{
			Log ("{0} {1} - No further moves will be attempted. Please log into the server to investigate." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			exit
			}
		}
}
# Looks up the DAG and returns an array of servers to contact.
# If serverName is specified, list it first.
# If dagName is specified, then add the dag netname first,
# followed by all of the servers in the dag (excluding serverName)
function Get-ClusterNames(
	[string] $dagName,
	[string] $ServerName
)
{
	$serversToTry = @( );

	# If they specified a server, use it first.
	if ( $ServerName )
	{
		$serversToTry += $ServerName;
	}
<#
	if ( $dagName )
	{
		$dag = @( get-databaseavailabilitygroup $dagName -erroraction silentlycontinue );
		if ( ( ! $dag ) -or ( $dag.Length -eq 0 ) )
		{
			Log ("{1}: {2} was unable to find any DAGs named '{0}'!" -f $dagName,"Get-ClusterNames","get-dag")
		}
		elseif ( $dag.Length -ne 1 )
		{
			Log ("{2}: {3} found {0} DAGs named '{1}'!" -f $dag.Length,$dagName,"Get-ClusterNames","get-dag")
		}
		else
		{
			# If they specified a valid DAG, try the netname before any of
			# the member servers.
			$serversToTry += $dag[0].Name

			foreach ( $server in $dag[0].Servers )
			{
				# Don't add it a second time!
				if ( $server.Name -ne $ServerName )
				{
					$serversToTry += $server.Name
				}
			}
		}

	}
#>
	silentlog ("{3} dagName='{0}' serverName='{1}' is returning '{2}'." -f $dagName,$ServerName,"$serversToTry","Get-ClusterNames")
	return $serversToTry;
}

# Common function to call cluster.exe
# Will try to connect to $dagName or $ServerName to execute the command. If it
# fails for certain reasons (such as the server being unavailable) it will
# retry on the other machines in the cluster.
#
# Parameters:
#  $dagName. Name of the DAG. Optional.
#  $ServerName. Name of the server to use. Optional.
#   One of $dagName or $ServerName MUST be supplied.
#  $clusterCommand. The command to execute.
#
# Returns:
#  @( $errorCode, $textOutput )
function Call-ClusterExe(
	[string] $dagName,
	[string] $ServerName,
	[string] $clusterCommand
)
{

	$script:exitCode = -1;
	$script:textOutput = $null;

	$exitCode = -1;
	$textOutput = $null;

	if ( ( ! $dagName ) -and ( ! $ServerName ) )
	{
		log "Call-ClusterExe was called with neither a dag name nor a server name!"
		exit
	}

	$namesToTry = @( Get-ClusterNames -dagName $dagName -serverName $ServerName )

	foreach ( $nodeName in $namesToTry )
	{

		# Start a script block so that $error.Clear doesn't affect
		# callers, and ErrorActionPreference is unchanged.
		&{
			# Simply specifying -erroraction is not enough.
			$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

			# Method B
			$clusCommand = "cluster.exe /cluster:$nodeName $clusterCommand 2>&1"
			SilentLog ("Running '{0}' on {1}." -f $clusCommand,$nodeName);

			# Run the command.
			$script:textOutput = invoke-expression $clusCommand -erroraction:silentlycontinue

			# We're handling the errors with the exit codes.
			$error.Clear()

			# Save the exit code.
			$script:exitCode = $LastExitCode

			# Superfluous verbose logging.
			SilentLog ("{2} inner block: cluster.exe returned {0}. Output is '{1}'." -f $script:exitCode,"$script:textOutput","Call-ClusterExe")
		}

		# Convert from $script scope to locals.
		$exitCode=$script:exitCode
		$textOutput=$script:textOutput

		SilentLog ("{2}: cluster.exe returned {0}. Output is '{1}'." -f $exitCode,"$textOutput","Call-ClusterExe")

		if ( $LastExitCode -eq 1722 )
		{
			# 1722 is RPC_S_SERVER_UNAVAILABLE
			Log ("{1}: Could not contact the server name {0}. RPC_S_SERVER_UNAVAILABLE (usually means the server is down)." -f $ServerName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -eq 1753 )
		{
			# 1753 is EPT_S_NOT_REGISTERED
			Log ("{1}: Could not contact the server name {0}. EPT_S_NOT_REGISTERED (usually means the server is up, but clussvc is down)." -f $ServerName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -eq 1825 )
		{
			# 1825 is RPC_S_SEC_PKG_ERROR
			Log ("{1}: Could not contact the server name {0}. RPC_S_SEC_PKG_ERROR (usually means the net name resource is down)." -f $ServerName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -ne 0 )
		{
			Log ("{1}: cluster.exe did not succeed, but {0} was not a {2} error code. Not attempting any other servers. This may be an expected error by the caller." -f $LastExitCode,"Call-ClusterExe","retry-able")
			break;
		}
		else
		{
			# It returned 0.
			break;
		}
	}

	return @( $exitCode, $textOutput );
}

##function to log activities - anything that gets changed will be logged in the path C:\Avanade\CAB\
function Log 
{
	param([string]$text)
	#Output to logfile
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
	#Output to screen
	Write-Output $text 
}
function SilentLog 
{
	param([string]$text)
	#Output to logfile
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
}
function BalanceDBs 
{
	Log ($line)
	Log ("{0} {1} - Balancing Databases by Activation Preference" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	
	$countmembers=0
	$check=Get-MailboxServer | ? {$_.DatabaseAvailabilityGroup -match $DagName} | Get-MailboxDatabase | Get-MailboxDatabaseCopyStatus | ?{$_.ContentIndexState -ne "healthy" }| select -ExpandProperty Name
	If ($check)
		{
		$failedcontentindex=1
		Log ("{0} {1} - Some databases have a content index state of Failed. Moving databases..."-f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
		}
	While ($countmembers -lt $Primary.count)
		{
		$check=$null
		$Server=$Primary[$countmembers]
		$DB=$DBs[$countmembers]
		If ($DB -match "RDB")
			{
			$countmembers++
			continue
			}
		$MountedOn=(Get-MailboxDatabase $DB -status).mountedonserver
		$MountedOn=(Get-MailboxServer $MountedOn).name
		If ($MountedOn -ne $Server)
			{
			Log ($line)
			Log ("{0} {1} - $DB is mounted on $MountedOn - attempting to move database to Primary node $Server" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			$check = Get-MailboxDatabaseCopyStatus -identity $DB | ?{$_.copyqueuelength -gt 10 -or $_.replayqueuelength -gt 10}
			If ($check)
				{
				Log ("{0} {1} - $DB has a Copy Queue or Replay Queue length of more than 10. Pausing for 30 seconds and checking again." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
				Start-Sleep -s 30
				$check = Get-MailboxDatabaseCopyStatus -identity $DB | ?{$_.copyqueuelength -gt 10 -or $_.replayqueuelength -gt 10}
				If ($check)
					{
					Log ("{0} {1} - $DB still has a Copy Queue or Replay Queue length of more than 10. No further moves will be attempted." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
					exit
					}
				}
			SilentLog ("{0} {1} - Move-ActiveMailboxDatabase -Identity $DB -ActivateOnServer $Server" -f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
			If ($failedcontentindex)
			{
			$Move = & Move-ActiveMailboxDatabase -Identity $DB -ActivateOnServer $Server -SkipClientExperienceChecks -confirm:$false | ?{$_.Status -ne "Succeeded"}
			}else{
			$Move = & Move-ActiveMailboxDatabase -Identity $DB -ActivateOnServer $Server -confirm:$false | ?{$_.Status -ne "Succeeded"}
			}
			If ($Move){
			Log ("{0} {1} - Failed to move database." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			$errors++
			}else{
			Log ("{0} {1} - $DB Successfully moved to $Server" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			$all+=$DB
			}
			}else{
			Log ($line)
			Log ("{0} {1} - $DB is mounted on Primary node $Server" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			}
			
		$countmembers++
		}
	If ($errors)
	{
	Log ("{0} {1} - $errors error(s) were encountered while attemting database moves. Please log into server to investigate." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	}
	else
	{
	Log ("{0} {1} - All databases have been moved to primary node. Confirming now that all databases are healthy." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))	
	}
	Start-Sleep -Seconds 15
	ForEach ($DB in $all)
		{
		$check = Get-MailboxDatabaseCopyStatus $DB | ?{$_.Status -ne "Mounted" -and $_.Status -ne "Healthy"}
		If ($check)
			{
			$switch=1
			}
		}
	If ($switch)
		{
		Log ("{0} {1} - Some databases are not healthy. Please log onto the server to investigate." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}else{
		Log ("{0} {1} - All databases successfully moved are healthy." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}
	}
#endregion

$Avanade = "C:\Avanade"
$CAB = "C:\Avanade\CAB"
If ((Test-Path $Avanade) -eq 0)
	{
	New-Item -ItemType directory -Path $Avanade | Out-Null
	New-Item -ItemType directory -Path $CAB | Out-Null
	}
If ((Test-Path $CAB) -eq 0)
	{
	New-Item -ItemType directory -Path $CAB | Out-Null
	}

##Checks to make sure that more than one flag is not active.
If ($MaintenanceStart -eq $true)
	{
	$check++
	}
If ($MaintenanceStop -eq $true)
	{
	$check++
	}
If ($BalanceDBsByActivationPreference -eq $true)
	{
	$check++
	}
If ($check -gt 1)
	{
	Write-Output "Please select only one Maintenance or Balance flag."
	exit
	}

If ($reg2007)
	{
	##removed this functionality due to a complication with AUM.
<#2007 failover code
	$clustergroups=GetClusterGroup
	Write-Output $line,"Exchange 2007 detected.",$line
	$Clusters = Get-MailboxServer | ? {$_.ClusteredStorageType -ne "Disabled"}
	ForEach ($Cluster in $Clusters) {
		$NodeNames+=$Cluster.Name
		$ClusterStatus = Get-ClusteredMailboxServerStatus -Identity $Cluster.Name | Select -Expand OperationalMachines | ForEach {If($_ -like "*Active*") {$_}}
		$ActiveNode += $ClusterStatus.Split(" ")[0]
		$ClusterStatus = Get-ClusteredMailboxServerStatus -Identity $Cluster.Name | Select -Expand OperationalMachines | ForEach {If($_ -notlike "*Active*") {$_}}
		$PassiveNode += $ClusterStatus.Split(" ")[0]
		}
	If ($ActiveNode.count -eq 1)
		{
		If ($ActiveNode -ne $ServerName)
			{
			Write-Output "Please run the failover script on the server that is hosting the clustered mailbox server. ($ActiveNode)"
			exit
			}
		}
	ForEach ($Node in $NodeNames)
		{
		If (!$Node)
			{
			continue
			}
		$ndx=[array]::IndexOf($NodeNames,$Node)
		$ANode=$ActiveNode[$ndx]
		$PNode=$PassiveNode[$ndx]
		If ($PNode -eq $ServerName)
			{
			Log ("Clustered Mailbox server ""$Node"" is hosted on $ANode. Not moving ""$Node"".")
			continue
			}
		Log ("Moving clustered mailbox server ""$Node"" to passive node $PNode...")
		SilentLog ("Move-ClusteredMailboxServer -Identity $Node -MoveComment ""Avanade Patching"" -TargetMachine $PNode)
		Move-ClusteredMailboxServer -Identity $Node -MoveComment "Avanade Patching" -TargetMachine $PNode -confirm:$false
		}
	ForEach ($Group in $clustergroups)
		{
		$node=$Group.Node
		$name=$Group.Group
		$status=$Group.status
		$name=$name -replace 'ble-Sto','ble Sto'
		$name=$name -replace 'ter-Gr','ter Gr'
		If ($name -eq "Available Storage" -or $name -eq "Cluster Group")
			{
		If ($node -ne $ServerName)
			{
			Log ("""$name"" is hosted on $node. Not moving ""$name"".")
			continue
			}
		Log ("Moving resource group ""$name"" to node $PNode...")
		$move = & cluster group $name /move:$PNode
			}
		}
#>
	}
ElseIf ($reg2010)
	{
##Grabs DagName of server
$DagName=((Get-MailboxServer $ServerName).databaseavailabilitygroup.name)

##Grabs unique database names from servers in the defined DAG
[array]$DBs = (Get-MailboxServer -status | ? {$_.DatabaseAvailabilityGroup -eq $DagName} | Get-MailboxDatabase | select -ExpandProperty name | select -uniq)

If ($MaintenanceStop -eq $false)
	{
	##This loop will go through each defined databases from above and grab the activation preferences of each - only needs to be ran with maintenance start and balancedb
	While ($countdb -lt $DBs.count)
		{
		If ($DB -ne $DBs[$countdb])
			{
			[string]$DB=$DBs[$countdb]
			[array]$APs = ((Get-MailboxDatabase $DB).activationpreference)
			If ($APs.Count -ne 3)
				{
				If ($APs.Count -eq 2)
					{
					[array]$Tertiary+="nothing"
					}
				If ($APs.Count -eq 1)
					{
					[array]$Tertiary+="nothing"
					[array]$Secondary+="nothing"
					}
				}
			}
		If ($countap -lt $APs.count)
			{
			[string]$AP=$APs[$countap]
			$AP = $AP.substring(1,$AP.Length-5)
			}
		If ($countap -eq 0)
			{
			[array]$Primary+=$AP
			}
		If ($countap -eq 1)
			{
			[array]$Secondary+=$AP
			}
		If ($countap -eq 2)
			{
			[array]$Tertiary+=$AP
			}
		$countap++
		If ($countap -eq $APs.count)
			{
			$countdb++
			[int]$countap = 0
			}
		}
	}
If ($BalanceDBsByActivationPreference -eq $true)
	{
	MakeDBsHealthy
	BalanceDBs
	}
If ($MaintenanceStart -eq $true)
	{
	Log ($line)
	Log ("{0} {1} - Starting Maintenance on $ServerName" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	Log ($line)
	MakeDBsHealthy
	$check = Get-MailboxDatabaseCopyStatus -Server $ServerName | ?{$_.copyqueuelength -gt 10 -or $_.replayqueuelength -gt 10}
	If ($check)
		{
		Log ("{0} {1} - $ServerName has a Copy Queue or Replay Queue length of more than 10. Pausing for 30 seconds and checking again." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		Start-Sleep -s 30
		$check = Get-MailboxDatabaseCopyStatus -Server $ServerName | ?{$_.copyqueuelength -gt 10 -or $_.replayqueuelength -gt 10}
		If ($check)
			{
			Log ("{0} {1} - $ServerName still has a Copy Queue or Replay Queue length of more than 10. No further moves will be attempted." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
			exit
			}
		}
	$dagName = $null
	$mbxServer = get-mailboxserver $ServerName -erroraction:silentlycontinue
	if ($mbxServer -and $mbxServer.DatabaseAvailabilityGroup)
		{
		$dagName = $mbxServer.DatabaseAvailabilityGroup.Name;
		}
	# Start with $ServerName (which may or may not be a FQDN) before
	# falling back to the (short) names of the DAG.
	$outputStruct = Call-ClusterExe -dagName $dagName -serverName $ServerName -clusterCommand "node $ServerName /pause"
	$LastExitCode = $outputStruct[ 0 ];
	$output = $outputStruct[ 1 ];
	if ($LastExitCode -eq 1753)
		{
		# 1753 is EPT_S_NOT_REGISTERED
		silentlog ("{1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED." -f $ServerName,"Start-DagServerMaintenance")
		}
		elseif ($LastExitCode -eq 1722)
		{
		# 1722 is RPC_S_SERVER_UNAVAILABLE
		silentlog ("{1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE." -f $ServerName,"Start-DagServerMaintenance")
		}
		elseif ($LastExitCode -ne 0)
		{
		Log ("{2}: Failed to suspend the server {0} from hosting the Primary Active Manager, returned {1}." -f $ServerName,$LastExitCode,"Start-DagServerMaintenance")
		exit
		}
	Log ("{0} {1} - Moving Databases off of $ServerName" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	$check=Get-MailboxServer | ? {$_.DatabaseAvailabilityGroup -match $DagName} | Get-MailboxDatabase | Get-MailboxDatabaseCopyStatus | ?{$_.ContentIndexState -eq "Failed" }| select -ExpandProperty Name
	If ($check)
		{
		Log ("{0} {1} - Some databases have a content index state of Failed. Moving databases..."-f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
		$Move=$null
		#Moves databases with failed content index states
		ForEach ($name in $check)
			{
			$activecopy=Get-MailboxDatabaseCopyStatus $name | select -ExpandProperty ActiveDatabaseCopy
			$dbname=Get-MailboxDatabaseCopyStatus $name | select -ExpandProperty DatabaseName
			If ($dbname -match "RDB")
				{
				continue
				}
			$ndx=[array]::IndexOf($DBs,$dbname)
			If ($activecopy -eq $ServerName)
				{
				If ($ServerName -eq $Primary[$ndx])
					{
					SilentLog ("{0} {1} - Move-ActiveMailboxDatabase $dbname -activateonserver $Secondary[$ndx] -skipclientexperiencechecks" -f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
					$Move = & Move-ActiveMailboxDatabase $dbname -activateonserver $Secondary[$ndx] -skipclientexperiencechecks -confirm:$false | ?{$_.Status -ne "Succeeded"}
					}else{
					SilentLog ("{0} {1} - Move-ActiveMailboxDatabase $dbname -activateonserver $Primary[$ndx] -skipclientexperiencechecks" -f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
					$Move = & Move-ActiveMailboxDatabase $dbname -activateonserver $Primary[$ndx] -skipclientexperiencechecks -confirm:$false | ?{$_.Status -ne "Succeeded"}
					}
				If ($Move)
					{
					$errors++
					}
				}
			}
		}
		#moves the rest of the databases - it moves nothing if there were content index failed statuses on all of the databases - this will move anything that has a healthy content index
		#that was missed by the unhealthy check.
		[array]$DatabasesOnServer=Get-MailboxServer $ServerName | Get-MailboxDatabaseCopyStatus | ?{$_.Status -eq "Mounted"} | Select -ExpandProperty Name
		If ($DatabasesOnServer)
			{
			ForEach ($name in $DatabasesOnServer)
				{
				$activecopy=Get-MailboxDatabaseCopyStatus $name | select -ExpandProperty ActiveDatabaseCopy
				$dbname=Get-MailboxDatabaseCopyStatus $name | select -ExpandProperty DatabaseName
				$ndx=[array]::IndexOf($DBs,$dbname)
				If ($Secondary[$ndx] -eq "nothing" -or $dbname -match "RDB")
					{
					continue
					}
				If ($activecopy -eq $ServerName)
					{
					If ($ServerName -eq $Primary[$ndx])
						{
						$activateonserver=$Secondary[$ndx]
						SilentLog ("{0} {1} - Move-ActiveMailboxDatabase $dbname -activateonserver $activateonserver" -f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
						$Move = & Move-ActiveMailboxDatabase $dbname -activateonserver $activateonserver -confirm:$false | ?{$_.Status -ne "Succeeded"}
						}else{
						$activateonserver=$Primary[$ndx]
						SilentLog ("{0} {1} - Move-ActiveMailboxDatabase $dbname -activateonserver $activateonserver" -f (Get-Date).tostring("MM-dd-yyyy"), (Get-Date).tostring("HH:mm:ss"))
						$Move = & Move-ActiveMailboxDatabase $dbname -activateonserver $activateonserver -confirm:$false | ?{$_.Status -ne "Succeeded"}
						}
					If ($Move)
						{
						$errors++
						}
					}
				}
			}
		
	If ($errors)
		{
		Log ("{0} {1} - Errors were encountered while attemting database moves. Please log into server to investigate. Checking to see if databases are healthy." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}else{
		Log ("{0} {1} - All databases have been moved off of server $ServerName. Confirming now that all databases are healthy." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))	
		}
	Start-Sleep -Seconds 25
	$check = Get-MailboxDatabaseCopyStatus -Server $ServerName | ?{$_.Status -ne "Mounted" -and $_.Status -ne "Healthy"}
	If ($check)
		{
		Log ("{0} {1} - Some databases on $ServerName are not healthy. Please log onto the server to investigate." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}else{
		Log ("{0} {1} - All databases on $ServerName are healthy." -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}
	}
If ($MaintenanceStop -eq $true)
	{
	Log ($line)
	Log ("{0} {1} - Stopping Maintenance on $ServerName" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	Log ($line)
	Import-Module failoverclusters
	$clusternodes=get-clusternode
	Foreach ($node in $clusternodes)
		{
		If ($node.state -ne "Up")
			{
			$nodecheck = 1
			}
		}
	If ($nodecheck -ne 1)
		{
		Log ("{0} {1} - Maintenance is already stopped on $ServerName" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
		}
		else
		{
	# Try to fetch $dagName if we can.
	$dagName = $null
	$mbxServer = get-mailboxserver $ServerName -erroraction:silentlycontinue
	if ( $mbxServer -and $mbxServer.DatabaseAvailabilityGroup )
	{
		$dagName = $mbxServer.DatabaseAvailabilityGroup.Name;
	}
	
	$outputStruct = Call-ClusterExe -dagName $dagName -serverName $ServerName -clusterCommand "node $ServerName /resume"
	$LastExitCode = $outputStruct[ 0 ];
	$output = $outputStruct[ 1 ];

	# 0 is success, 5058 is ERROR_CLUSTER_NODE_NOT_PAUSED.
	if ( $LastExitCode -eq 5058 )
	{
		silentlog ("The server {0} is already able to host the Primary Active Manager." -f $ServerName)
	}
	elseif ( $LastExitCode -eq 1753 )
	{
		# 1753 is EPT_S_NOT_REGISTERED
		silentlog ("{1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED." -f $severName,"Stop-DagServerMaintenance")
	}
	elseif ( $LastExitCode -eq 1722 )
	{
		# 1722 is RPC_S_SERVER_UNAVAILABLE
		silentlog ("{1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE." -f $severName,"Stop-DagServerMaintenance")
	}
	elseif ( $LastExitCode -eq 0 )
	{
		Log ("{0} {1} - Successfully stopped maintenance on $ServerName" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	}
	elseif ( $LastExitCode -ne 0 )
	{
		Log ("{4}: Failed to resume the ability of the server {0} to host the Primary Active Manager, 'cluster.exe /cluster:{1} node {2} /resume' returned {3}." -f $ServerName,$ServerName,$shortServerName,$LastExitCode,"Start-DagServerMaintenance") -stop 
	}
	}
	}
If ($MoveClusterGroups -eq $true)
	{
	Log ($line)
	Log ("{0} {1} - Moving Cluster Resources" -f (Get-Date).tostring("MM-dd-yyyy"),(Get-Date).tostring("HH:mm:ss"))
	Log ($line)
	Import-Module failoverclusters
	$clustergroups=Get-ClusterGroup
	Foreach ($group in $clustergroups)
		{
		$name=$group.name
		$node=$group.ownernode
		$state=$group.state
		If ($name -eq "Cluster Group" -or $name -eq "Available Storage")
			{
			If ($node -match $ServerName)
				{
				$pnodecheck=Get-ClusterNode | ?{$_.name -ne $ServerName} | select -First 1 | select -ExpandProperty name
				If ($pnodecheck)
					{
					$pnode=Get-ClusterNode | ?{$_.name -ne $ServerName -and $_.State -eq 'Up'} | select -First 1 | select -ExpandProperty name
					If ($pnode)
						{
						Log ("Moving resource group ""$name"" to node $pnode...")
						$Move = & Move-ClusterGroup $name -Node $pnode
						}else{
						$pnodestate=Get-ClusterNode | ?{$_.name -ne $ServerName} | select -First 1 | select -ExpandProperty state
						Log ("Node $pnodecheck is currently $pnodestate. No moves will be attempted for group ""$name"".")
						}
					}
					else
					{
					Log ("""$name"" is hosted on $node. Not moving ""$name"".")
					}
					
				}
			}
			else
			{
			Log ("***Not moving resource group ""$name""***")
			}
		}
	}
}
ElseIf ($Exch2013)
	{
	Write-Output "Not yet implemented for Exchange 2013"

	}
else
	{
	Write-Output "Exchange is not installed on this machine."
	}