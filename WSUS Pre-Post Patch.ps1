Param([switch]$Prepatch,[switch]$Midpatch,[switch]$Postpatch,[string]$TargetServer)
get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue
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
#endregion
#github test
$UseClusterEXE=$null
$ServerName=$Env:COMPUTERNAME
$SkipExchangeTests=$false
$filename="C:\Avanade\CAB\$az_casenumber - Services.csv"
$filename1="C:\Avanade\CAB\$az_casenumber - Cluster.csv"
$filename2="C:\Avanade\CAB\$az_casenumber - Databases.csv"
$Avanade = "C:\Avanade"
$CAB = "C:\Avanade\CAB"
$WSUS = "C:\Avanade\Install Reports"
$check=0
$line="-----------------------------"

If ($TargetServer)
	{
	$SecureX=Get-Service -ComputerName $TargetServer -Name Azaleos.SecureXAgent5 -ErrorAction SilentlyContinue
	If ($SecureX)
		{
		$Status=Get-Service -ComputerName $TargetServer -Name Azaleos.SecureXAgent5 -ErrorAction SilentlyContinue | select -ExpandProperty status
		If ($Status -eq "Running")
			{
			Write-Output "SecureX Service is Running on server $TargetServer. Restarting service now..."
			Stop-Service -InputObject (Get-Service -ComputerName $TargetServer -Name Azaleos.SecureXAgent5 -ErrorAction SilentlyContinue)
			Start-Service -InputObject (Get-Service -ComputerName $TargetServer -Name Azaleos.SecureXAgent5 -ErrorAction SilentlyContinue)
			}
		If ($Status -eq "Stoppped")
			{
			Write-Output "SecureX Service is Stopped on server $TargetServer. Starting service now..."
			Start-Service -InputObject (Get-Service -ComputerName $TargetServer -Name Azaleos.SecureXAgent5 -ErrorAction SilentlyContinue)
			}
		If ($Status -eq "Starting")
			{
			Write-Output "SecureX Service is in the process of starting."
			}
		If ($Status -eq "Stopping")
			{
			Write-Output "SecureX Service is in the process of stopping."
			}
		}else{
		Write-Output "Unable to reach indicated server to restart SecureX service."
		}
	exit
	}

If ($Prepatch -eq $true)
	{
	$check++
	}
If ($Midpatch -eq $true)
	{
	$check++
	}
If ($Postpatch -eq $true)
	{
	$check++
	}
If ($check -gt 1)
	{
	Write-Output "Please select only one patch flag."
	exit
	}
$check=0
If ((Test-Path $Avanade) -eq 0)
{
New-Item -ItemType directory -Path $Avanade | Out-Null
New-Item -ItemType directory -Path $CAB | Out-Null
New-Item -ItemType directory -Path $WSUS | Out-Null
}
If ((Test-Path $CAB) -eq 0)
{
New-Item -ItemType directory -Path $CAB | Out-Null
}
If ((Test-Path $WSUS) -eq 0)
{
New-Item -ItemType directory -Path $WSUS | Out-Null
}
If ($Postpatch -eq $true -or $Midpatch -eq $true)
	{
	If ((Test-Path $filename) -eq 0)
		{
		Write-Output "-----------------------------------------------------------------------------","Prepatch services file not found. Please run Prepatch flag prior to patching.","-----------------------------------------------------------------------------"
		exit
		}
	}

##checks for BES services
$BB5=Get-Service "Blackberry Router" -erroraction silentlycontinue
$BB10=Get-Service "BES - BlackBerry Dispatcher" -erroraction silentlycontinue

##checks for Cluster services
$Clus=Get-Service ClusSvc -erroraction silentlycontinue

##checks for Exchange services
$Exch=Get-Service *MSExchange* -ErrorAction silentlycontinue

##checks for Lync Services
$Lync=Get-Service rtcsrv -ErrorAction SilentlyContinue

##checks for DPM Services
$DPM10=Get-Service 'MSSQL$MSDPM2010' -erroraction SilentlyContinue
$DPM12=Get-Service 'MSSQL$MSDPM2012' -ErrorAction SilentlyContinue

##variables to check local registry
$Server = $env:COMPUTERNAME
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Server)

If ($Exch)
	{
	#exchange registry values to check to see what version is installed - 2013 will not work in AUM yet, so it does not do anything more than stop AZMON
	$Exch2007="SOFTWARE\\Microsoft\\Exchange\\V8.0"
	$Exch2010="SOFTWARE\\Microsoft\\ExchangeServer\\v14"
	$Exch2013="SOFTWARE\\Microsoft\\ExchangeServer\\v15"
	$reg2007=$reg.OpenSubKey($Exch2007)
	$reg2010=$reg.OpenSubKey($Exch2010)
	$reg2013=$reg.OpenSubKey($Exch2013)
	If ($reg2007)
		{
		$MailboxRoleInstalled=$reg2007.OpenSubKey("MailboxRole")
		}
	If ($reg2010)
		{
		$MailboxRoleInstalled=$reg2010.OpenSubKey("MailboxRole")
		If ($Prepatch -eq $true)
			{
			If ($MailboxRoleInstalled)
				{
				##Grabs DagName of server
				$DagName=((Get-MailboxServer $ServerName).databaseavailabilitygroup.name)

				##Grabs unique database names from servers in the defined DAG
				[array]$DBs = (Get-MailboxServer -status | ? {$_.DatabaseAvailabilityGroup -eq $DagName} | Get-MailboxDatabase | select -ExpandProperty name | select -uniq)
		
				$AllInfo = $DBs | ForEach-Object {
				$MountedOnServer=Get-MailboxDatabase $_ -status | select -ExpandProperty MountedOnServer
				New-Object PSObject -Property @{
				"Name" = "{0}" -f $_
				"MountedOnServer" = "{0}" -f $MountedOnServer
						}
					}
				$AllInfo | Export-Csv -Path $filename2 -NoTypeInformation
				}
			}
		}
	If ($reg2013)
		{
		$MailboxRoleInstalled=$reg2013.OpenSubKey("MailboxRole")
		$SkipExchangeTests=$true
		}
	}

If ($Prepatch -eq $true)
{
get-wmiobject win32_service | select DisplayName, State, StartMode| export-csv $filename -NoTypeInformation
Write-Output "Service status saved at $filename",$line

If ($DPM10)
	{
	Write-Output "DPM2010 Role is detected."
	Write-Output "Stopping and Disabling Azaleos Monitoring Service..."
	Stop-Service AzaleosMonitoringService
	Get-Service AzaleosMonitoringService | Set-Service -StartupType Disabled
	$name=Get-Service -Name 'MSSQL$MSDPM2010' | select -ExpandProperty DisplayName
	Write-Output "Stopping and Disabling $name...",$line
	$DepServices = Get-Service -Name 'MSSQL$MSDPM2010' -DependentServices | ?{$_.status -eq 'Running'}
	If ($DepServices)
		{
	foreach ($Service in $DepServices)
		{
		Stop-Service -InputObject (Get-Service -Name $Service.Name) -Force
		}
		}
	Stop-Service -InputObject (Get-Service -Name 'MSSQL$MSDPM2010') -Force
	Get-Service 'MSSQL$MSDPM2010' | Set-Service -StartupType Disabled
	}
If ($DPM12)
	{
	Write-Output "DPM2012 Role was detected."
	Write-Output "Stopping and Disabling Azaleos Monitoring Service...",$line
	Stop-Service AzaleosMonitoringService
	Get-Service AzaleosMonitoringService | Set-Service -StartupType Disabled
	$name=Get-Service -Name 'MSSQL$MSDPM2012' | select -ExpandProperty DisplayName
	Write-Output "Stopping and Disabling $name...",$line
	$DepServices = Get-Service -Name 'MSSQL$MSDPM2012' -DependentServices | ?{$_.status -eq 'Running'}
	If ($DepServices)
		{
	foreach ($Service in $DepServices)
		{
		Stop-Service -InputObject (Get-Service -Name $Service.name) -Force
		}
		}
	Stop-Service -InputObject (Get-Service -Name 'MSSQL$MSDPM2012') -Force
	Get-Service 'MSSQL$MSDPM2012' | Set-Service -StartupType Disabled
	}
If ($Exch)
	{
	If ($MailboxRoleInstalled)
		{
		Write-Output "Mailbox Role is detected.`nStopping and Disabling Azaleos Monitoring Service...",$line
		Stop-Service AzaleosMonitoringService
		Get-Service AzaleosMonitoringService | Set-Service -StartupType Disabled
		}
	}
If ($Lync)
	{
	Write-Output "Lync Front End detected.`nStopping Azaleos Monitoring Service...",$line
	Stop-Service AzaleosMonitoringService
	}
If ($Clus)
	{
	$clustergroup=GetClusterGroup
	$clusternode=GetClusterNode
	[array]$nodenames=$clusternode | select -ExpandProperty node
	[array]$state=$clusternode | select -ExpandProperty status
	[array]$groups=$clustergroup | select -ExpandProperty group
	[array]$nodes=$clustergroup | select -expandproperty node
	[array]$statuses=$clustergroup| select -ExpandProperty status
	Foreach ($group in $groups)
		{
		If ($group -eq "Available-Storage")
			{
			$ndx=[array]::IndexOf($groups, $group)
			$groups[$ndx]="Available Storage"
			continue
			}
		If ($group -eq "Cluster-Group")
			{
			$ndx=[array]::IndexOf($groups, $group)
			$groups[$ndx]="Cluster Group"
			}
		}
	If ($groups.count -gt $nodenames.count)
		{
		While ($groups.count -gt $nodenames.count)
			{
			$nodenames+="0"
			$state+="0"
			}
		}else{
		While ($nodenames.count -gt $groups.count)
			{
			$groups+="0"
			$nodes+="0"
			$statuses+="0"
			}
		}
	$AllInfo = $groups | ForEach-Object {
		$ndx=[array]::IndexOf($groups, $_)
		New-Object PSObject -Property @{
		"Group" = "{0}" -f $_
		"Node" = "{0}" -f $nodes[$ndx]
		"Status" = "{0}" -f $statuses[$ndx]
		"Node Names" = "{0}" -f $nodenames[$ndx]
		"Node State" = "{0}" -f $state[$ndx]
			}
		}
	$AllInfo | Export-Csv $filename1 -NoTypeInformation
	Write-Output "Cluster status saved at $filename1",$line
	}
If ($Exch)
	{
	If ($reg2010)
		{
		If ($MailboxRoleInstalled)
			{Write-Output "Database statuses saved at $filename2",$line}
		}
	}
}

If ($Postpatch -eq $true -or $Midpatch -eq $true)
{
[String[]]$Ignore = @("Microsoft .NET Framework NGEN v4.0.30319_X64", "Microsoft .NET Framework NGEN v4.0.30319_X86", "Performance Logs and Alerts", "Shell Hardware Detection", "Software Protection")
[string[]]$skip=@("Azaleos Monitoring Service","SQL Server (MSDPM2010)","SQL Server (MSDPM2012)","SQL Server Agent (MSDPM2010)","SQL Server Agent (MSDPM2012)","DPM (MSDPM)","DPM")
$Status = Import-Csv $filename | select -ExpandProperty State
$DisplayName = Import-Csv $filename | select -ExpandProperty DisplayName
$StartMode = Import-Csv $filename | select -ExpandProperty StartMode
If ($Postpatch -eq $true)
	{
ForEach ($Name in $DisplayName)
	{
	$ndx=[array]::IndexOf($DisplayName, $Name)
	$Start=Get-WmiObject win32_service | ?{$_.displayname -eq $Name} | select -ExpandProperty startmode
	If ($Ignore -contains $Name)
		{
		continue
		}
	If ($skip -contains $Name)
		{
		If ($Start -ne "Auto")
			{
			Write-Output "Service $Name is set to $Start. It is being set to Automatic..."
			Get-Service $Name | Set-Service -StartupType "Automatic"
			}
		continue
		}
	If ($Start -eq $StartMode[$ndx])
		{
		continue
		}else{
		$StartModeTemp=$StartMode[$ndx]
		If ($StartModeTemp -eq "Auto")
			{
			$StartModeTemp="Automatic"
			}
		Write-Output "Service $Name is set to $Start. It was set to $StartModeTemp. Changing start mode now..."
		Get-Service $Name | Set-Service -StartupType $StartModeTemp
		}
	}
	}else{
	ForEach ($Name in $DisplayName)
		{
		If ($Ignore -contains $Name)
			{
			continue
			}
		If ($skip -contains $Name)
			{
			continue
			}
		$ndx=[array]::IndexOf($DisplayName, $Name)
		$Start=Get-WmiObject win32_service | ?{$_.displayname -eq $Name} | select -ExpandProperty startmode
		If ($Start -eq $StartMode[$ndx])
			{
			continue
			}else{
			$StartModeTemp=$StartMode[$ndx]
			If ($StartModeTemp -eq "Auto")
				{
				$StartModeTemp="Automatic"
				}
			Write-Output "Service $Name is set to $Start. It was set to $StartModeTemp. Changing start mode now..."
			Get-Service $Name | Set-Service -StartupType $StartModeTemp
			}
		}
	}
If ($BB5 -or $BB10)
	{
	$BB = Get-Service "Blackberry Router" -ErrorAction SilentlyContinue
	If ($BB -ne $null)
		{
		$BesServicesStart = @("BlackBerry Router","BlackBerry Dispatcher","BlackBerry Controller") 
		}else{
		$BesServicesStart = @("BES - BlackBerry Dispatcher","BES - BlackBerry Controller")
		}
	foreach ($Service in $BesServicesStart)
		{
		$where = [array]::IndexOf($DisplayName, $Service)
		$checkservice=(Get-Service $Service).status
		If ($Status[$where] -eq "Running" -and $checkservice -ne "Running")
			{
			Write-Output "Starting $Service..."
			Get-Service $Service | Start-Service -WarningAction SilentlyContinue | Out-Null
			$line
			}
		}
	foreach ($Service in $BesServicesStart)
		{
		$checkservice=(Get-Service $Service).status
		If ($checkservice -ne "Running")
			{
			Write-Output "BES service $Service failed to start - log into server to investigate.",$line
			}
		}
	}
	
If ($Exch)
	{
	If ($SkipExchangeTests -eq $false)
		{
		$testhealth = test-servicehealth | select -ExpandProperty requiredservicesrunning
		If ($testhealth -contains $false)
			{
			Write-Output "Test-ServiceHealth command indicates some Exchange services are not running. Starting services..."
			$testhealth = Test-ServiceHealth | select -ExpandProperty servicesnotrunning | select -Unique
			If ($testhealth -match "MSExchangeADTopology")
				{
				$DepServices = Get-Service -Name MSExchangeADTopology -DependentServices | ? {$_.status -eq "Stopped"}
				If ($DepServices)
					{
					Foreach ($DepService in $DepServices)
						{
						$Name=$DepService.DisplayName
						Write-Output "Starting $Name..."
						Start-Service -InputObject (Get-Service -Name $DepService.Name) | Out-Null
						}
					}
				$test=Get-Service -Name MSExchangeADTopology
				If ($test.status -eq "Stopped")
					{
					Write-Output "Starting Microsoft Exchange Active Directory Topology..."
					Start-Service -InputObject (Get-Service -Name MSExchangeADTopology) | Out-Null
					}
				$testhealth = Test-ServiceHealth | select -ExpandProperty servicesnotrunning | select -Unique
				}
			If ($testhealth)
				{
				Foreach ($Service in $testhealth)
					{
					Write-Output "Starting $Service..."
					Start-Service -InputObject (Get-Service -Name $Service) | Out-Null
					}
				}
			}
		If ($Midpatch -eq $true)
			{
			$badevent=Get-EventLog -LogName application -after (get-date).AddDays(-1)|?{$_.eventid -eq 2601 -or $_.eventid -eq 2604 -or $_.eventid -eq 2501}
			If ($badevent)
				{
				$line
				Write-Output "Found event 2601 or 2501 or 2604. This event requires a restart of Exchange AD Topology Service. Restarting service and its dependant services now..."
				$DepServices = Get-Service -Name MSExchangeADTopology -DependentServices
				If ($DepServices)
					{
					Foreach ($DepService in $DepServices)
						{
						$Name=$DepService.DisplayName
						Write-Output "Stopping $Name..."
						Stop-Service -InputObject (Get-Service -Name $DepService.Name) -Force | Out-Null
						}
					}
				Write-Output "Restarting Exchange AD Topology..."	
				Restart-Service -InputObject (Get-Service -Name MSExchangeADTopology) -Force | Out-Null
				If ($DepServices)
					{
					Foreach ($DepService in $DepServices)
						{
						$Name=$DepService.DisplayName
						Write-Output "Starting $Name..."
						Start-Service -InputObject (Get-Service -Name $DepService.Name) | Out-Null
						}
					}
				$line
				}
			}
		}
	}

[int]$countmembers = 0
[int]$check = 0
While ($countmembers -lt $DisplayName.count)
	{
	$D = $DisplayName[$countmembers]
	$S = $Status[$countmembers]
	$SM = $StartMode[$countmembers]
	If ($Ignore -contains $Name)
		{
		$countmembers++
		continue
		}
	If ($Midpatch -eq $true)
		{
		If ($skip -contains $D)
			{
			$countmembers++
			continue
			}
		}
	$servicestatus = (Get-Service $D).status
	If ($skip -contains $D)
		{
		If ($servicestatus -match "Stopped")
			{
			Write-Output "$D is stopped. Starting service now...",$line
			Get-Service $D | Start-Service -WarningAction SilentlyContinue | Out-Null
			}
		$countmembers++
		continue
		}
	If ($SM -ne "Auto")
		{
		$countmembers++
		continue
		}
	If ($servicestatus -ne $S)
		{
		if ($S -eq "Running")
			{
			Write-Output "$D was running pre-patch. Starting service now..."
			Get-Service $D | Start-Service -WarningAction SilentlyContinue | Out-Null
			$checkstatus = "Running"
			}
		if ($S -eq "Stopped")
			{
			Write-Output "$D was stoppped pre-patch. Stoping service now..."
			Get-Service $D | Stop-Service -WarningAction SilentlyContinue| Out-Null
			$checkstatus = "Stopped"
			}
		$servicestatus = (Get-Service $D).status
		If ($servicestatus -eq $checkstatus)
			{
			Write-Output "$D was set to $checkstatus successfully",$line
			}
			else
			{
			Write-Output "$D was not successfully set to $checkstatus, please log into server $env:computername and investigate",$line
			$check++
			}
		}
	$countmembers++
	}

If ($check -eq 0)
	{
	Write-Output "All services have been set to pre-patch statuses successfully.",$line
	}
	else
	{
	If ($check -eq 1)
		{
		Write-Output "1 error occured. Please log into server to investigate."$line
		}else{
		Write-Output "$check errors occured. Please log into server to inestigate."$line
		}
	}
	
##Checks Cluster status
If ($Clus)
	{
	[array]$clustergroup=Import-Csv $filename1 | select -ExpandProperty Group
	[array]$clusternode=Import-Csv $filename1 | select -ExpandProperty Node
	[array]$clusterstatus=Import-Csv $filename1 | select -ExpandProperty status
	[array]$clusternodename=Import-Csv $filename1 | select -ExpandProperty 'node names'
	[array]$clusternodestate=Import-Csv $filename1 | select -ExpandProperty 'node state'
	[array]$currentclusterstatus=getclustergroup | ?{$_.group -eq "Available-Storage" -or $_.group -eq "Cluster-Group"}
	[array]$currentclusternodestatus=getclusternode
<#
	Foreach ($currentcluster in $currentclusternodestatus)
		{
		$nodename=$currentcluster.node
		$ndx=[array]::IndexOf($clusternodename,$nodename)
		If ($clusternodestate[$ndx] -ne $currentcluster.status)
			{
			$currentstate=$currentcluster.status
			$prestate=$clusternodestate[$ndx]
			Write-Output "Cluster node $nodename is currently $currentstate. Prepatch it was $prestate. Changing state now..."
			
			}
		}
#>
	If ($clustergroup.count -gt 2)
		{
		Write-Output 'This script will only edit cluster groups "Cluster Group" and "Available Storage".','There are other cluster groups installed, please log into server to ensure they are set to the proper OwnerNode and State.',$line
		}
	Foreach ($cluster in $currentclusterstatus)
		{
		If ($cluster.group -eq "Available-Storage" -or $cluster.group -eq "Cluster-Group")
			{
			$group=$cluster.group
			$group=$group -replace 'ble-Sto','ble Sto'
			$group=$group -replace 'ter-Gr','ter Gr'
			$status=$cluster.status
			$ndx=[array]::IndexOf($clustergroup,$group)
			If ($cluster.node -ne $clusternode[$ndx])
				{
				$prenode=$clusternode[$ndx]
				Write-Output "Cluster group ""$group"" is not mounted on pre-patch OwnerNode. Moving ""$group"" now..."
				$output=Invoke-expression "cluster group $group /move:$prenode"
				}else{
				Write-Output "Cluster group ""$group"" is mounted on pre-patch OwnerNode."
				$line
				}
			$prestatus=$clusterstatus[$ndx]
			If ($cluster.status -ne $clusterstatus[$ndx])
				{
				Write-Output "Cluster group ""$group"" is ""$status"" - pre-patch it was ""$prestatus"". Changing state to ""$status""..."
				$output=Invoke-expression "cluster group $group /$prestatus"
				}else{
				Write-Output "Cluster group ""$group"" is ""$status"" - pre-patch it was ""$prestatus"". No change needed."
				$line
				}
			}
			else
			{
			continue
			}
		}
	}
}
If ($Postpatch -eq $true)
	{
	$Rundate=get-date -Format MM-dd-yy
	$Installpath="C:\Avanade\Install Reports\Install Report $Rundate.html"
	$tempcsvpath="C:\Avanade\Install Reports\tempinstallcsv.csv"
	If ((Test-Path $tempcsvpath) -eq $true)
	{
	$allinstalls=Import-Csv $tempcsvpath
		
	$Output="<html>
<body>
<font size=""1"" face=""Arial,sans-serif"">
<h3 align=""center"">Installed KB Report</h3>
<h5 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>"

$Output+="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
<col width=""50%""><col width=""20%""><col width=""30%"">
<tr align=""center"" bgcolor=""#FF8000 ""><th>Hotfix Title</th><th>Hotfix ID (KB number)</th><th>KB Article</th></tr>
"
$Output+="</tr>"
	$AlternateRow=0;
	Write-Output "The Following Hotfixes were applied -"
	foreach ($Hotfix in $allinstalls)
	{
	[string]$KB=$Hotfix.kb
	$title=$Hotfix.Title
	If ($KB)
		{
		[string]$KBArticle="http://support2.microsoft.com/?kbid="+$KB.Substring(2)
		}else{
		$KBArticle="-"
		$KB="-"
		}
		$Output+="<tr"
		if ($AlternateRow)
		{
			$Output+=" style=""background-color:#dddddd"""
			$AlternateRow=0
		} else
		{
			$AlternateRow=1
		}
		$output+="><td>$($title)</td>"
		$Output+="<td>$($kb)</td>"
		$Output+="<td>$($KBArticle)</td>"
		$Output+="</tr>";
	Write-Output "$title"
	}
	$Output+="</table><br />"



$Output+="</body></html>"
$Output | Out-File $Installpath
Remove-Item $tempcsvpath
}else{
Write-Output "Hotfix install data is not present."
}
}