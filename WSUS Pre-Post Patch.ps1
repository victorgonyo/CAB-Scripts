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