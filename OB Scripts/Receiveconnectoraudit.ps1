#Grab All Receive Connectors
$AllConnectors=Get-ReceiveConnector
#Grab Transport server to use for comparisons
$BaseServer=Get-TransportServer |Select -First 1
Write-Output ("User {0} as the base object" -f $BaseServer.Name)
$BaseConnectors=$AllConnectors|Where {$_.Server -eq $BaseServer.Name -and ($_.Name -notmatch "Default " -and $_.Name -notmatch "Client ")}

foreach ($BaseConn in $BaseConnectors) {
	Write-Output ("Checking: {0}" -f $BaseConn.Name)
	#Base AD Permissions
	$BaseADRights=$baseconn|Get-ADPermission |where {$_.ExtendedRights -ne $null}|Select Identity, ExtendedRights|Sort ExtendedRights
	
	$CompareConns=$AllConnectors |Where {$_.Name -eq $BaseConn.Name -and $_.Server -ne $BaseConn.Server}
	if ($CompareConns) {
		foreach ($CompareConn in $CompareConns) {
			#Compare Connector
			$Result=Compare-Object -ReferenceObject ($BaseConn.RemoteIPRanges|Sort) -DifferenceObject ($CompareConn.REmoteIPRanges|Sort)
			if ($Result) {Write-Output ("Mismatch: {0} -- {1}" -f $BaseConn.Identity.ToString(), $CompareConn.Identity.ToString())}
			
		}
	}
}