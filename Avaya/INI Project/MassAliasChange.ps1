$AllInfo=Import-Csv C:\Azaleos\HeliosPhase2.csv
$databases=Import-Csv C:\Azaleos\HeliosPhase2.csv | select -ExpandProperty Database | Select -Unique
Write-Output "-------","Stage 1","-------"
Foreach ($Info in $AllInfo)
	{
	$Alias=$Info.alias
	Write-Output "Disabling Mailbox with alias $Alias."
	Disable-Mailbox $Info.alias -confirm:$false
	sleep 10
	}
Write-Output "-------","Stage 2","-------"
Foreach ($database in $databases)
	{
	Write-Output "Cleaning Database $database."
	clean-mailboxdatabase $database
	sleep 10
	}

Write-Output "-------","Stage 3","-------"
Foreach ($Info in $AllInfo)
	{
	$Alias=$Info.alias
	$NewAlias=$Info.newalias
	Write-Output "Connecting Mailbox that had old alias of $Alias and new alias of $NewAlias."
	Connect-Mailbox -identity $Info.DisplayName -User $Info.newalias -Alias $Info.newalias -Database $Info.database
	Sleep 10
	$PostSMTP=$Info.PostSMTP
	$PrimarySMTP=$Info.PrimarySMTPAddress
	$combinedSMTP="$PostSMTP" + "," + "$PrimarySMTP"
	set-mailbox $Info.newalias -PrimarySMTPAddress $PostSMTP -emailaddresspolicyenabled $false
	sleep 3
	set-mailbox $Info.newalias -PrimarySMTPAddress $PrimarySMTP -emailaddresspolicyenabled $false
	sleep 10
	}
Write-Output "-------","Stage 4","-------"
Foreach ($database in $databases)
	{
	Write-Output "Cleaning Database $database."
	clean-mailboxdatabase $database
	sleep 10
	}
Write-Output "-------","Stage 5","-------"
Write-Output "Confirming success of alias change..."
$newAlias=$null
$oldAlias=$null
$somefail=$null
Foreach ($Info in $AllInfo)
	{
	[array]$newAlias+=$Info.newalias
	[array]$oldAlias+=$Info.alias
	$nAlias=$Info.newalias
	$oAlias=$Info.Alias
	$check=Get-mailbox $Info.newalias -erroraction silentlycontinue
	If ($check)
		{
		[array]$Success+="Success"
		Write-Output "Change from old alias $oAlias to new alias $nAlias was successful!"
		}
		else
		{
		$somefail=1
		[array]$Success+="Fail"
		Write-Output "Change from old alias $oAlias to new alias $nAlias failed."
		}
	}
	$AllInfo = $newAlias | ForEach-Object {
	$ndx = [array]::IndexOf($newAlias, $_)
	New-Object PSObject -Property @{
	"NewAlias" = "{0}" -f $_
	"OldAlias" = "{0}" -f $oldAlias[$ndx]
	"Result" = "{0}" -f $Success[$ndx]
	}
	}
	$AllInfo | Export-Csv -Path C:\Azaleos\HeliosSuccessLog.csv -NoTypeInformation
If (!$somefail)
	{
	Write-Output " _______           _______  _______  _______  _______  _______  _
(  ____ \|\     /|(  ____ \(  ____ \(  ____ \(  ____ \(  ____ \( )
| (    \/| )   ( || (    \/| (    \/| (    \/| (    \/| (    \/| |
| (_____ | |   | || |      | |      | (__    | (_____ | (_____ | |
(_____  )| |   | || |      | |      |  __)   (_____  )(_____  )| |
      ) || |   | || |      | |      | (            ) |      ) |(_)
/\____) || (___) || (____/\| (____/\| (____/\/\____) |/\____) | _
\_______)(_______)(_______/(_______/(_______/\_______)\_______)(_)"
	}
	else
	{
	Write-Output "The following aliases failed to change (oldalias\newalias)",$failed
	}