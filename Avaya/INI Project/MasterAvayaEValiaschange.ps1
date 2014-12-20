Param([string]$InputFilename,[switch]$PhaseOnly)
If ($PhaseOnly -eq $false)
	{
	If (!$InputFilename)
		{
		$InputFilename=Read-Host "Please give the location of the starting CSV file"
		}
	If ((Test-Path $InputFilename) -eq 0)
		{
		Write-Output "That file does not exist. Please re-run script and supply proper filename"
		exit
		}
	$check=Read-Host "This script is designed to find information based off the CSV file having the mailboxes old alias and new alias. This information is what is in the CSV? (Y/N)"
	If ($check -notmatch "Y" -or $check -notmatch "Yes")
		{
		Write-Output "Please supply a CSV file that has old alias and new alias mailbox information."
		exit
		}
	$alias=Import-Csv $InputFilename | select -expandproperty alias
	$newalias = Import-Csv $InputFilename | select -expandproperty newalias
	$AllInfo = $alias | ForEach-Object {
		$mailbox=get-mailbox $_
		$ndx = [array]::IndexOf($alias, $_)
		New-Object PSObject -Property @{
		"Alias" = "{0}" -f $_
		"NewAlias" = "{0}" -f $newalias[$ndx]
		"DisplayName" = "{0}" -f $mailbox.displayname
		"PrimarySMTPaddress" = "{0}" -f $mailbox.primarysmtpaddress
		"LegacyExchangeDN" = "{0}" -f $mailbox.legacyexchangedn
		"Database" = "{0}" -f $mailbox.database
		}
		}

	$OutputFilename = Read-Host "Please give the location where you would like to save the output file with all of the needed information"
	$AllInfo | Export-Csv -Path $OutputFilename -NoTypeInformation
	}
If (!$OutputFilename)
	{
	$OutputFilename = Read-Host "Please give the location of the CSV file with all of the required information (Alias, NewAlias, DisplayName, PrimarySMTPaddress, LegacyExchangeDN, and Database)"
	}
	Write-Output "Starting Process for enabling and disabling mailboxes."

	$AllInfo=Import-Csv $OutputFilename
	$databases=Import-Csv $OutputFilename | select -ExpandProperty Database | Select -Unique
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
	Set-mailbox $Info.newalias -PrimarySMTPAddress $Info.PrimarySMTPAddress -emailaddresspolicyenabled $false
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
$failed=$null
Foreach ($Info in $AllInfo)
	{
	[array]$newAlias+=$Info.newalias
	[array]$oldAlias+=$Info.alias
	$nAlias=$Info.newalias
	$oAlias=$Info.Alias
	[string]$oldandnew="$oAlias" + "\" + "nAlias"
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
		[array]$failed+=$oldandnew
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
$OutputFilename = Read-Host "Please give the location where you would like to save the results log"
$AllInfo | Export-Csv -Path $OutputFilename -NoTypeInformation
If (!$somefail)
	{
	Write-Output "Script successfully changed aliases on all mailboxes."
	}
	else
	{
	Write-Output "The following aliases failed to change (oldalias\newalias)",$failed
	}