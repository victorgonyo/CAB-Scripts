[array]$alias=Import-Csv C:\Azaleos\AvayaHeliosTest2.csv | select Alias
[array]$newalias = Import-Csv C:\Azaleos\AvayaHeliosTest2.csv | select NewAlias
[array]$PostSMTP = Import-Csv C:\Azaleos\AvayaHeliosTest2.csv | select -expandproperty PostSMTP
$AllInfo = $alias | ForEach-Object {
	$things=$_
	$mailbox=get-mailbox $things.alias
	$ndx = [array]::IndexOf($alias, $_)
	If ($ndx -eq $null)
		{$ndx = 0}
	$new=$newalias[$ndx]
	New-Object PSObject -Property @{
	"alias" = "{0}" -f $things.alias
	"newalias" = "{0}" -f $new.newalias
	"PostSMTP" = "{0}" -f $PostSMTP[$ndx]
	"DisplayName" = "{0}" -f $mailbox.DisplayName
	"primarysmtpaddress" = "{0}" -f $mailbox.primarysmtpaddress
	"legacyexchangedn" = "{0}" -f $mailbox.legacyexchangedn
	"database" = "{0}" -f $mailbox.database
	}
	}
	$AllInfo | Export-Csv -Path 'C:\Azaleos\HeliosPhase2.csv' -NoTypeInformation
