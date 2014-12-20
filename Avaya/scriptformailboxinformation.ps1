$filepath="C:\Avanade\CSV Files"
$filename="sendasreportforglobal"
[int]$filenumber=0
[int]$check=0
If ((Test-Path $filepath) -eq 0)
{
New-Item -ItemType directory -Path 'C:\Avanade' | Out-Null
New-Item -ItemType directory -Path $filepath | Out-Null
}
If ((Test-Path ("$filepath\$filename.csv")) -eq $true)
	{
	$filenumber++
	While ($check -eq 0)
	{
	If ((Test-Path ("$filepath\$filename $filenumber.csv")) -eq $true)
		{
		$filenumber++
		}else{
		$check=1
		}
	}
	}
[array]$Allinfo=@()
[array]$Mailboxname=@()
[array]$allsmtpaddress=@()
[array]$mailboxinfo=Import-Csv "C:\Avanade\AvayaSendAsRequest.csv" | select -ExpandProperty Identity
$countmailboxes=0
$mailbox=$null
while ($countmailboxes -lt $mailboxinfo.count)
	{
	If ($mailbox -ne $mailboxinfo[$countmailboxes])
		{
		[string]$mailbox=$mailboxinfo[$countmailboxes]
		Write-Output "Gathering information for mailbox $mailbox..."
		$ADPermission=Get-Mailbox $mailbox | Get-ADpermission |? {$_.extendedrights -like "*send-as*"}|? {$_.user -notlike "NT AUTHORITY\SELF"}
		$Allinfo+=$ADPermission
		$mailboxsmtpaddress=Get-Mailbox $mailbox -ea silentlycontinue | select -ExpandProperty primarysmtpaddress
		}
	If ($Mailboxname.count -ne $Allinfo.count)
		{
		$allsmtpaddress+=$mailboxsmtpaddress
		$Mailboxname+=$mailbox
		}else{
		$countmailboxes++
		}
	}
Write-Output "Generating CSV file..."
$newpsobject=$Allinfo | ForEach-Object {
	$ndx=[array]::IndexOf($Allinfo,$_)
	[string]$tempuser=$_.user
	If ($tempuser -match 'global')
		{
		$tempuser = $tempuser -replace 'GLOBAL',''
		$tempuser = $tempuser.Substring(1,$tempuser.length-1)
		$smtpaddress=&get-mailbox $tempuser -ea silentlycontinue | select -expand primarysmtpaddress
		}else{
		$smtpaddress=""
		}
	$properties = @{"Identity" = $_.identity;"Primary SMTP Address" = $smtpaddress;"Inherited" = $_.isinherited;"Deny" = $_.deny;"User" = $_.user;"Mailbox name" = $mailboxname[$ndx];"Mailbox Primary SMTP Address" = $allsmtpaddress[$ndx]}
	New-Object -Type PSObject -Property $properties
	}
If ($filenumber -eq 0)
	{
	Write-Output "CSV File saved at $filepath\$filename.csv"
	$newpsobject | Export-Csv -path "$filepath\$filename.csv" -NoTypeInformation
	}else{
	Write-Output "CSV File saved at $filepath\$filename $filenumber.csv"
	$newpsobject | Export-Csv -path "$filepath\$filename $filenumber.csv" -NoTypeInformation
	}
