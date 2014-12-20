$path = Read-host "Enter path to CSV"
$servers = Import-csv $path

$lstat = DATA {
ConvertFrom-StringData -StringData @'
0 = Unlicensed
1 = Licensed
2 = OOB Grace
3 = OOT Grace
4 = Non-Genuine Grace
5 = Notification
6 = Extended Grace
'@
}
$report = @()

Foreach($computer in $servers){
$test = @()

$test = Get-WmiObject SoftwareLicensingProduct -ComputerName $computer.name | where {$_.PartialProductKey} | select __SERVER, @{N="LicenseStatus"; E={$lstat["$($_.LicenseStatus)"]} },name
$report += $test
	
}

$report | ft -auto	