$b=[system.convert]::frombase64string("JGxpbmU9Ii0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tIg0KJFN0YXJzPSIqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioiDQpXcml0ZS1PdXRwdXQgIg0KQ0NNUyBQcm9mZXNzaW9uYWwgU2VydmljZXMgdG8gT3BlcmF0aW9ucyBIYW5kb2ZmIENoZWNrbGlzdCAtIEV4Y2hhbmdlIDIwMTMNCkxhc3QgVXBkYXRlZDogNS8zMC8yMDE0DQoNClNlcnZlciBDb25maWd1cmF0aW9uDQoNClN5c3RlbSBJbmZvcm1hdGlvbg0KJGxpbmUNCiINCg0KR2V0LUV4Y2hhbmdlU2VydmVy")
$a=[system.text.encoding]::utf8.getstring($b)
Out-File C:\Avanade\CAB\test.ps1 -InputObject $a
$computername = $env:COMPUTERNAME
$pssession=New-PSSession -ComputerName $computername
Enter-PSSession $pssession

$line="--------------------------"
$Stars="******************************************"
Write-Output "
CCMS Professional Services to Operations Handoff Checklist - Exchange 2013
Last Updated: 5/30/2014

Server Configuration

System Information
$line
"

Get-ExchangeServer