get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue
Write-Output "
================================================================
Checking for completed setup operations in Exchange setup log...
================================================================
"
$LogFolder = "C:\ExchangeSetupLogs\ExchangeSetup.log"
$search = "The Microsoft Exchange Server setup operation "
$LogFile = Get-Content $LogFolder | Select-String -AllMatches -SimpleMatch -Pattern $search
$LogFile|Select-Object -Last 5|%{$_.Line}|out-string