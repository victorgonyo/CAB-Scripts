Param([Switch]$Start,[Switch]$Stop,[Switch]$CheckStatus,[String]$TargetFQDN)
get-pssnapin -registered|add-pssnapin

If ($TargetFQDN -eq $null)
{
Write-Output "Please input a target FQDN Server"
exit
}

$test = get-exchangeserver $TargetFQDN -erroraction SilentlyContinue

If ($test -eq $null)
{
Write-Output "Target FQDN server indicated does not exist"
exit
}

If ($Start -eq $true)
{
$ServerName = $env:COMPUTERNAME 
Set-ServerComponentState $ServerName -Component HubTransport -State Draining -Requester Maintenance 
Restart-Service MSExchangeTransport 
Restart-Service MSExchangeFrontEndTransport 
Set-ServerComponentState $ServerName -Component UMCallRouter -State Draining -Requester Maintenance 
Redirect-Message -Server $ServerName -Target $TargetFQDN 
Suspend-ClusterNode $ServerName 
Set-MailboxServer $ServerName -DatabaseCopyActivationDisabledAndMoveNow $True 
Set-MailboxServer $ServerName -DatabaseCopyAutoActivationPolicy Blocked 
Set-ServerComponentState $ServerName -Component ServerWideOffline -State Inactive -Requester Maintenance 
}

If ($Stop -eq $true)
{
$ServerName = $env:COMPUTERNAME 
Set-ServerComponentState $ServerName -Component ServerWideOffline -State Active -Requester Maintenance 
Set-ServerComponentState $ServerName -Component UMCallRouter -State Active -Requester Maintenance 
Resume-ClusterNode $ServerName 
Set-MailboxServer $ServerName -DatabaseCopyActivationDisabledAndMoveNow $False 
Set-MailboxServer $ServerName -DatabaseCopyAutoActivationPolicy Unrestricted 
Set-ServerComponentState $ServerName -Component HubTransport -State Active -Requester Maintenance 
Restart-Service MSExchangeTransport 
Restart-Service MSExchangeFrontEndTransport 
}

If ($CheckStatus -eq $true)
{
$ServerName = $env:COMPUTERNAME 
Get-ServerComponentState $ServerName | ft Component,State -Autosize Get-MailboxServer $ServerName | ft DatabaseCopy* -Autosize 
Get-ClusterNode $ServerName | fl 
Get-Queue
}