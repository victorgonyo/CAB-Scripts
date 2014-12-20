Param([Switch]$StopTrustedInstaller,[String]$TargetServerName)
get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue

If ($TargetServerName -ne $null)
{
$RemoteTask = Get-WmiObject -class Win32_Service -computername $TargetServerName -ErrorAction SilentlyContinue | ?{$_.DisplayName -eq "Windows Modules Installer"}
}

If ($RemoteTask -eq $null)
{
Write-Output "The specified server $TargetServerName is not reachable."
exit
}

$RemoteTask = (Get-Service -ComputerName $TargetServerName -DisplayName "Windows Modules Installer").status

Write-Output "Trusted Installer is currently $RemoteTask"

If ($StopTrustedInstaller -eq $true -and $RemoteTask -ne "Stopped")
{
taskkill /s $TargetServerName /IM trustedinstaller.exe /F
}

##$RemoteTask1 = (Get-Service -ComputerName $TargetServerName -DisplayName "Windows Modules Installer").status

##If ($RemoteTask -eq $RemoteTask1)
##{
##sc \\$TargetServerName stop "Windows Module Installer"
##}