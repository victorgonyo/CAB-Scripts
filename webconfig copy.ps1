$RunDate = (Get-Date).tostring("MM-dd-yyyy")
$RunTime = (Get-Date).ToShortTimeString()
$RunTime = $RunTime -replace ":","_"
$RunTime = $RunTime -replace " ","_"
$Avanade = "C:\Avanade"
$CAB = "C:\Avanade\CAB"
$FileName = ("C:\Avanade\CAB\owa-webconfig-" + $RunDate + "_" + $RunTime + "_.txt")
$FileName2 = ("C:\Avanade\CAB\exchweb-ews-webconfig-" + $RunDate + "_" + $RunTime + "_.txt")
If ((Test-Path $Avanade) -eq 0){
New-Item -ItemType directory -Path $Avanade | Out-Null
New-Item -ItemType directory -Path $CAB | Out-Null
}
If ((Test-Path $CAB) -eq 0){
New-Item -ItemType directory -Path $CAB | Out-Null
}
If (Test-Path $FileName)
{
Write-Output "
$FileName already exists. Please wait a minute so no files are overwritten."
}
else
{
$Server = $env:COMPUTERNAME
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
$ExchVer = (Get-ExchangeServer $Server).AdminDisplayVersion
$regKey= $reg.OpenSubKey($REG_ExSetup)
$MsiInstallPathValue = "MsiInstallPath"
$installPath = ($regkey.getvalue($MsiInstallPathValue))

$CopyPath1 = $installPath + "ClientAccess\Owa\web.config"
$CopyPath2 = $installPath + "ClientAccess\exchweb\ews\web.config"

Copy-Item $CopyPath1 "$FileName"
Copy-Item $CopyPath2 "$FileName2"
Write-Output "
$FileName has been created
$FileName2 has been created"
}