Param([Switch]$PreinstallCheck,[Switch]$AntiSpamCheck,[Switch]$SSLHTTPPre,[Switch]$SSLHTTPPost,[Switch]$BackupWebConfig,[Switch]$InterimUpdateCheck,[Switch]$AVTransportAgentCheck)
get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue

#region begining checks
$c=1
$Avanade = "C:\Avanade"
$CAB = "C:\Avanade\CAB"
If ((Test-Path $Avanade) -eq 0)
	{
	New-Item -ItemType directory -Path $Avanade | Out-Null
	New-Item -ItemType directory -Path $CAB | Out-Null
	}
If ((Test-Path $CAB) -eq 0)
	{
	New-Item -ItemType directory -Path $CAB | Out-Null
	}
$Server = $env:COMPUTERNAME
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Server)
$ExchVer = (get-exchangeserver $Server -ea silentlycontinue).AdminDisplayVersion
If ($ExchVer -eq $null)
{
$ExchVer = "Version 8"
}
$Programs = (Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall | % {Get-ItemProperty $_.PsPath} | where {$_.Displayname -and ($_.Displayname -match ".*")} | sort Displayname | select DisplayName, DisplayVersion, Publisher)
$Exchange = $Programs | where {$_.DisplayName -like "*exchange*"}
If ($Exchange -eq $null)
	{
	Write-Output "Exchange is not installed on this machine"
	exit
	}

#Creates values for the Registry that I want to pull eventually

$DisplayNameValue = "DisplayName"
$InstalledValue = "Installed"
$MsiInstallPathValue = "MsiInstallPath"
$LocalPackageValue = "LocalPackage"
$AllPatchesValue = "AllPatches"
$InstallPropertiesValue = "InstallProperties"

#Check to see if Exchange 2007, 2010, or 2013 is installed.

if ($ExchVer -match "Version 15")
	{
	$REG_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
	$Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v15\\Setup"
	$Reg_Patches = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Patches"
	$Reg_Ex = "SOFTWARE\\Microsoft\\ExchangeServer\\v15\\"
	$ServerName = (Get-ExchangeServer $server -status).Name
	$ServerRole = [string] (Get-ExchangeServer $server -status).ServerRole
	}
elseif ($ExchVer -match "Version 14")
	{
	$REG_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
	$Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v14\\Setup"
	$Reg_Patches = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Patches"
	$Reg_Ex = "SOFTWARE\\Microsoft\\ExchangeServer\\v14\\"
	$Reg_Anti = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products"
	$reg_msi = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\InstallProperties"
	$ServerName = (Get-ExchangeServer $server -status).Name
	$ServerRole = [string] (Get-ExchangeServer $server -status).ServerRole
	}
elseif	($ExchVer -match "Version 8")
	{
	$REG_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\461C2B4266EDEF444B864AD6D9E5B613\Patches"
	$Reg_ExSetup = "SOFTWARE\Microsoft\Exchange\Setup"
	$Reg_Patches = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches"
	$Reg_Ex = "SOFTWARE\Microsoft\Exchange\v8.0"
	$Reg_Anti = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
	$reg_msi = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\461C2B4266EDEF444B864AD6D9E5B613\InstallProperties"
	}
else
	{
	Write-Output "Exchange is not installed on this machine"
	exit
	}
If ($PreinstallCheck -eq $true -and $ExchVer -match "Version 8")
	{
	$2k7Servers = Get-ExchangeServer -ea silentlycontinue | ?{$_.AdminDisplayVersion -match "Version 8" -and $_.$_.ServerRole -ne "Edge"} | sort Name | select -ExpandProperty Name
	If ($2k7Servers -contains $Server)
		{
		$ServerName = (Get-ExchangeServer $server -status).Name
		$ServerRole = [string] (Get-ExchangeServer $server -status).ServerRole
		$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ServerName)
		}

If ($2k7Servers -notcontains $Server)
	{
	$ServerName = $env:COMPUTERNAME
	$ServerRole = "Mailbox"
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ServerName)
	}
}
#endregion

#region Preinstall Check
If ($PreinstallCheck -eq $true){
#Checks to see what roles are installed, and checks to see if there is a watermark
$regkey1 = $reg.OpenSubKey($Reg_Ex)
If($ServerRole.Contains("ClientAccess"))
	{
	$CASVersionCON = $regkey1.OpenSubKey("ClientAccessRole").GetValue("ConfiguredVersion")
	$CASVersionUNP = $regkey1.OpenSubKey("ClientAccessRole").GetValue("UnpackedVersion")
	$WMCA = $regkey1.OpenSubKey("ClientAccessRole").GetValue("WaterMark")
	}

If($ServerRole.Contains("Hub"))
	{
	$HUBVersionCON = $regkey1.OpenSubKey("HubTransportRole").GetValue("ConfiguredVersion")
	$HUBVersionUNP = $regkey1.OpenSubKey("HubTransportRole").GetValue("UnpackedVersion")
	$WMHT = $regkey1.OpenSubKey("HubTransportRole").GetValue("WaterMark")
	}

If($ServerRole.Contains("Mailbox"))
	{
	$MAILVersionCON = $regkey1.OpenSubKey("MailboxRole").GetValue("ConfiguredVersion")
	$MAILVersionUNP = $regkey1.OpenSubKey("MailboxRole").GetValue("UnpackedVersion")
	$WMMB = $regkey1.OpenSubKey("MailboxRole").GetValue("WaterMark")
	If ($ExchVer -eq "Version 8")
		{
		$WMCMS = $regkey1.OpenSubKey("ClusteredMailboxServer").GetValue("Watermark")
		}
	}

#Writes all the info out all pretty

Write-output "
=========================
Checking Installed Roles:
=========================

Server Name: $ServerName

Installed Roles: $ServerRole
"
If($ServerRole.Contains("ClientAccess"))
	{
    Write-output "CAS Configuered Version: $CASVersionCON","CAS Unpacked Version: $CASVersionUNP",""
	}
If($ServerRole.Contains("Hub"))
	{
	Write-Output "Hub Configuered Version: $HUBVersionCON","Hub Unpacked Version: $HUBVersionUNP",""
	}
If($ServerRole.Contains("Mailbox"))
	{
	Write-Output "Mailbox Configuered Version: $MAILVersionCON","Mailbox Unpacked Version: $MAILVersionUNP",""
	}

#If any watermarks are present, they will be displayed with the information located in the key.

If ($WMCA -ne $null)
	{
	Write-output "Client Access WaterMark is present $WMCA"
	$c=2
	}

If ($WMHT -ne $null)
	{
	Write-Output "Hub Transport WaterMark is present $WMHT"
	$c=2
	}

If ($WMMB -ne $null)
	{
	Write-Output "Mailbox WaterMark is present $WMMB"
	$c=2
	}
	
If ($WMCMS -ne $null)
	{
	Write-Output "Clustered Mailbox Server WaterMark is present $WMCMS"
	$c=2
	}
If ($c -eq 1)
	{
	Write-Output "No WaterMark is present."
	}

$regKey= $reg.OpenSubKey($REG_ExSetup)
$installPath = ($regkey.getvalue($MsiInstallPathValue) | foreach {$_ -replace (":","`$")})
$binFile = "Bin\ExSetup.exe"
$exSetupVer = ((Get-Command "\\$Server\$installPath$binFile").FileVersionInfo | ForEach {$_.FileVersion})
$regKey = $reg.OpenSubKey($REG_KEY).GetSubKeyNames() | ForEach {"$Reg_Key\\$_"}

#these pull the value displayname and installed from registry (no longer using as it pulls them in a non-logical order)

#$dispName = [array] ($regkey | %{$reg.OpenSubKey($_).getvalue($DisplayNameValue)})
#$instDate = [array] ($regkey | %{$reg.OpenSubKey($_).getvalue($InstalledValue)})
   
#this pulls the value of all patches so that it will drill down to the proper folders in the patches registry path

$regkey2 = $reg.OpenSubKey($REG_KEY)
$AllPatches = ($regkey2.GetValue($AllPatchesValue))

$regkey2 = $reg.OpenSubKey($REG_Patches)
$regkey1 = $reg.OpenSubKey($REG_KEY)
$regkey3 = $reg.OpenSubKey($reg_msi)

$MSILocalPackage = ($regkey3.GetValue($LocalPackageValue))
[string]$GUID = ($regkey3.GetValue('UninstallString'))
$GUID=[regex]::match($GUID,"\{([^\}]+)\}").groups[1].value

If (Test-Path $MSILocalPackage)
	{
	[string]$test = "Service Pack MSI file exists."
	}
	else
	{
	[string]$test = "***Service Pack MSI file DOES NOT exist - $MSILocalPackage***","GUID:$GUID"
	}
	
Write-Output "
Exchange Setup Version: $exSetupVer

$test

===========================
Checking Installed Rollups:
===========================
"

$countmembers = 0

#this is a loop that displays the RU's installed and their msp file locations and checks that the path exists
if ($regkey -ne $null)
    {
    while ($countmembers -lt $AllPatches.Count)
    {
    $AP = $AllPatches[$countmembers]
    $dispName = $regkey1.OpenSubKey($AP).GetValue($DisplayNameValue)
    $instDate = $regkey1.OpenSubKey($AP).GetValue($InstalledValue)
    $pinstDate = ($instDate.substring(0,4)+"/"+$instDate.substring(4,2)+"/"+$instDate.substring(6,2))
    Write-Output "$dispName","Installed on: $pinstDate"	
    $MSP = $regkey2.OpenSubKey($AP).GetValue($LocalPackageValue)
    Write-Output "$MSP"
    If (Test-Path $MSP) {
        Write-Output "MSP file exists
        "
        }else{
        Write-Output "**MSP file does not exist for this RU**
        "
        }
        $countmembers++
        }
    }
    else
    {
        Write-output "No Rollup Updates are installed.
        "
    }

}
#endregion

#region Anti Spam Check
If ($AntiSpamCheck -eq $true)
	{
	Write-Output "============================================","Checking 2007 Anti-spam and Block MSI files:","============================================",""
	If ($ServerRole.Contains("Hub"))
		{
		$regkey = @($reg.OpenSubKey($Reg_Anti).GetSubKeyNames() | ForEach {"$Reg_Anti\\$_"})
		ForEach ($r in $regkey)
			{
			$regkey1 = $reg.OpenSubKey($r)
			$dispName = [string] $regkey1.OpenSubKey("InstallProperties").GetValue($DisplayNameValue)
			$MSI = $regkey1.OpenSubKey("InstallProperties").GetValue($LocalPackageValue)
			If ($dispName.contains("Microsoft Exchange 2007"))
				{
				Write-Output "$dispName","$MSI"
				If (Test-Path $MSI) {
					Write-Output "MSI file exists
					"
					}else{
					Write-Output "**MSI file does not exist**
					"
					}
				}
			}
		}else{
		Write-Output "Hub Transport Role is not installed on this server"
		}
	}
#endregion

If ($SSLHTTPPre -eq $true -and $SSLHTTPPost -eq $true)
	{
	Write-Output "Please select only one SSL Offloading/Http Redirect flag at a time"
	$SSLHTTPPost=$false
	$SSLHTTPPre=$false
	}
#region SSL Offloading and HTTP Redirect PreUpdate
If ($SSLHTTPPre -eq $true)
    {
[array]$SSLEnabled=@()
[array]$HTTPRedirWebsite=@()
[array]$SSLDisabled=@()
[array]$All=@()
[array]$SSLValue=@()
$SSLHTTPFilename="C:\Avanade\CAB\SSLandHTTPsettings.csv"
Import-Module webadministration
$Server = $env:COMPUTERNAME
$OA = (Get-OutlookAnywhere -Server $server).ssloffloading
$OWA = (Get-ItemProperty -path HKLM:\\SYSTEM\CurrentControlSet\Services\"MSExchange OWA")
$DW = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site" -filter "system.webServer/security/access" -name "sslFlags"
$AllApps=Get-Childitem 'IIS:\Sites\Default Web Site' | ?{$_.Schema.Name -eq 'application'} | select -ExpandProperty pschildname
$AllVDs=Get-Childitem 'IIS:\Sites\Default Web Site' | ?{$_.Schema.Name -eq 'VirtualDirectory'} | select -ExpandProperty pschildname
Foreach ($App in $AllApps)
	{
	$SSL=Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "system.webServer/security/access" -name "sslFlags"
	If ($SSL -eq "Ssl,Ssl128")
		{
		continue
		}
	If ($SSL -eq "SslNegotiateCert")
		{
		$SSLDisabled+=$App
		$All+=$App
		$SSLValue += $false
		continue
		}
	If ($SSL -cmatch "Ssl")
		{
		$SSLEnabled+=$App
		$All+=$App
		$SSLValue += $true
		}
		else
		{
		$SSLDisabled+=$App
		$All+=$App
		$SSLValue += $false
		}
	}
If ($DW -cmatch "Ssl")
	{
	$SSLEnabled+="Default Web Site"
	$All+="Default Web Site"
	$SSLValue += $true
	}
	else
	{
	$SSLDisabled+="Default Web Site"
	$All+="Default Web Site"
	$SSLValue += $false
	}
Foreach ($VD in $AllVDs)
	{
	$SSL=Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$VD" -filter "system.webServer/security/access" -name "sslFlags"
	If ($SSL -cmatch "Ssl")
		{
		$SSLEnabled+=$VD
		$All+=$VD
		$SSLValue += $true
		}
		else
		{
		$SSLDisabled+=$VD
		$All+=$VD
		$SSLValue += $false
		}
	}
If ($OA -ne $true)
{
$SSLEnabled+="Outlook Anywhere"
$All+="Outlook Anywhere"
$SSLValue+=$true
}else{
$SSLDisabled+="Outlook Anywhere"
$All+="Outlook Anywhere"
$SSLValue+=$false
}

If ($OWA -contains "SSLOffloading")
{
$SSLEnabled+="MSExchangeOWA"
$All+="MSExchangeOWA"
$SSLValue+=$true
}else{
$SSLDisabled+="MSExchangeOWA"
$All+="MSExchangeOWA"
$SSLValue+=$false
}

Write-Output "================================","Checking SSL Offloading settings","================================"

If ($SSLEnabled)
	{
	Write-Output "=========================","The following are ENABLED","=========================",""
	$countmembers=0
	While ($countmembers -lt $SSLEnabled.count)
		{
		Write-Output $SSLEnabled[$countmembers]
		$countmembers++
		}
	}

####

If ($SSLDisabled)
	{
	Write-Output "","==========================","The following are DISABLED","==========================",""
	$countmembers=0
	While ($countmembers -lt $SSLDisabled.count)
		{
		Write-Output $SSLDisabled[$countmembers]
		$countmembers++
		}
	}
	$countmembers=0
	While ($countmembers -lt $All.count)
		{
		$HTTPRedirWebsite+=0
		$countmembers++
		}
	
	Import-Module webadministration
	Write-Output "","===============================","Checking HTTP Redirect settings","==============================="
	$y=0
	Foreach ($App in $All)
		{
		$HTTPRedir=(Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "/system.webServer/httpRedirect" -name "enabled").value
		$ndx = [array]::IndexOf($All, $App)
		If ($HTTPRedir -eq $true)
			{
			$y=1
			$HTTPRedirWeb=(Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "/system.webServer/httpRedirect" -name "destination").value
			[string]$Website="$App"+":"+"$HTTPRedirWeb"
			$HTTPRedirWebsite[$ndx]=$HTTPRedirWeb
			Write-Output $Website
			}
		}
	
	If ($y -eq 0)
		{
		Write-Output "No Redirect settings found"
		}
	$AllInfo = $All | ForEach-Object {
	$ndx = [array]::IndexOf($All, $_)
	New-Object PSObject -Property @{
	"Name" = "{0}" -f $_
	"SSLOffloading" = "{0}" -f $SSLValue[$ndx]
	"HTTPRedirect" = "{0}" -f $HTTPRedirWebsite[$ndx]
	}
	}
	$AllInfo | Export-Csv -Path $SSLHTTPFilename -NoTypeInformation
	}
#endregion

#region SSL Offloading and HTTP Redirect PostUpdate
If ($SSLHTTPPost -eq $true)
	{
	$SSLHTTPFilename="C:\Avanade\CAB\SSLandHTTPsettings.csv"
	If (Test-Path $SSLHTTPFilename)
	{
	[array]$Name = Import-Csv $filename | select -ExpandProperty Name
	[array]$HTTPRedirect = Import-Csv $filename | select -ExpandProperty HTTPRedirect
	[array]$SSLOffloading = Import-Csv $filename | select -ExpandProperty SSLOffloading
	[array]$SSLEnabled=@()
	[array]$HTTPRedirWebsite=@()
	[array]$SSLDisabled=@()
	[array]$All=@()
	[array]$SSLValue=@()
	Import-Module webadministration
	$Server = $env:COMPUTERNAME
	$OA = (Get-OutlookAnywhere -Server $server).ssloffloading
	$OWA = (Get-ItemProperty -path HKLM:\\SYSTEM\CurrentControlSet\Services\"MSExchange OWA")
	$DW = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site" -filter "system.webServer/security/access" -name "sslFlags"
	$AllApps=Get-Childitem 'IIS:\Sites\Default Web Site' | ?{$_.Schema.Name -eq 'application'} | select -ExpandProperty pschildname
	$AllVDs=Get-Childitem 'IIS:\Sites\Default Web Site' | ?{$_.Schema.Name -eq 'VirtualDirectory'} | select -ExpandProperty pschildname
	Foreach ($App in $AllApps)
		{
		$SSL=Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "system.webServer/security/access" -name "sslFlags"
		If ($SSL -eq "Ssl,Ssl128")
			{
			continue
			}
		If ($SSL -eq "SslNegotiateCert")
			{
			$SSLDisabled+=$App
			$All+=$App
			$SSLValue += $false
			continue
			}
		If ($SSL -cmatch "Ssl")
			{
			$SSLEnabled+=$App
			$All+=$App
			$SSLValue += $true
			}
			else
			{
			$SSLDisabled+=$App
			$All+=$App
			$SSLValue += $false
			}
		}
	If ($DW -cmatch "Ssl")
		{
		$SSLEnabled+="Default Web Site"
		$All+="Default Web Site"
		$SSLValue += $true
		}
		else
		{
		$SSLDisabled+="Default Web Site"
		$All+="Default Web Site"
		$SSLValue += $false
		}
	Foreach ($VD in $AllVDs)
		{
		$SSL=Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$VD" -filter "system.webServer/security/access" -name "sslFlags"
		If ($SSL -cmatch "Ssl")
			{
			$SSLEnabled+=$VD
			$All+=$VD
			$SSLValue += $true
			}
			else
			{
			$SSLDisabled+=$VD
			$All+=$VD
			$SSLValue += $false
			}
		}
	$countmembers=0
	While ($countmembers -lt $All.count)
		{
		
		}
	Foreach ($App in $All)
		{
		$HTTPRedir=(Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "/system.webServer/httpRedirect" -name "enabled").value
		$ndx = [array]::IndexOf($All, $App)
		If ($HTTPRedir -eq $true)
			{
			$HTTPRedirWeb=(Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/$App" -filter "/system.webServer/httpRedirect" -name "destination").value
			
			
			
			}
		}
	
	}else{
	Write-Output "Please run the SSLHTTPPre flag before running the Post flag"
	}
	
	}
#endregion

#region Backup Web Config
If ($BackupWebConfig -eq $true)
	{
	$RunDate = (Get-Date).tostring("MM-dd-yyyy")
	$RunTime = (Get-Date).ToShortTimeString()
	$RunTime = $RunTime -replace ":","_"
	$RunTime = $RunTime -replace " ","_"
	$FileName = ("C:\Avanade\CAB\owa-webconfig-" + $RunDate + "_" + $RunTime + "_.txt")
	$FileName2 = ("C:\Avanade\CAB\exchweb-ews-webconfig-" + $RunDate + "_" + $RunTime + "_.txt")
	If (Test-Path $FileName)
		{
		Write-Output "
		$FileName already exists. Please wait a minute so no files are overwritten."
		}else{
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
	}
#endregion

#region Interim Update Check
If ($InterimUpdateCheck -eq $true)
	{
	$Programs = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall | 
	% {Get-ItemProperty $_.PsPath} | where {$_.Displayname -and ($_.Displayname -match ".*")} | 
	sort Displayname | select DisplayName, DisplayVersion, Publisher
	$Interim = $Programs | where {$_.DisplayName -like "*interim*"}
	if ($Interim | where {$_.DisplayName -like "*interim*"})
		{
		Write-Output "","--------------------------------------"
		Write-Output "Exchange Interim Updates are installed"
		$Interim
		Write-Output "--------------------------------------"
		}else{
		Write-Output "","------------------------------------------"
		Write-Output "Exchange Interim Updates are NOT installed"
		Write-Output "------------------------------------------"
		}
	}
#endregion

#region AV Transport Agent Check
If ($AVTransportAgentCheck -eq $true)
	{
	$TransportAgents = Get-TransportAgent 
	$TransportAgentMc = $TransportAgents | where {$_.Identity -like "*McAfee*"}
	If ($TransportAgentMc | where {$_.Identity -like "*McAfee*"})
		{
		Write-Output "","--------------------------------------"
		Write-Output "McAfee Transport Agents are installed"
		$TransportAgentMc | select Identity, Enabled, Priority
		}else{
		Write-Output "","--------------------------------------"
		Write-Output "McAfee Agents are NOT installed"
		}
	$TransportAgentSy = $TransportAgents | where {$_.Identity -like "*SMS*"}
	if ($TransportAgentSY | where {$_.Identity -like "*SMS*"})
		{
		Write-Output "--------------------------------------"
		Write-Output "Symantec Transport Agents are installed"
		$TransportAgentSY | select Identity, Enabled, Priority
		Write-Output "--------------------------------------"
		}else{
		Write-Output "--------------------------------------"
		Write-Output "Symantec Transport Agents are NOT installed"
		Write-Output "--------------------------------------"
		}
	}
#endregion
