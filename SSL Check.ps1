get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue
Import-Module webadministration

$Server = $env:COMPUTERNAME
$OA = (Get-OutlookAnywhere -Server $server).ssloffloading
$OWA = (Get-ItemProperty -path HKLM:\\SYSTEM\CurrentControlSet\Services\"MSExchange OWA")
$ECP = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/ECP" -filter "system.webServer/security/access" -name "sslFlags"
$OAB = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/OAB" -filter "system.webServer/security/access" -name "sslFlags"
$EWS = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/EWS" -filter "system.webServer/security/access" -name "sslFlags"
$AS = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/Microsoft-Server-ActiveSync" -filter "system.webServer/security/access" -name "sslFlags"
$AD = Get-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST" -location "Default Web Site/Autodiscover" -filter "system.webServer/security/access" -name "sslFlags"

$ALL = @($OA, $OWA, $ECP, $OAB, $EWS, $AS, $AD)

ForEach ($x in $ALL){
    if ($x -cmatch "Ssl" -or $OA -ne $true -or $OWA -contains "SSLOffloading")
    {
        if ($y -eq $null){    
        Write-Output "==========================","The following are ENABLED","==========================",""
        $y=1
        }
    }
    }

If ($OA -ne $true)
{
Write-Output "Outlook Anywhere"
}

If ($OWA -contains "SSLOffloading")
{
Write-Output "Owa"
}

If ($ECP -eq "Ssl")
{
Write-Output "ECP"
}

If ($OAB -eq "Ssl")
{
Write-Output "OAB"
}

If ($EWS -eq "Ssl")
{
Write-Output "EWS"
}

If ($AS -eq "Ssl")
{
Write-Output "Active Sync"
}

If ($AD -eq "Ssl")
{
Write-Output "Autodiscover"
}

####

ForEach ($x in $ALL){
    if ($x.value -eq 0 -or $OA -eq $true -or $OWA -notcontains "SSLOffloading")
    {
        if ($z -eq $null){    
        Write-Output "","===========================","The following are DISABLED","===========================",""
        $z=1
        }
    }
    }


####

If ($OA -eq $true)
{
Write-Output "Outlook Anywhere"
}

If ($OWA -notcontains "SSLOffloading")
{
Write-Output "Owa"
}

If ($ECP -ne "Ssl")
{
Write-Output "ECP"
}

If ($OAB -ne "Ssl")
{
Write-Output "OAB"
}

If ($EWS -ne "Ssl")
{
Write-Output "EWS"
}

If ($AS -ne "Ssl")
{
Write-Output "Active Sync"
}

If ($AD -ne "Ssl")
{
Write-Output "Autodiscover"
}