[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="low")]
Param([Switch]$DownloadOnly,[Switch]$Restart,[Switch]$Shutdown)

$filename=("C:\Azaleos\{0}.txt" -f $env:ComputerName)
$line="--------------------------------------------------"

#Send message to logged in users
msg * /time:600 "This computer is being patched and will be restarted shortly. Please save your work immediately and logoff. $az_casenumber "  $az_caseowner.FullName

function Log {
	param([string]$text)
	#Output to logfile
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
	#Output to screen
	Write-Output $text 
}
$starttime=(get-date)
#Region CheckRebootStatus
Try
{
    $SystemInfo= New-Object -ComObject "Microsoft.Update.SystemInfo"
    if($SystemInfo.RebootRequired)
    {
        Log ($line)
		Log ("{0} - WARNING - Reboot is required to continue" -f (Get-Date))
		Log ($line)
		
        return
    }
}
Catch
{      
	Log ($line)
    Log ("{0} - Can't check Reboot Status" -f (Get-Date))
    Log ($line)

}
#endregion


#create session object
$ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
$Session = New-Object -ComObject "Microsoft.Update.Session"
$Searcher = $Session.CreateUpdateSearcher()

#check source of updates   
$service=$serviceManager.Services|Where {$_.IsDefaultAUService -eq $true}
if($Service){
	$serviceName = $Service.Name
}
else {
	Log ("{0} - Can't Find Update Service" -f (Get-Date))
	Return
} 
        
Log ($line)
Log ("{0} - Connecting to {1}, please wait..." -f (Get-Date),$serviceName)
Log ($line)

#check updates

Try
{        
    $Results = $Searcher.Search("IsInstalled=0")
}
Catch
{
    if($_ -match "HRESULT: 0x80072EE2")
    {
		Log ($line)
       	Log ("{0} - WARNING: Unable to connect to Windows Update server" -f (Get-Date))
	   	Log ($line)

    }
	$_
    Return
}
Log ($line)
Log ("{0} - Found {1} updates to download" -f (Get-Date),$Results.Updates.count)
Log ($line)

#set update status    
$Updates = New-Object -ComObject "Microsoft.Update.UpdateColl"
foreach($Update in $Results.Updates)
{
	$size = [System.Math]::Round($Update.MaxDownloadSize/1MB,2)
	if($Update.KBArticleIDs -ne "") {$KB = "KB"+$Update.KBArticleIDs} 
	else {$KB = ""}
    
	$log = New-Object psobject
	$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
	$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
	$log | Add-Member -MemberType NoteProperty -Name Size -Value $size
	        
    #accept eula and add to collection
    if ( $Update.EulaAccepted -eq 0 ) { $Update.AcceptEula() }
    $Updates.Add($Update) | out-null   
	Log ("{0} - {1} - {2} MB" -f $log.Title, $log.KB, $Log.Size)
}


if($Updates.count)
{
    Log ($line)	
	Log ("{0} - Downloading Updates, please wait..." -f (Get-Date))
	Log ($line)
}
else
{
	#No Updates to install - Exit Script
	Log ($line)
	Log ("{0} - There are 0 updates to install" -f (Get-Date))
	Log ($line)
	$endtime=(Get-Date)
	$duration=$endtime - $starttime
	Log ("Duration (minutes): {0:N2}" -f $duration.TotalMinutes)
    return
}

#Region DownloadUpdates    
$UpdatesInstall = New-Object -ComObject "Microsoft.Update.UpdateColl"
foreach($Update in $Updates)
{   
    $UpdatesTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
    $UpdatesTmp.Add($Update) | out-null
        
    $Downloader = $Session.CreateUpdateDownloader() 
    $Downloader.Updates = $UpdatesTmp
    try
    {
		
        $DownloadResult = $Downloader.Download()
    }
    Catch
    {
        if($_ -match "HRESULT: 0x80240044")
        {
            Log ($line)
			Log ("{0} - WARNING: Your security policy doesn't allow a non-administator to perform this task" -f (Get-Date))
			Log ($line)
        }
        return
    } 
    
    switch -exact ($DownloadResult.ResultCode)
    {
        0   { $Status = "NotStarted"}
        1   { $Status = "InProgress"}
        2   { $Status = "Downloaded"}
        3   { $Status = "DownloadedWithErrors"}
        4   { $Status = "Failed"}
        5   { $Status = "Aborted"}
    }
            
    $log = New-Object psobject
                    
    if($Update.KBArticleIDs -ne "")    {$KB = "KB"+$Update.KBArticleIDs} else {$KB = ""}
    $size = [System.Math]::Round($Update.MaxDownloadSize/1MB,2)
                    
	$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
	$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
	$log | Add-Member -MemberType NoteProperty -Name Size -Value $size
	$log | Add-Member -MemberType NoteProperty -Name Status -Value $Status
   
	
	#output Results
	Log ("{0} - {1} - {2} MB - {3}" -f $log.Title, $log.KB, $Log.Size, $log.Status )
          
    if($DownloadResult.ResultCode -eq 2)
    {
        $UpdatesInstall.Add($Update) | out-null
    }
}
#endregion 


#Region InstallUpdates
if(!$DownloadOnly)
{
    $needsReboot = $false
   
    if($UpdatesInstall.count)
    {
        Log ($line)
		Log ("{0} - Installing updates, please wait..." -f (Get-Date))
		Log ($line)
    }
    else{
		#No Updates to install - Exit Script
		Log ($line)
		Log ("{0} - There are 0 updates to install" -f (Get-Date))
		Log ($line)
		return
	}

	$Installer = $Session.CreateUpdateInstaller()
	$Installer.Updates = $UpdatesInstall
	          
	try
	{                
	    $InstallResult = $Installer.Install()
	}
	Catch
	{
	    if($_ -match "HRESULT: 0x80240044")
	    {
			Log ($line) 
			Log ("{0} - WARNING: Your security policy doesn't allow a non-administator to perform this task" -f (Get-Date))
			Log ($line)
	    }
	    return
	}
	        
	if(!$needsReboot) {$needsReboot = $installResult.RebootRequired}  

	#Calculate Duration of the scripted install
	$endtime=(Get-Date)
	$duration=$endtime - $starttime

	#endregion 
#region Summary
	Log ($line)
	Log ("{0} - Validating Installation, please wait..." -f (Get-Date))
	Log ($line)
	
	$LogCollection=@()
	for ($i=0; $i -lt $Updates.count ; $i++) 
	{
		#Set Status
		If ($InstallResult.GetUpdateResult($i).ResultCode -eq 2){$Status = "Installed"}
		Else{$Status = "Failed"}
	                    
	    #if($Update.KBArticleIDs -ne "")    {$KB = "KB"+$Update.KBArticleIDs} else {$KB = ""}
	    
		$log = New-Object psobject               
		$log | Add-Member -MemberType NoteProperty -Name Title -Value $Updates.item($i).Title
		$log | Add-Member -MemberType NoteProperty -Name Status -Value $Status
	   	
		#output Results
		Write-Output ("{0} - {1}" -f $log.Title, $log.Status )
		#Add to log collection
		$LogCollection+=$log
	}
		

	#List Summary of Patches - List any Failed Patches
	Log ($line)
	Log ("{0} - Summary" -f (Get-Date))
	Log ($line)
	Log ($LogCollection|group Status -NoElement|Out-string)

	$FailedUpdates=$LogCollection|Where {$_.Status -eq "Failed"}
	If ($FailedUpdates) {
		Log ("Failed Updates")
		Log ($line)
		foreach($log in $FailedUpdates){
			Log ("Title:`t{0}" -f $log.Title)
			Log ("KB:`t{0}" -f $log.KB)
		}
	}
	Log ("Duration (minutes): {0:N2}" -f $duration.TotalMinutes)
#endregion
	
	#Reboot or Shutdown server
	if($needsReboot -eq $false) {
		Log ($line)
		Log ("{0} - Restart NOT required" -f (Get-Date))
		Log ($line)
		return
	}
	if($Restart -eq $true -and $needsReboot -eq $true){
		Log ($line)
		Log ("{0} - Restarting computer" -f (Get-Date))
		Log ($line)
		#Restart Server
		Restart-Computer -force
	}
	if($Shutdown) {
		Log ($line)
		Log ("{0} - Shutting down computer" -f (Get-Date))
		Log ($line)
		#Shutdown server
		Stop-Computer -Force 
	}
}