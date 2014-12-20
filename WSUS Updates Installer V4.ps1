#WSUS Updates Installer Script V4.1.2

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="low")]	
Param([Switch]$DownloadOnly,[Switch]$InstallOnly,[Switch]$Restart,[Switch]$Shutdown)

$ServerName=$env:ComputerName
$filename=("C:\Azaleos\{0}.txt" -f $env:ComputerName)
$line="--------------------------------------------------"
$Stars="******************************************"

function Log 
{
	param([string]$text)
	#Output to logfile
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
	#Output to screen
	Write-Output $text 
}
$starttime=(get-date)

$computer = "LocalHost" 
$namespace = "root\cimv2" 

$Removable = Get-WmiObject -class Win32_LogicalDisk -computername $computer -namespace $namespace
$RDrive = $Removable | Select Drivetype
If ($RDrive.Drivetype -eq 2)
	{
	Log ($Stars)
	Log ("*******Removable drive is detected********")
	Log ($Stars)
	}

#Region CheckRebootStatus if not Restart or Shutdown are selected
If($Restart -eq $True -or $Shutdown -eq $True)
{
	Try
	{
		$SystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"	
		If($SystemInfo.RebootRequired)
		{
     		Log ($line)
			Log ("{1} - {0} - WARNING - Reboot is required. Rebooting PC now." -f (Get-Date),$ServerName)
			Log ($line)
			Restart-Computer -Force
			Return
		}
	}
	Catch
	{      
		Log ($line)
	   	Log ("{1} - {0} - Can't check Reboot Status" -f (Get-Date),$ServerName)
	   	Log ($line)
	}
}
#endregion
		
#region	Connect to WSUS

#create session object
$ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
$Session = New-Object -ComObject "Microsoft.Update.Session" 
$Searcher = $Session.CreateUpdateSearcher()

$Service=$ServiceManager.Services|Where {$_.IsDefaultAUService -eq $true}

If($Service)
{
	$serviceName = $Service.Name
}
Else
{
	Log ("{1} - {0} - Can't Find Update Service" -f (Get-Date),$ServerName)
	Return
}
Log ($line)
Log ("{2} - {0} - Connecting to {1}, please wait..." -f (Get-Date),$serviceName,$ServerName)
Log ($line)
#endregion

#region Check for updates
Try
{
	$Results = $Searcher.Search("IsInstalled=0")
}
Catch
{
	If($_ -match "HRESULT: 0x80072EE2")
	{
		Log ($line)
		Log ("{1} - 0} - WARNING: Unable to connect to Windows Update server" -f (Get-Date),$ServerName)
		Log ($line)
	}
	Return
}
Log ($line)
Log ("{2} - {0} - Found {1} updates to download" -f (Get-Date),$Results.Updates.count,$ServerName)
Log ($line)
#endregion

#region Set Update Status/Accept Eula
$CollectedUpdates = New-Object -ComObject "Microsoft.Update.UpdateColl" 
Foreach($Update in $Results.Updates)
{	
	$size = [System.Math]::Round($Update.MaxDownloadSize/1MB,2)

	if($Update.KBArticleIDs -ne "")
	{
		$KB = "KB"+$Update.KBArticleIDs
	} 
	else
	{
		$KB = ""
	}

	$log = New-Object psobject
	$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
	$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
	$log | Add-Member -MemberType NoteProperty -Name Size -Value $size

    #accept eula and add to collection
    if($Update.EulaAccepted -eq 0)
	{
		$Update.AcceptEula()
	}
	$CollectedUpdates.Add($Update) | out-null   
	Log ("{3} -> {0} - {1} - {2} MB" -f $log.Title, $log.KB, $Log.Size, $ServerName)
}
if($CollectedUpdates.count)
{
    Log ($line)	
	Log ("{1} - {0} - Downloading Updates, please wait..." -f (Get-Date),$ServerName)
	Log ($line)
}
else
{
	Log ($line)
	Log ("{1} - {0} - There are 0 updates to install" -f (Get-Date),$ServerName)
	Log ($line)

	$endtime=(Get-Date)
	$duration=$endtime - $starttime

	Log ("Duration (minutes): {0:N2}" -f $duration.TotalMinutes)
    return
}
#endregion

#region Download Updates	
$Download = New-Object -ComObject "Microsoft.Update.UpdateColl"
$DownloadSummary = New-Object -ComObject "Microsoft.Update.UpdateColl"
foreach($Update in $CollectedUpdates)
{ 
	$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
	$objCollectionTmp.Add($Update) | Out-Null

	$Downloader = $Session.CreateUpdateDownloader() 
	$Downloader.Updates = $objCollectionTmp
	Try
	{
		$DownloadResult = $Downloader.Download()
	}
	Catch
	{
        if($_ -match "HRESULT: 0x80240044")
		{
			Log ($line)
			Log ("{1} - {0} - WARNING: Your security policy doesn't allow a non-administator to perform this task" -f (Get-Date),$ServerName)
			Log ($line)
		}
		Return
		If($_ -match "HRESULT: 0x80072F8F")
		{
			Log ($line)
			Log ("{1} - {0} - WARNING: 80072F8F Error occured. May be a Cert issue." -f (Get-Date),$ServerName)
			Log ($line)
		}
		return
	} 
	Switch -exact ($DownloadResult.ResultCode)
	{
		0   { $Status = "NotStarted" }
		1   { $Status = "InProgress" }
		2   { $Status = "Downloaded" }
		3   { $Status = "DownloadedWithErrors" }
		4   { $Status = "Failed" }
		5   { $Status = "Aborted" }
	}
    $log = New-Object psobject
	if($Update.KBArticleIDs -ne "")
	{
		$KB = "KB"+$Update.KBArticleIDs
	}
	else
	{
		$KB = ""
	}

	If($DownloadResult.ResultCode -ne 2)
	{
		$Status = "Failed"
	}

	$size = [System.Math]::Round($Update.MaxDownloadSize/1MB,2)

	$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
	$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
	$log | Add-Member -MemberType NoteProperty -Name Size -Value $size
	$log | Add-Member -MemberType NoteProperty -Name Status -Value $Status

	#output Results
	If($DownloadResult.ResultCode -eq 2)
	{
		$Download.Add($Update) | Out-Null
	}
	Else
	{
		Log ("{4} - **Failed to Download** - {3} - {0} - {1} - {2} MB" -f $log.Title, $log.KB, $Log.Size, $log.Status, $ServerName)
	}
}
#endregion

If($Download.Count -le $CollectedUpdates.Count)
{
	$DownloadSummary = $CollectedUpdates.count - $Download.Count
}
If($DownloadSummary -eq 0)
{
	Log ("All updates downloaded successfully.")
}
If($DownloadSummary -gt 0)
{
	Log ("** {0} Updates Failed to Download **" -f $DownloadSummary)
}
#Region Check Reboot Status for no selection and Download Only
If($Restart -eq $False -and $Shutdown -eq $False)
{
	Try
	{
		$SystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"	
		If($SystemInfo.RebootRequired)
		{
     		Log ($line)
			Log ("{0} - WARNING - Reboot is required. Log in and reboot it." -f (Get-Date))
			Log ($line)
			Return
		}
	}
	Catch
	{      
		Log ($line)
    	Log ("{0} - Can't check Reboot Status" -f (Get-Date))
    	Log ($line)
		Return
	}
}
#EndRegion

#region Install Updates

If($Restart -eq $True -or $Shutdown -eq $True -or $InstallOnly -eq $True)
{
	#Send message to logged in users
	msg * /time:600 "Avanade Managed Services is installing updates and rebooting this server. Please save your work immediately and log off. If you have any questions, contact us at 1-866-475-5557 in regards to case $az_casenumber and ask for" $az_caseowner.FullName

	$NeedsReboot = $false
	if($Download.count)
	{
		Log ($line)
		Log ("{2} - {1} - Installing {0} updates, please wait..." -f $Download.count, (Get-Date), $ServerName)
		Log ($line)
    }
    else
	{
		#No Updates to install - Exit Script
		Log ($line)
		Log ("{0} - There are 0 updates to install" -f (Get-Date))
		Log ($line)
		return
	}
	
	$FailedInstalls = New-Object -ComObject "Microsoft.Update.UpdateColl"

	Foreach($Update in $Download)
	{
		$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
		$objCollectionTmp.Add($Update) | Out-Null

		$objInstaller = $Session.CreateUpdateInstaller()
		$objInstaller.Updates = $objCollectionTmp
		
		$InstallStartTime=(Get-Date)
		
		Try{$InstallResult = $objInstaller.Install()}
		Catch
		{
			If($_ -match "HRESULT: 0x80240044")
			{
			Log ($line) 
			Log ("{1} - {0} - WARNING: Your security policy doesn't allow a non-administator to perform this task" -f (Get-Date), $ServerName)
			Log ($line)
			}
			Return
		}
		Switch -exact ($InstallResult.ResultCode)
		{
			0   { $Status = "NotStarted"}
			1   { $Status = "InProgress"}
			2   { $Status = "Installed"}
			3   { $Status = "InstalledWithErrors"}
			4   { $Status = "Failed"}
			5   { $Status = "Aborted"}
		}
		If(!$NeedsReboot) 
		{ 
			$NeedsReboot = $InstallResult.RebootRequired 
		}
		
    	$log = New-Object psobject
		if($Update.KBArticleIDs -ne "")
		{
			$KB = "KB"+$Update.KBArticleIDs} else {$KB = ""
		}
   		
		$size = [System.Math]::Round($Update.MaxDownloadSize/1MB,2)

		If($InstallResult.ResultCode -ne 2)
		{
			$Status = "Failed"
		}
		$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
		$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
		$log | Add-Member -MemberType NoteProperty -Name Size -Value $size
		$log | Add-Member -MemberType NoteProperty -Name Status -Value $Status
		
		$installendtime=(Get-Date)
		$installtime = ($installendtime - $installstarttime)
		
		#output Results
		Log ("{7} -> {3} - Duration: {4:00}:{5:00}.{6:N0} Minutes - {0} - {1} - {2} MB" -f $log.Title, $log.KB, $Log.Size, $log.Status, $installtime.Minutes, $installtime.Seconds, $installtime.milliseconds, $ServerName)
		If($InstallResult.ResultCode -ne 2)
		{
			$FailedInstalls.add($Update) | Out-Null
		}
	}
#endregion

#region Validation
	Log ($line)
	Log ("{1} - {0} - Validating Installation, please wait..." -f (Get-Date), $ServerName)
	Log ($line)

	If($FailedInstalls)
	{
		Foreach($Update in $FailedInstalls)
		{
			If($FailedInstalls.ResultCode -ne 2)
			{
				$Status = "Failed"
			}
			Else
			{
				$Status = "Installed"
			}
			If($FailedInstalls.ResultCode -ne 2)
			{
				$log = New-Object psobject
				if($Update.KBArticleIDs -ne "")    {$KB = "KB"+$Update.KBArticleIDs} else {$KB = ""}
				$log | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
				$log | Add-Member -MemberType NoteProperty -Name KB -Value $KB
				$log | Add-Member -MemberType NoteProperty -Name Status -Value $Status
				
				Log ("{3} **Failed to install** - {2} - {0} - {1}" -f $log.Title, $log.KB, $Log.Status, $ServerName)
			}
		}
	}
#endregion

#region Summary of Installed and Failed Patches

	Log ($line)
	Log ("{1} - {0} - Summary" -f (Get-Date), $ServerName)
	Log ($line)
	
	If($DownloadSummary -eq 0)
	{
		Log ("{0} - All updates downloaded successfully." -f $ServerName)
	}
	If($DownloadSummary -gt 0)
	{
		Log ("{1} ** {0} Updates Failed to Download **" -f $DownloadSummary, $ServerName)
	}
	If($FailedInstalls.Count)
	{
		Log ("{1} ** {0} Updates Failed to Install **" -f $FailedInstalls.Count, $ServerName)
	}
	Else
	{
		Log ("{0} - All Updates Installed Successfully" -f $ServerName)
	}
	
	#Calculate Duration of the scripted install
	$endtime=(Get-Date)
	$duration=$endtime - $starttime	
	Log ("Duration (minutes): {0:N2}" -f $duration.TotalMinutes)
	#endregion

#region Reboot or Shutdown server
	if($needsReboot -eq $false)
	{
		Log ($line)
		Log ("{1} - {0} - Restart NOT required" -f (Get-Date), $ServerName)
		Log ($line)
	}
	if($needsReboot -eq $true)
	{
		Log ($line)
		Log ("{1} - {0} - Restart required" -f (Get-Date), $ServerName)
		Log ($line)
	}
	If($Restart -eq $True -or $Shutdown -eq $True)
	{

		if($Restart -eq $true -and $needsReboot -eq $true)
		{
			Log ($line)
			Log ("{1} - {0} - Restarting computer" -f (Get-Date), $ServerName)
			Log ($line)
			#Restart Server
			Restart-Computer -force
		}
		if($Shutdown)
		{
			Log ($line)
			Log ("{1} - {0} - Shutting down computer" -f (Get-Date), $ServerName)
			Log ($line)
			#Shutdown server
			Stop-Computer -Force 
		}
	}
}