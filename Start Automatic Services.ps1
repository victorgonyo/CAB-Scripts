get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue
$Ignore = @( 
		    'Microsoft .NET Framework NGEN v4.0.30319_X64', 
		    'Microsoft .NET Framework NGEN v4.0.30319_X86', 
		    'Performance Logs and Alerts',
            'Shell Hardware Detection',   
		    'Software Protection'; 
		)

$computer = "LocalHost"
$namespace = "root\cimv2"  

Write-Output "Checking if there are any required Services that are not started..."

$AutoServices = Get-WmiObject -class Win32_Service -computername $computer -ErrorAction Stop | Where {$_.StartMode -eq 'Auto' -and $Ignore -notcontains $_.DisplayName -and $_.State -ne 'Running'}

        If($AutoServices -eq $null) 

             {"Services which are set to Automatic are started."}

        Else 
    
{"Attempting to start Services..."		
			
ForEach ($Service in $AutoServices){

$Service.StartService()

$Pause = Start-Sleep 5

"Checking if any required services are not started..."

$AutoServices = Get-WmiObject -class Win32_Service -computername $computer -namespace $namespace -ErrorAction Stop | Where {$_.StartMode -eq 'Auto' -and $Ignore -notcontains $_.DisplayName -and $_.State -ne 'Running'} | FL DisplayName

			If($AutoServices -eq $null) 

                 {"Services which are set to Automatic are now started."}

            Else {"Some Services did not start."}
			return $AutoServices

}
}