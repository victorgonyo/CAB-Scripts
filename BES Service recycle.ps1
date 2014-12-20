Param([Switch]$RestartServices,[Switch]$StopServices)

get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue

Write-Output "Checking Bes Services..."

$a = Get-Service Blackberry* -erroraction silentlycontinue

if ($a -eq $null) 
    {
    write-output "Blackberry Services are NOT installed on this machine. Exiting..."
    exit
    }

$a = Get-Service "Blackberry Router" -ErrorAction SilentlyContinue

If ($a -ne $null)
    {
    Set-Variable -name BesServicesStop -value @("BlackBerry Controller", "BlackBerry Dispatcher","BlackBerry Router")
    Set-Variable -name BesServicesStart -value @("BlackBerry Router","BlackBerry Dispatcher","BlackBerry Controller") 
    }
    else
    {
    $BesServicesStop = @("BES - BlackBerry Controller","BES - BlackBerry Dispatcher")
    $BesServicesStart = @("BES - BlackBerry Dispatcher","BES - BlackBerry Controller")
    }
 
foreach($Service in $BesServicesStop)
    {
    $CheckService = Get-Service $Service
    $ServiceStatus = $CheckService.Status
    Write-Output "$Service is $ServiceStatus"
    }

If ($RestartServices -eq $true)
    {
    Write-Output "Stopping BES Services..."
    foreach($Service in $BesServicesStop)
        {
        if((Get-Service -Name $Service).Status -ne "Stopped") 
            {
            Write-Output "Stopping $Service..."
            Stop-Service -Name $Service
            if((Get-Service -Name $Service).Status -ne "Stopped")
                {
                Write-Output "$Service failed to stop. Exiting..."
                exit
                }
            else
                {
                Write-Output "$Service successfully stopped..."
                }
            }
               else
              {
                   Write-Output "$selectedService already stopped..."
               }
             }    
    Start-Sleep 5
    Write-Output "Starting BES Services..."
    foreach($Service in $BesServicesStart)
     {
       if((Get-Service -Name $Service).Status -ne "Running")
       {
           Write-Output "Starting $Service..."
             Start-Service -Name $Service
               if((Get-Service -Name $Service).Status -ne "Running")
           {    
                Write-Output "$Service failed to start. Exiting..."
                   exit
             }
               else
             {
                      Write-Output "$Service successfully started."
              }
               }
                else
                  {
                      Write-Output "$Service already running!"
                   }
                    } 

      Write-Output "Checking BES Services..."
      Write-Output "------------------------"
      Start-Sleep 2
      foreach($Service in $BesServicesStop)
       {
         $CheckService = Get-Service $Service
         $ServiceStatus = $CheckService.Status
         Write-Output "$Service is now $ServiceStatus"
        }
}