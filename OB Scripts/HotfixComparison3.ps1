$path = Read-host "Enter path to CSV"


$header = "<H3>Hotfix Comparison Table</H3>" 
$head = '<!--mce:0-->' 
$report = @() 


$servers = Import-csv $path


$roles = $servers | sort-object -unique Role | select role


Foreach($r in $roles){
$hotfixes = @() 
$result = @() 
$header = "<H3>Hotfix Comparison Table $($r.role)</H3>" 


$computers = $servers | where{$_.role -eq $r.role}


 
foreach ($computer in $computers) 
{ 
    foreach ($hotfix in (get-hotfix -computer $computer.name | select HotfixId)) 
    { 
        $h = New-Object System.Object 
        $h | Add-Member -type NoteProperty -name "Server" -value $computer.name 
        $h | Add-Member -type NoteProperty -name "Hotfix" -value $hotfix.HotfixId 
        $hotfixes += $h 
    } 
} 
     
$ComputerList = $hotfixes | Select-Object -unique Server | Sort-Object Server 
 
foreach ($hotfix in $hotfixes | Select-Object -unique Hotfix | Sort-Object Hotfix) 
{ 
    $h = New-Object System.Object 
    $h | Add-Member -type NoteProperty -name "Hotfix" -value $hotfix.Hotfix 
         
    foreach ($computer in $ComputerList) 
    { 
        if ($hotfixes | Select-Object |Where-Object {($computer.server -eq $_.server) -and 

($hotfix.Hotfix -eq $_.Hotfix)})  
        {$h | Add-Member -type NoteProperty -name $computer.server -value "Installed"} 
        else 
        {$h | Add-Member -type NoteProperty -name $computer.server -value "Missing"} 
    } 
    $result += $h 
} 

$result2 = $result

$result | ConvertTo-Html -head $head -body $header | Out-File $env:temp\InstalledHotfixes.html 
Invoke-Item $env:temp\InstalledHotfixes.html 

$report +="<H3>Hotfix Comparison Table $($r.role)</H3>" 
$add = $result2 | ConvertTo-Html -fragment
$report += $add 


}

Convertto-html -body $report | out-file $env:temp\ReportInstalledHotfixes.html
Invoke-Item $env:temp\ReportInstalledHotfixes.html 

