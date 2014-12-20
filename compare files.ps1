$filename=("C:\Avanade\CAB\OWA Compare Files {0}.txt" -f $env:ComputerName)


$line="--------------------------------------------------"

function Log 
{
	param([string]$text)
	#Output to logfile
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
	#Output to screen
	Write-Output $text 
}

$Folder = "C:\Avanade\CAB\"

$owafiles = Get-ChildItem -Path $Folder -Filter "owa-web*"
$exchwebfiles = Get-ChildItem -Path $Folder -Filter "exchweb-*"

$owafile = ($owafiles |Sort CreationTime -Descending |Select-Object -First 1).name
$exchwebfile = ($exchwebfiles |Sort CreationTime -Descending |Select-Object -First 1).name

$owafile = $Folder + $owafile
$exchwebfile = $Folder + $exchwebfile

$currentfile = "C:\Program Files\Microsoft\Exchange Server\V14\ClientAccess\Owa\web.config"

$pattern = ".*"

$content1 = Get-Content $currentfile 
$content2 = Get-Content $owafile

$comparedLines = Compare-Object $content1 $content2 -IncludeEqual | 
    Sort-Object { $_.InputObject.ReadCount } 
    
$lineNumber = 0 
$comparedLines | foreach {

    if($_.SideIndicator -eq "==" -or $_.SideIndicator -eq "=>") 
    { 
        $lineNumber = $_.InputObject.ReadCount 
    } 
    
    if($_.InputObject -match $pattern) 
    { 
        if($_.SideIndicator -ne "==") 
        { 
            if($_.SideIndicator -eq "=>") 
            { 
                $lineOperation = "added" 
            } 
            elseif($_.SideIndicator -eq "<=") 
            { 
                $lineOperation = "deleted" 
            }

			$log = New-Object psobject
			$log | Add-Member -MemberType NoteProperty -Name Line -Value $lineNumber
			$log | Add-Member -MemberType NoteProperty -Name Operation -Value $lineOperation
			$log | Add-Member -MemberType NoteProperty -Name Text -Value $_.InputObject
	Log ("<----{2} - {1} - {0}" -f $lineNumber, $lineOperation, $_.InputObject) | Out-Null
 
            
        } 
    } 
}
Write-Output "Log located at : $FileName"

$filename=("C:\Avanade\CAB\EXCHweb Compare Files {0}.txt" -f $env:ComputerName)
$currentfile = "C:\Program Files\Microsoft\Exchange Server\V14\ClientAccess\exchweb\ews\web.config"
$content1 = Get-Content $currentfile 
$content2 = Get-Content $exchwebfile

$comparedLines = Compare-Object $content1 $content2 -IncludeEqual | 
    Sort-Object { $_.InputObject.ReadCount } 
    
$lineNumber = 0 
$comparedLines | foreach {

    if($_.SideIndicator -eq "==" -or $_.SideIndicator -eq "=>") 
    { 
        $lineNumber = $_.InputObject.ReadCount 
    } 
    
    if($_.InputObject -match $pattern) 
    { 
        if($_.SideIndicator -ne "==") 
        { 
            if($_.SideIndicator -eq "=>") 
            { 
                $lineOperation = "added" 
            } 
            elseif($_.SideIndicator -eq "<=") 
            { 
                $lineOperation = "deleted" 
            }

			$log = New-Object psobject
			$log | Add-Member -MemberType NoteProperty -Name Line -Value $lineNumber
			$log | Add-Member -MemberType NoteProperty -Name Operation -Value $lineOperation
			$log | Add-Member -MemberType NoteProperty -Name Text -Value $_.InputObject
	Log ("<----{2} - {1} - {0}" -f $lineNumber, $lineOperation, $_.InputObject) | Out-Null
 
            
        } 
    } 
}
Write-Output "Log located at : $FileName
**Script Completed** "