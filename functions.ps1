#Generates a prompt that asks the user to input which computer to access

$computer= Read-Host "Enter the computer name"

If (!$computer)
{
$computer=$env:COMPUTERNAME
}

 

#Create a function that list the software installed on the computer specified by input argument

function getSoftwareInfo {

 

#Gets the Programs installed on the computer

$softwareInfo = gwmi Win32_Product -computername $computer |  select Vendor, Name
}

 

#Create a function that lists the free disk space and percentage of disk space used on the computer

function diskCheck {

 

#Gets the freespace of the C: Drive

$freeSpace = gwmi -class win32_logicaldisk -computername $computer  | where {$_.deviceid -eq "C:"} | select-object -expandproperty freespace
$freeSpace = $freeSpace / 1GB
#Gets the total space of the C: Drive

$totalSpace = gwmi -class win32_logicaldisk -computername $computer | where {$_.deviceid -eq "C:"} | select-object -ExpandProperty size
$totalSpace = $totalSpace / 1GB

 

#Gets percentage of disk space used

 

$percentUsed = (($totalSpace - $freeSpace)/$totalSpace) * 100

}

 

#calls the functions

getSoftwareInfo

diskCheck

 

#Create readable texts for both functions



$diskfree = "Free space available on C: drive: $freeSpace GB"

$diskused = "Percent of disk space used on C: $percentUsed %"

 

#create a custom psobject to accurately export software information using a foreach loop
$AllSoftware = $softwareinfo | ForEach-Object {
	$vendor=$_.vendor
	$name=$_.name
	New-Object PSObject -Property @{
	"Vendor" = "{0}" -f $vendor
	"Name" = "{0}" -f $name
	}
	}
#Exports softwareInfo to CSV file
$AllSoftware | Export-Csv -Path C:\Scripts\Computer_Software.csv -NoTypeInformation	


 

#creates customer psobject and Exports freespace and percentused to CSV file

New-Object -TypeName pscustomobject -Property @{
"Free Space" = $freeSpace
"Used Space" = $percentUsed
} | Export-Csv -path C:\Scripts\Computer_hdspace.csv -NoTypeInformation

 

#Writes output to the screen

Write-Output "The programs that are installed in the $computer are as follows:",$softwareInfo


Write-Host $diskfree

Write-Host $diskused

 