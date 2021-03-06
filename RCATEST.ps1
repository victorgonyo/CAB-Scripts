[array]$all=@()
[string]$line=""
##reads each line in the file, and if it contains outlook.exe the outs the whole line to an array.
foreach ($line in (Get-Content 'C:\Users\victor.gonyo\Desktop\Scripts\stuff\RCA TEST.txt'))
{
If ($line -match "outlook.exe,")
{
$all+=$line
}
}

[int]$next=1
[string]$line=""
[int]$countmembers=0

##this while loop will remove everything before the version number.
While ($countmembers -lt $all.Count)
	{
	If ($next -eq 1)
		{
		[string]$line=$all[$countmembers]
		$next=0
		}
	If ($line.Substring(0,11) -notmatch "OUTLOOK.EXE")
		{
		$line=$line.Substring(1,$line.Length-1)
		}
	If ($line.Substring(0,11) -match "OUTLOOK.EXE")
		{
		$line = $line.substring(12,$line.Length-12)
		$all[$countmembers]=$line
		$next=1
		$countmembers++
		}
	}
[int]$next=1
[string]$line=""
[int]$countmembers=0
[int]$position=0
##this while loop removes everything after the version number.
While ($countmembers -lt $all.Count)
	{
	If ($next -eq 1)
		{
		[string]$line=$all[$countmembers]
		$next=0
		}
	If ($line.Substring($position,1) -ne ",")
		{
		$position++
		}
	If ($line.Substring($position,1) -eq ",")
		{
		$line=$line.Substring(0,$position)
		$all[$countmembers]=$line
		$next=1
		$position=0
		$countmembers++
		}
	}
##this is just here to display the array with the version numbers.	
$all