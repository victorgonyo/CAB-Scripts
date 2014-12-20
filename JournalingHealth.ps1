param([string]$magicalmailboxname
#Journaling Mailbox Health
#Scott Williams

#Pull Journal Databases Size & Whitespace

#US
	
$SizeandwhitespaceUS = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*us*"} | Sort-Object DatabaseSize -Descending | Format-Table Name, DatabaseSize, AvailableNewMailboxSpace
[array]$NameUS = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*us*"} | Sort-Object DatabaseSize -Descending | Select -ExpandProperty Name
[array]$DatabaseSizeUS = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*us*"} | Sort-Object DatabaseSize -Descending | select -ExpandProperty DatabaseSize
[array]$AvailableNewMailboxSpaceUS = 
#FF
	
$SizeandwhitespaceFF = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*FF*"} | Sort-Object DatabaseSize -Descending | Format-Table Name, DatabaseSize, AvailableNewMailboxSpace
		
#Pull Message count and size of Journal Mailboxes

#United States Journals

$JRNMBSizeandmessagecountUS = Get-MailboxDatabase *jrn* |Get-Mailbox| ? {$_.DisplayName -like "*JRN*" -and $_.Servername -like "*us*"} |Get-MailboxStatistics
		
#Frankfurt Journals
		
$JRNMBSizeandmessagecountFF = Get-MailboxDatabase *jrn* |Get-Mailbox| ? {$_.DisplayName -like "*JRN*" -and $_.Servername -like "*FF*"} |Get-MailboxStatistics

#Trespassed Mailboxes
	
#Pull list of trespassed United States Mailboxes

$TrespassedUS = Get-MailboxDatabase *jrn* |Get-Mailbox |? {$_.Alias -notlike "*jrn*" -and $_.Servername -like "*us*"}

#Pull List of trespassed Frankfurt	
		
$TrespassedFF = Get-MailboxDatabase *jrn* |Get-Mailbox |? {$_.Alias -notlike "*jrn*" -and $_.Servername -like "*us*"}

Write-host "

Journal Health Report Avaya US
	
	Database Size & Whitespace
	
		$SizeandwhitespaceUS
		
	Mailbox Message Counts
		
		$JRNMBSizeandmessagecountUS
		
	Trespassed Mailboxes
		
		$TrespassedUS
		
Journal Health Report Avaya FF
	
	Database Size & Whitespace
	
		$SizeandwhitespaceFF
		
	Mailbox Message Counts
		
		$JRNMBSizeandmessagecountFF
		
	Trespassed Mailboxes
		
		$TrespassedFF
"
