#Journaling Mailbox Health
#Scott Williams
#Pull Journal Databases Size & Whitespace
#US
$SizeandwhitespaceUS = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*US*"} | Sort-Object DatabaseSize -Descending | select Name, DatabaseSize, AvailableNewMailboxSpace
$SizeandwhitespaceUS
#FF
#$SizeandwhitespaceFF = Get-MailboxDatabase *jrn* -Status | ? {$_.Name -like "*EU*"} | Sort-Object DatabaseSize -Descending | select Name, DatabaseSize, AvailableNewMailboxSpace

#Pull Message count and size of Journal Mailboxes
#United States Journals

$JRNMBSizeandmessagecountUS = Get-MailboxDatabase *jrn* |Get-Mailbox| ? {$_.DisplayName -like "*JRN*" -and $_.Servername -like "*us*"} |Get-MailboxStatistics
Write-Verbose $JRNMBSizeandmessagecountUS
#Frankfurt Journals
#$JRNMBSizeandmessagecountFF = Get-MailboxDatabase *jrn* |Get-Mailbox| ? {$_.DisplayName -like "*JRN*" -and $_.Servername -like "*ff*"} |Get-MailboxStatistics

#Trespassed Mailboxes

#Pull list of trespassed United States Mailboxes
$TrespassedUS = Get-MailboxDatabase *jrn* |Get-Mailbox |? {$_.Alias -notlike "*jrn*" -and $_.Servername -like "*us*"} |Get-MailboxStatistics
$TrespassedUS
#Pull List of trespassed Frankfurt 
#$TrespassedFF = Get-MailboxDatabase *jrn* |Get-Mailbox |? {$_.Alias -notlike "*jrn*" -and $_.Servername -like "*ff*"} |Get-MailboxStatistics
                              
#Write-Output "
#
#Journal Health Report Avaya US
#
#Database Size & Whitespace"
#
#$SizeandwhitespaceUS
#
#Write-Output "
#Mailbox Message Counts
#"
#$JRNMBSizeandmessagecountUS
#
#Write-output "
#
#Trespassed Mailboxes
#
#"
#$TrespassedUS


#Write-host "
#
#Journal Health Report Avaya FF
#
#Database Size & Whitespace
#"
#$SizeandwhitespaceFF
#"
#Mailbox Message Counts
#"
#$JRNMBSizeandmessagecountFF
#"
#Trespassed Mailboxes
#"
#$TrespassedFF