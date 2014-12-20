get-pssnapin -registered | add-pssnapin -ErrorAction SilentlyContinue
$AllInfo=get-mailbox | select alias, LegacyExchangeDN
$ZAPEvaultEnableStringMaster="[Directory]
DirectoryComputerName = Blah.global.avaya.com
Sitename = 1

[Mailbox]
DistinguishedName = LegacyExchangeDN

[Folder]
Name = Mailboxroot
Suspended = False
Enabled = True"
$countmembers = 0
While ($countmembers -lt $AllInfo.count)
	{
	$Info=$AllInfo[$countmembers]
	$ZAPEvaultEnableStringSlave=$ZAPEvaultEnableStringMaster
	$Alias=$Info.alias
	[string]$LegacyExchangeDN=$Info.LegacyExchangeDN
	$ZAPEvaultEnableStringSlave=$ZAPEvaultEnableStringSlave.Replace("LegacyExchangeDN","$LegacyExchangeDN")
	Out-File "C:\Avanade\test\$Alias - ZAPEvaultTestenable.ini" -inputobject $ZAPEvaultEnableStringSlave -encoding unicode
	$countmembers++
	}
$ZAPEvaultEnableStringMaster="[Directory] 
DirectoryComputerName = USEVEAST2-2.global.avaya.com 
Sitename = Avaya

[ArchivePermissions]
ArchiveName=Evault Movetest
Zap=True"
