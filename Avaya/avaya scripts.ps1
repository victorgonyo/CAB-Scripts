get-pssnapin -registered|add-pssnapin -ErrorAction SilentlyContinue
$filename = "C:\Avanade\AFRICANOC.csv"
$countmembers=0
$Names=Import-Csv  $filename
While ($countmembers -lt $Names.count){
$Name=$Names[$countmembers].name
Set-Mailbox africanoc -grantsendonbehalfto @{Add="$Name"}
$countmembers++
}

$countmembers=0
$filename = "C:\Avanade\CRMS.csv"
$Names=Import-Csv  $filename
While ($countmembers -lt $Names.count){
$Name=$Names[$countmembers].name
Set-Mailbox crms -grantsendonbehalfto @{Add="$Name"}
$countmembers++
}
$countmembers=0
$filename = "C:\Avanade\INDIANOC.csv"
$Names=Import-Csv  $filename
While ($countmembers -lt $Names.count){
$Name=$Names[$countmembers].name
Set-Mailbox indianoc -grantsendonbehalfto @{Add="$Name"}
$countmembers++
}