$date = "7/17/2013 10:10:10 AM".ToString()
"date = {0}" -f $date

$formats = [string[]]("M/d/yyyy h:m:s tt","M/dd/yyyy h:m:s tt","MM/d/yyyy h:m:s tt","MM/dd/yyyy h:m:s tt")
$formattedDate = [datetime]::ParseExact($x, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None)

"Formatted Date = {0}" -f $formattedDate.ToLongDateString() 