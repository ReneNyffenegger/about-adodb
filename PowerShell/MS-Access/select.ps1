$curDir     = get-location
$accessFile ="$($curDir)\test-db.accdb"

$provider =
   # "Provider=Microsoft.ACE.OLEDB.12.0;" +
     "Provider=Microsoft.Jet.OLEDB.4.0;"  +
     "Data Source=$accessFile"

$adoConnection = new-object -comObject ADODB.connection
$adoConnection.connectionString = $provider
$adoConnection.open()

$recordSet = $adoConnection.execute(@'
  select
     num,
     txt
  from
     tab_one
'@)

while (! $recordSet.eof) {
   write-output "$($recordSet.fields('num').value) | $($recordSet.fields('txt').value)"
   $recordSet.moveNext()
}

$adoConnection.close()
