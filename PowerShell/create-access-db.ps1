#
#  Make sure the correct bitness of PowerShell is running!!!!
#
$adoxCat    = new-object -comObject ADOX.catalog
$curDir     = get-location
$accessFile ="$($curDir)\test-db.accdb"

#
#  Remove Access file if it already exists:
#
remove-item $accessFile -errorAction ignore


#
#  OLE DB Provider string
#
$provider =
   # "Provider=Microsoft.ACE.OLEDB.12.0;" +
     "Provider=Microsoft.Jet.OLEDB.4.0;"  +
     "Data Source=$accessFile"

$catalog = $adoxCat.create($provider) # | out-null

$tab_one = $catalog.execute(@'
  create table tab_one (
    num  integer primary key,
    txt  varchar(20) not null
  )
'@)

$adoxCat.activeConnection.close()
