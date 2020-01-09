$curDir     = get-location
$accessFile ="$($curDir)\test-db.accdb"

$provider =
   # "Provider=Microsoft.ACE.OLEDB.12.0;" +
     "Provider=Microsoft.Jet.OLEDB.4.0;"  +
     "Data Source=$accessFile"

$adoConnection = new-object -comObject ADODB.connection
$adoConnection.connectionString = $provider
$adoConnection.open()

$insertStmt = new-object -comObject ADODB.command

$insertStmt.activeConnection = $adoConnection
$insertStmt.commandText      ='insert into tab_one values (:num, :txt)'
$insertStmt.commandType      = 1 # adCmdText

$paramNum = $insertStmt.createParameter('num',   3, 1,  4) #   3 = adInteger, 1 = adParamInput,  4 the size
$paramTxt = $insertStmt.createParameter('txt', 200, 1, 20) # 200 = adVarchar, 1 = adParamInput, 20 the size

$insertStmt.parameters.append($paramNum)
$insertStmt.parameters.append($paramTxt)

$paramNum.value = 1; $paramTxt.value ='one'  ; $insertStmt.execute() | out-null
$paramNum.value = 2; $paramTxt.value ='two'  ; $insertStmt.execute() | out-null
$paramNum.value = 3; $paramTxt.value ='three'; $insertStmt.execute() | out-null

$adoConnection.close()
