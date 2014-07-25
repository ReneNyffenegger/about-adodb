' \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word stored_procedure -c go rene rene ORA_MANUALLY_CREATED
'
' Prior run
'    stored_procedure.bas
' for this example.
'
' ADODB
' call addReference(application, "{2A75196C-D9EB-4129-B803-931327F72D5C}") 

option explicit

sub go(dbUser as string, dbPassword as string, dbName as string) ' {


  dim cn as ADODB.connection
  set cn = openConnection(dbUser, dbPassword, dbName)

  dim cm as ADODB.command
  set cm = new ADODB.command
  set cm.activeConnection = cn
  cm.commandText = "tq84_procedure"
  cm.commandType = adCmdStoredProc

  cm.parameters.append cm.createParameter(, adDouble, adParamInput,, 10)
  cm.parameters.append cm.createParameter(, adDouble, adParamInput,, 20)

  cm.execute,,adExecuteNoRecords

  msgBox ("Check tq84_table, it has been filled with values")

end sub ' }

private function openConnection(dbUser as string, dbPassword as string, dbName as string) as ADODB.connection ' {

  on error goto error_handler

  dim cn as    ADODB.connection
  set cn = new ADODB.connection

  cn.open ( _
     "User ID="     & dbUser       & _
    ";Password="    & dbPassword   & _
    ";Data Source=" & dbName       & _
    ";Provider=MSDAORA.1")

  set openConnection = cn

  exit function

error_handler:
  if   err.number = -2147467259 then
              msgBox("Oracle Fehler beim Öffnen der Datenbankverbindung: " & err.description)
  else 
              msgBox(err.number & " " & err.description)
  end if

end function ' }
