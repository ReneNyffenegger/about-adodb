' \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word anonymous_block -c go rene rene ORA_MANUALLY_CREATED

option explicit

sub go(dbUser as string, dbPassword as string, dbName as string) ' {

  ' ADODB
' call addReference(application, "{2A75196C-D9EB-4129-B803-931327F72D5C}") 

  dim cn as ADODB.connection
  set cn = openConnection(dbUser, dbPassword, dbName)

  dim plsql as string

' Obviously, a block cannot start with declare
'
' See also http://stackoverflow.com/questions/2373401/with-ado-how-do-i-call-an-oracle-pl-sql-block-and-specify-input-output-bind-var
'
  plsql =         "begin "

  plsql = plsql & "declare"
  plsql = plsql & "  num_in  number := ?;"
  plsql = plsql & "  num_out number; "
  plsql = plsql & "begin"
  plsql = plsql & "  num_out := num_in * 5;"
  plsql = plsql & "  ? := num_out;"
  plsql = plsql & "end;"

  plsql = plsql & "end;"

  dim cm as ADODB.command
  set cm = new ADODB.command
  set cm.activeConnection = cn
  cm.commandText = plsql
  cm.commandType = adCmdText
    cm.parameters.append cm.createParameter(, adDouble, adParamInput,, 3)
    cm.parameters.append cm.createParameter(, adDouble, adParamOutput)

  cm.execute,,adExecuteNoRecords

  msgBox(cm.parameters(1))

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
              msgBox("Oracle error while opening connection: " & err.description)
  else 
              msgBox(err.number & " " & err.description)
  end if

end function ' }
