' \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word ref_cursor -c go rene rene ORA_MANUALLY_CREATED

'  Priorly run
'    ref_cursor_db_objects.sql
'  for this example.
'  
' 
' ADODB
' call addReference(application, "{2A75196C-D9EB-4129-B803-931327F72D5C}") 

option explicit


sub go(dbUser as string, dbPassword as string, dbName as string) ' {


  dim cn as ADODB.connection
  set cn = openConnection(dbUser, dbPassword, dbName)


  dim rs as ADODB.recordSet

  dim cm as ADODB.command
  set cm = new ADODB.command
  set cm.activeConnection = cn

  cm.commandText = "tq84_refcursor_test_pck.proc_1"

  cm.commandType = adCmdStoredProc
  cm.parameters.append cm.createParameter("justAName", adDouble, adParamInput,, 3)
  cm.parameters.append cm.createParameter("justAName", adVarChar, adParamOutput,10)

  set rs = cm.execute ' , , adExecuteNoRecords

  do while not rs.eof
     
     dim i as long
     for i = 0 to rs.fields.count -1
         msgBox i & ": " & rs.fields(i)
     next
     rs.moveNext

  loop

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
              msgBox("Oracle Error while opening the database: " & err.description)
  else 
              msgBox(err.number & " " & err.description)
  end if

end function ' }
