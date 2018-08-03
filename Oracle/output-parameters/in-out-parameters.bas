' \lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word stored_procedure -c go rene rene ORA_MANUALLY_CREATED
'
' Prior run
'    in_out_parameters.sql
' for this example.
'
' ADODB
' call addReference(application, "{2A75196C-D9EB-4129-B803-931327F72D5C}")

option explicit

sub in_out_parameters(dbUser as string, dbPassword as string, dbName as string) ' {

  dim cn as ADODB.connection
  set cn = openConnection(dbUser, dbPassword, dbName)

  call executeProc(cn,  7.1#)
  call executeProc(cn, null)

end sub ' }

function variantToString(v as variant) as string ' {
    if isNull(v) then
       variantToString = "null"
    else
       variantToString = v
    end if
end function

sub executeProc(cn as adodb.connection, param_one as variant) ' {

    dim cm as new ADODB.command
'   set cm = new ADODB.command
    set cm.activeConnection = cn
    cm.commandText = "tq84_in_out_parameters.proc"
    cm.commandType = adCmdStoredProc

    dim param_two   as variant
    dim param_three as variant

    cm.parameters.append cm.createParameter("param_one"  , adDouble, adParamInput ,, param_one        )
    cm.parameters.append cm.createParameter("param_two"  , adDouble, adParamOutput,, param_two  )
    cm.parameters.append cm.createParameter("param_three", adDouble, adParamOutput,, param_three)

    cm.execute,,adExecuteNoRecords

    debug.print "called procedure with param_one = " & variantToString(param_one)
    debug.print "  param_two   = " & variantToString(cm.parameters("param_two"  ).value)
    debug.print "  param_three = " & variantToString(cm.parameters("param_three").value)

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
       msgBox("Error opening connection to oracle: " & err.description)
  else
       msgBox(err.number & " " & err.description)
  end if

end function ' }
