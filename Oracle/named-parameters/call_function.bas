'
'   Adding ADODB reference to VBA project:
'     thisWorkbook.VBProject.references.addFromGuid guid := "{2A75196C-D9EB-4129-B803-931327F72D5C}", major := 0, minor := 0
'

option explicit

sub main(dbUser as string, dbPassword as string, dbName as string) ' {

  dim cn as ADODB.connection
  set cn = openConnection(dbUser, dbPassword, dbName)

  dim cm as new ADODB.command
  set cm.activeConnection = cn

  dim retVal as ADODB.parameter
  dim p_two  as ADODB.parameter

  dim outSize as long: outSize = 1000

'
' Apparently, the returnd parameter needs not empty name. Thus, "anyName"
' is chosen:
'
  set retVal = cm.createParameter("anyName", adVarChar, adParamReturnValue, outSize,""   )

'
' We want to pass the second paramter but leave the first parameter as default:
'
  set p_two  = cm.createParameter("p_two"  , adNumeric, adParamInput      ,  , 99)

  cm.commandText = "tq84.func"
  cm.commandType =  adCmdStoredProc

'
' In order to pass paramters by name rather than by position, the
' namedParameters property of the command object needs to be set
' to true:
  cm.namedParameters = true

  cm.parameters.append retVal
  cm.parameters.append p_two

  cm.execute ' ,,adExecuteNoRecords

  debug.print "retVal: " & retVal.value

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
