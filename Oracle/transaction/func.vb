option explicit

dim connWriter as adodb.connection
dim connReader as adodb.connection
dim stmtInsert as adodb.command
dim parmInsert as adodb.parameter


sub main()

    setGlobalVariables

    connWriter.execute("create table tq84_ado_trx_test(txt varchar2(50), id number generated always as identity)")

    writeDataWithoutTransaction
    readData "After writing data without transaction"

    writeDataWithinTransactions
    readData "After writing data within transactions"

    connWriter.commitTrans
    readData "After commiting transaction"

    connWriter.execute("drop table tq84_ado_trx_test")

end sub


sub writeDataWithoutTransaction() ' {
    insertValue "inserted with autocommit"
end sub ' }

sub writeDataWithinTransactions() ' {

    connWriter.beginTrans
    insertValue "committed record"
    connWriter.commitTrans

    connWriter.beginTrans
    insertValue "rolled back record"
    connWriter.rollbackTrans

    connWriter.beginTrans
    insertValue "committed or rolled back?"

end sub ' }

sub insertValue(val as string) ' {
    parmInsert.value = val
    stmtInsert.execute
end sub ' }

sub readData(step as string) ' {
    debug.print step

    dim rs as adodb.recordSet
    set rs = connReader.execute("select txt from tq84_ado_trx_test order by id")

    do while not rs.eof
       debug.print "  " & rs!txt
       rs.moveNext
    loop
end sub ' }


sub setGlobalVariables() ' {

    set connWriter = openOracleConnection
    set connReader = openOracleConnection

    set stmtInsert = new adodb.command
    set stmtInsert.activeConnection = connWriter

    stmtInsert.commandText = "insert into tq84_ado_trx_test(txt) values(:txt)"
    stmtInsert.commandType =  adCmdText

    set parmInsert = stmtInsert.createParameter(":txt", adVarchar, adParamInput, 50)
    stmtInsert.parameters.append parmInsert

end sub ' }

function openOracleConnection() as adodb.connection ' {

    set openOracleConnection = new adodb.connection

    openOracleConnection.open _
         "Provider=OraOLEDB.Oracle.1;"   & _
         "Persist Security Info=False;"  & _
         "User ID=RENE;"                 & _
         "Password=RENE;"                & _
         "Data Source=Ora19;"            & _
         "Extended Properties="""""

end function ' }
