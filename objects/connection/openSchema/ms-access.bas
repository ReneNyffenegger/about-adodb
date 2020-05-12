'
'  Add reference to ADODB with
'
'      application.VBE.vbProjects(1).references.addFromGuid guid := "{B691E011-1797-432E-907A-4D8C69339129}", major := 6, minor := 1
'

option explicit

sub main() ' {

    dim conn as adodb.connection

    set conn = currentProject.connection

    createSchema conn

    showSchema conn

end sub ' }

sub showSchema(conn as adodb.connection) ' {

    dim rs as adodb.recordSet

    set rs = conn.openSchema(adSchemaPrimaryKeys, array(empty, empty, "tab_p"))

    debug.print("Primary key of tab_p:")
    do while not rs.eof ' {

       debug.print("  Column Name     : " & rs!COLUMN_NAME)
       debug.print("  Primary key name: " & rs!PK_NAME)

       rs.moveNext
    loop ' }

  ' --------------------------------------------
    
    set rs = conn.openSchema(adSchemaForeignKeys, array(empty, empty, empty, empty, empty, "tab_c"))

    debug.print("Foreign key of tab_c:")
    do while not rs.eof ' {
       debug.print("  Column Name     : " & rs!FK_COLUMN_NAME)
       debug.print("  Foreign key name: " & rs!FK_NAME       )
       debug.print(" references")
       debug.print("  Table name      : " & rs!PK_TABLE_NAME )
       debug.print("  Column name     : " & rs!PK_COLUMN_NAME)

       rs.moveNext
    loop ' }

end sub ' }

sub createSchema(conn as adodb.connection) ' {

    dim rs as adodb.recordSet

    dropTableIfExists conn, "tab_c"
    dropTableIfExists conn, "tab_p"

    conn.execute("create table tab_p (id number primary key, val varchar(10))")
    conn.execute("create table tab_c (id number primary key, id_p number references tab_p, val varchar(10))")

end sub ' }

sub dropTableIfExists(conn as adodb.Connection, tabName as string) ' {

    if not isNull(DLookup("name", "MSysObjects", "Name='" & tabName & "' and type = 1")) then
       conn.execute("drop table " & tabName)
    end if

end sub ' }
