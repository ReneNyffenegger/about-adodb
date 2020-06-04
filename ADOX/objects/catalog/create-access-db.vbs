option explicit

dim fs
set fs = createObject("scripting.FileSystemObject")

'
'   Determine the absolute path of accdb file to be created.
'   The resulting accdb file will be in the same directory
'   where the vbs file is located.
'
dim accdbFile
accdbFile = fs.getParentFolderName(wscript.scriptFullName) & "\" & "created-from-vbs.accdb"

'
'   Check if the Access database was already created, and
'   delete it, if so:
'
if  fs.fileExists(accdbFile) then
    wscript.echo(accdbFile & " exists, going to delete it")
    fs.deleteFile(accdbFile)
end if

'
'   Use the adox.catalog object to create the Access database:
'
dim cat
set cat = createObject("adox.catalog")

cat.create("provider=Microsoft.ACE.OLEDB.12.0;" & _
           "data source=" & accdbFile)

'
'   The activeConnection property of the adox.catalog
'   object is an ADODB connection. It can be used
'   to execute SQL (DDL) statements:
'
dim con
set con = cat.activeConnection
con.execute("create table tab_one(id integer primary key, val varchar(10))")
con.execute("create table tab_two(id integer primary key, val varchar(10), id_one integer not null references tab_one)")
