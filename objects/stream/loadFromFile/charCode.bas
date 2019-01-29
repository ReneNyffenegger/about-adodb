option explicit

sub main() ' {

    dim s as new adodb.stream
    s.charSet = "utf-8"

    s.open
    s.loadFromFile(environ$("userprofile") & "\utf-8.txt")

    dim txt as string
    txt = s.readText

    s.close

    debug.print(txt)

end sub ' }
