Option Explicit
Dim obj

Set obj = CreateObject("wscript.shell")

dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://the.earth.li/~sgtatham/putty/latest/w64/putty.exe", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "C:\Users\nisha\AppData\Local\Temp\putty2.exe", 2 '//overwrite
end with

obj.Exec("C:\Users\nisha\AppData\Local\Temp\putty2.exe")
