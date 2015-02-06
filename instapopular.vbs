'get last popular images

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1

Set xml = CreateObject("Microsoft.XMLHTTP")
URL2="https://api.instagram.com/v1/media/popular?client_id=58e5502e27644cee9bb2770ec28213c2"

xml.Open "GET", URL2, False
xml.Send
txt2=xml.responseText
'Msgbox  txt2

start=1
do while 1<10
'Msgbox start

ipos1=InStr(start, txt2,"standard_resolution") +29
'Msgbox ipos1

if ipos1<30 then exit do 

ipos2=InStr(ipos1, txt2,"""")
'Msgbox ipos2

start=iPos2

img1URL=mid(txt2,ipos1, ipos2-ipos1)
'Msgbox img1URL

img1URL2=replace(img1URL,"\","")
'Msgbox img1URL2


fnPos=InStrRev(img1URL2,"/")
urlLen=Len(img1URL2)


rem  standard_resolution":{"url":"


ImageFile = Mid(img1URL2,fnPos+1, urlLen-fnPos)

'Msgbox ImageFile


xml.Open "GET", img1URL2, False
xml.Send


set oStream = createobject("Adodb.Stream")

oStream.type = adTypeBinary
oStream.open
oStream.write xml.responseBody

 oStream.savetofile a+"_"+ImageFile, adSaveCreateOverWrite

oStream.close


Loop

set oStream = nothing
Set xml = Nothing
