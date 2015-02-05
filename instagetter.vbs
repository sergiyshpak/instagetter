' takes latest picture from instagram user feed
'
'  first - get user id by name
'  second - get JSON with links
'  third - fing pic urg and save it locally
'
'   client_id can be received by registration process on instagramm

a=InputBox("Enter insta name")     
'katwade

rece1="https://api.instagram.com/v1/users/search?q="
rece2="&client_id=58e5502e27644cee9bb2770ec28213c2"
url1=rece1+trim(a)+rece2

'Msgbox url1
'https://api.instagram.com/v1/users/search?q=katwade&client_id=58e5502e27644cee9bb2770ec28213c2
' 213394178

Set xml = CreateObject("Microsoft.XMLHTTP")
xml.Open "GET", url1, False
xml.Send

'Msgbox xml.responseText

txt1=xml.responseText

rem ","id":"

pos1=InStr(txt1,""",""id"":""") +8
'Msgbox pos1
pos2=InStr(pos1, txt1,"""")
'Msgbox pos2

userId=mid(txt1,pos1, pos2-pos1)
'Msgbox userId

'https://api.instagram.com/v1/users/213394178/media/recent/?client_id=58e5502e27644cee9bb2770ec28213c2


rece1="https://api.instagram.com/v1/users/"
rece2="/media/recent/?client_id=58e5502e27644cee9bb2770ec28213c2"
URL2=rece1+userId+rece2

xml.Open "GET", URL2, False
xml.Send
txt2=xml.responseText
'Msgbox  txt2

ipos1=InStr(txt2,"standard_resolution")+29
'Msgbox ipos1

ipos2=InStr(ipos1, txt2,"""")
'Msgbox ipos2

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
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1

oStream.type = adTypeBinary
oStream.open
oStream.write xml.responseBody

 oStream.savetofile ImageFile, adSaveCreateOverWrite

oStream.close

set oStream = nothing
Set xml = Nothing
