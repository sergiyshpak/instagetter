Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1


a=InputBox("Enter insta name")     
'katwade

rece1="https://api.instagram.com/v1/users/search?q="
rece2="&client_id=e8d6b06f7550461e897b45b02d84c23e"
url1=rece1+trim(a)+rece2

'Msgbox url1
'https://api.instagram.com/v1/users/search?q=katwade&client_id=58e5502e27644cee9bb2770ec28213c2
' 213394178

'https://api.instagram.com/v1/media/popular?client_id=58e5502e27644cee9bb2770ec28213c2

'https://api.instagram.com/v1/tags/lovki/media/recent?client_id=58e5502e27644cee9bb2770ec28213c2

'

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
rece2="/media/recent/?client_id=e8d6b06f7550461e897b45b02d84c23e"
URL2=rece1+userId+rece2

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
