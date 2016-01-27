Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1


Function processit(stringurl)
param1=stringurl
'Msgbox  param1
    xml.Open "GET", param1, False
xml.Send
txt2=xml.responseText
'Msgbox  txt2

start=1

endurl1 =InStr(start, txt2, "pagination"+Chr(34)+":{}")
'Msgbox endurl1

nexturl4 = "NO"
if endurl1 > 0 then 
   nexturl4 = "NO"
else
' {"pagination":{"next_url":"https:\/\/api.instagram.com\/v1\/users\/213394178\/media\/recent?max_id=1161796352561204531_213394178\u0026client_id=e8d6b06f7550461e897b45b02d84c23e","ne
nexturl1=InStr(start, txt2, "next_url") +11
nexturl2=InStr(start, txt2, "next_max_id")-3
nexturl3=mid(txt2,nexturl1, nexturl2-nexturl1)
' Msgbox nexturl3
nexturl39=replace(nexturl3,"\u0026","&")
nexturl4=replace(nexturl39,"\","")
' Msgbox nexturl4
' derk_razor_sharp
end if

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


img1URL2=replace(img1URL2,"s640x640/sh0.08/e35/","")
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


processit=nexturl4
End Function


'------------------------------
'---------------------


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

'https://api.instagram.com/v1/users/213394178/media/recent/?client_id=e8d6b06f7550461e897b45b02d84c23e


rece1="https://api.instagram.com/v1/users/"
rece2="/media/recent/?client_id=e8d6b06f7550461e897b45b02d84c23e"
URL2=rece1+userId+rece2


do 
rez=processit(URL2)
'Msgbox "result from processit "+rez
'rez1=processit(rez)              '     derk_razor_sharp
'rez2=processit(rez1)
'rez3=processit(rez2)
URL2=rez
Loop until rez="NO"
   

Set xml = Nothing