'get last popular images

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1

Set xml = CreateObject("Microsoft.XMLHTTP")
URL2="https://api.instagram.com/v1/media/popular?client_id=e8d6b06f7550461e897b45b02d84c23e"

xml.Open "GET", URL2, False
xml.Send
jsonTxt=xml.responseText
'Msgbox  jsonTxt


Dim WshShell
Dim strSafeDate, strSafeTime, strDateTime
Set WshShell = CreateObject("WScript.Shell")
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

Set htmlFSO = CreateObject( "Scripting.FileSystemObject" )
Set htmlFile = htmlFSO.OpenTextFile( a+strSafeDate+strSafeTime+"_popular.html", 2, True )
htmlFile.Write("<html><head></head><body> Images for tag "+a+"<br>")

start=1
do while 1<10
'Msgbox start

ipos1=InStr(start, jsonTxt,"standard_resolution") +29
'Msgbox ipos1

if ipos1<30 then exit do 

ipos2=InStr(ipos1, jsonTxt,"""")
'Msgbox ipos2

start=iPos2

img1URL=mid(jsonTxt,ipos1, ipos2-ipos1)
'Msgbox img1URL

img1URL2=replace(img1URL,"\","")
'Msgbox img1URL2


fnPos=InStrRev(img1URL2,"/")
urlLen=Len(img1URL2)

rem  standard_resolution":{"url":"

ImageFile = Mid(img1URL2,fnPos+1, urlLen-fnPos)
'Msgbox ImageFile

usernamePosStrart=InStr(start, jsonTxt,"username") +11
usernamePosEnd=InStr(usernamePosStrart, jsonTxt,"""")
usernameTxt=mid(jsonTxt,usernamePosStrart, usernamePosEnd-usernamePosStrart)
'Msgbox usernameTxt

xml.Open "GET", img1URL2, False
xml.Send

set oStream = createobject("Adodb.Stream")
oStream.type = adTypeBinary
oStream.open
oStream.write xml.responseBody
oStream.savetofile usernameTxt+"_"+ImageFile, adSaveCreateOverWrite

oStream.close

 htmlFile.Write("<a href=http://instagram.com/"+usernameTxt+">")
 htmlFile.Write("<img src="+usernameTxt+"_"+ImageFile +">")
 htmlFile.Write("</a><br>")


Loop

set oStream = nothing
Set xml = Nothing



htmlFile.Write("</body></html>")
htmlFile.Close( )
