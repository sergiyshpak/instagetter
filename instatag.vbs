Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1


a=InputBox("Enter insta tag")     

rece1="https://api.instagram.com/v1/tags/"
rece2="/media/recent?client_id=e8d6b06f7550461e897b45b02d84c23e"

Set xml = CreateObject("Microsoft.XMLHTTP")
URL2=rece1+a+rece2

xml.Open "GET", URL2, False
xml.Send
txt2=xml.responseText
'Msgbox  txt2

Set htmlFSO = CreateObject( "Scripting.FileSystemObject" )
Set htmlFile = htmlFSO.OpenTextFile( a+"_tag.html", 2, True )
htmlFile.Write("<html><head></head><body> Images for tag "+a+"<br>")

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


img1URL2=replace(img1URL2,"s640x640/sh0.08/e35/","")


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

oStream.savetofile "tag_"+a+"_"+ImageFile, adSaveCreateOverWrite
oStream.close

 htmlFile.Write("<img src="+ "tag_"+a+"_"+ImageFile +">")

Loop

set oStream = nothing
Set xml = Nothing

htmlFile.Write("</body></html>")
htmlFile.Close( )
