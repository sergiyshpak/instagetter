var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/media/popular?client_id=58e5502e27644cee9bb2770ec28213c2", false);
xmlHttpReq.send();
var popularJSON=xmlHttpReq.responseText;

var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");

 var myMsgBox=new ActiveXObject("wscript.shell")
// myMsgBox.Popup (objJSON.meta.code)

var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var fh = fso.CreateTextFile("pop.html", 2, true); 
fh.WriteLine( "<html><head></head><body> last 20 popular <br>" );

for(var i in objJSON.data)
{
     var imgUrl = objJSON.data[i].images.standard_resolution.url;
     var name = objJSON.data[i].user.username;
     //myMsgBox.Popup ( imgUrl );
    
    var ss=imgUrl.split("/");
    var filename= ss[ss.length-1] ;
     
    xmlHttpReq.open("GET", imgUrl, false);
    xmlHttpReq.send();
    var objADOStream = new ActiveXObject("ADODB.Stream");
    objADOStream.open();
    objADOStream.type = 1; // Binary
    objADOStream.write(xmlHttpReq.responseBody);
    objADOStream.position = 0;
    objADOStream.saveToFile(filename, 2);
    objADOStream.close();
    fh.WriteLine("<a href=http://instagram.com/"+name+">"+name+"</a><br><img src="+ filename +"><hr>")
}

fh.WriteLine( "</body></html>" ); 
fh.Close(); 

var shell = WScript.CreateObject("WScript.Shell");
shell.Run( "cmd /c  start _pop"+dateStr+timeStr+".html");
