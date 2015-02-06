var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/media/popular?client_id=58e5502e27644cee9bb2770ec28213c2", false);
xmlHttpReq.send();
var popularJSON=xmlHttpReq.responseText;

var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");

 var myMsgBox=new ActiveXObject("wscript.shell")
// myMsgBox.Popup (objJSON.meta.code)

for(var i in objJSON.data)
{
     var imgUrl = objJSON.data[i].images.standard_resolution.url;
     //var name = objJSON.data[i].user.username;
     //myMsgBox.Popup ( imgUrl );
     
     
    xmlHttpReq.open("GET", imgUrl, false);
    xmlHttpReq.send();
    
            var objADOStream = new ActiveXObject("ADODB.Stream");
            objADOStream.open();
            objADOStream.type = 1; // Binary
            objADOStream.write(xmlHttpReq.responseBody);
            objADOStream.position = 0;
            objADOStream.saveToFile("file"+i+".jpg", 2);
            objADOStream.close();
    
}
