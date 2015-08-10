var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/media/popular?client_id=e8d6b06f7550461e897b45b02d84c23e", false);
xmlHttpReq.send();
var popularJSON=xmlHttpReq.responseText;

var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");

 var myMsgBox=new ActiveXObject("wscript.shell");
 //myMsgBox.Popup (objJSON.meta.code)

var d = new Date();
var dateStr = (d.getFullYear()*100 + d.getMonth()+1)*100 + d.getDate();
var timeStr = (d.getHours()*100+  d.getMinutes())*100+ d.getSeconds();

var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var fh = fso.CreateTextFile("_pop"+dateStr+timeStr+".html", 2, true); 
fh.WriteLine( "<html><head></head><body> last 20 popular <br>" );

for(var i in objJSON.data)
{
    var username = objJSON.data[i].user.username;

    if (objJSON.data[i].type == "video")
    {
         var vidUrl = objJSON.data[i].videos.standard_resolution.url;
         var ssv=vidUrl.split("/");
         var videofilename= username+"_vid_"+ssv[ssv.length-1] ;
         
         xmlHttpReq.open("GET", vidUrl, false);
	 xmlHttpReq.send();
	 var objADOStream = new ActiveXObject("ADODB.Stream");
	 objADOStream.open();
	 objADOStream.type = 1; // Binary
	 objADOStream.write(xmlHttpReq.responseBody);
	 objADOStream.position = 0;
	 objADOStream.saveToFile(videofilename, 2);
	 objADOStream.close();
         fh.WriteLine("<a href="+videofilename+"> VIDEO <hr> ")
    }


    var imgUrl = objJSON.data[i].images.standard_resolution.url;
    var ss=imgUrl.split("/");
    var filename= username+"_"+ss[ss.length-1] ;
     
    xmlHttpReq.open("GET", imgUrl, false);
    xmlHttpReq.send();
    var objADOStream = new ActiveXObject("ADODB.Stream");
    objADOStream.open();
    objADOStream.type = 1; // Binary
    objADOStream.write(xmlHttpReq.responseBody);
    objADOStream.position = 0;
    objADOStream.saveToFile(filename, 2);
    objADOStream.close();
    fh.WriteLine("<a href=http://instagram.com/"+username+">"+username+"</a><br><img src="+ filename +"><hr>")
    
    
}

fh.WriteLine( "</body></html>" ); 
fh.Close(); 
