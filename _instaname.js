//////////////////////
/////   mumbo jumbo from  http://with-love-from-siberia.blogspot.com/2009/12/msgbox-inputbox-in-jscript.html
/////  because jscript does not have  normal dialog boxes
////////////
var vb = {};
vb.Function = function(func)
{
    return function()
    {
        return vb.Function.eval.call(this, func, arguments);
    };
};
vb.Function.eval = function(func)
{
    var args = Array.prototype.slice.call(arguments[1]);
    for (var i = 0; i < args.length; i++) {
        if ( typeof args[i] != 'string' ) {
            continue;
        }
        args[i] = '"' + args[i].replace(/"/g, '" + Chr(34) + "') + '"';
    }
    var vbe;
    vbe = new ActiveXObject('ScriptControl');
    vbe.Language = 'VBScript';
    return vbe.eval(func + '(' + args.join(', ') + ')');
};
var InputBox = vb.Function('InputBox');
var MsgBox = vb.Function('MsgBox');

///////////////////////////////////
/////////


var searchTag = InputBox('Enter name', 'insta');

//https://api.instagram.com/v1/tags/gun/media/recent?client_id=e8d6b06f7550461e897b45b02d84c23e


var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/users/search?q="+searchTag+"&client_id=e8d6b06f7550461e897b45b02d84c23e", false);
xmlHttpReq.send();

var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");


var myMsgBox=new ActiveXObject("wscript.shell")
//myMsgBox.Popup (objJSON.meta.code)


if (objJSON.meta.code != 200)
{
   myMsgBox.Popup("request error, probably dead clientid");
   WScript.Quit(1);
}


//mscorolla7
// https://api.instagram.com/v1/users/search?q=mscorolla7&client_id=e8d6b06f7550461e897b45b02d84c23e
//{"meta":{"code":200},"data":[
//            {"username":"mscorolla7",
//            "profile_picture":"https:\/\/igcdn-photos-b-a.akamaihd.net\/hphotos-ak-xfa1\/t51.2885-19\/11352884_437620919744865_492722065_a.jpg",
//            "id":"328711111",
//            "full_name":"\u2702\ufe0fHairStylist\u2702 \ud83d\udc8bMakeupArtist\ud83d\udc8b"}]}
//myMsgBox.Popup (xmlHttpReq.responseText)

var personId=objJSON.data[0].id
myMsgBox.Popup (personId)

xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/users/"+personId+"/media/recent/?client_id=e8d6b06f7550461e897b45b02d84c23e",false)
xmlHttpReq.send();

var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");



/// "data":[]}
if (objJSON.data.length == 0)
{
   myMsgBox.Popup("no rezults found");
   WScript.Quit(1);
}


/// "code":400,"error_message":"The client used for authentication is no longer active."}}



var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var fh = fso.CreateTextFile("_"+searchTag+"_name.html", 2, true); 
fh.WriteLine( "<html><head></head><body> last 20 by tag "+searchTag+"<br>" );

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

//myMsgBox.Popup("More?",1,"ddd");


fh.WriteLine( "</body></html>" ); 
fh.Close(); 


var shell = WScript.CreateObject("WScript.Shell");
shell.Run( "cmd /c  start _"+searchTag+"_tag.html");

