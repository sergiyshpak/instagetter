//////////////////////
///     REALLY COOL    mumbo jumbo
///  from  http://with-love-from-siberia.blogspot.com/2009/12/msgbox-inputbox-in-jscript.html
/////  because jscript does not have  normal input text dialog boxes
////////////
////    shit does not work directly from win64
///            evil billgatez

///// 
/// var myMsgBox=new ActiveXObject("wscript.shell");
///    ///Popup(strText,[nSecondsToWait],[strTitle],[nType]) 
/// var answer = myMsgBox.Popup("Save data?",0,"title",4 )
/// myMsgBox.Popup(" ansver is "+answer);    //// yes6  no7  cancel2

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
var clientId="85f422e5d2d84861a0e3c69856741c1f";

var searchTag = InputBox('Enter tag', 'insta');

var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");

var getURL="https://api.instagram.com/v1/tags/"+searchTag+"/media/recent?client_id="+clientId;


///start loop here, to get next next next

xmlHttpReq.open("GET", getURL, false);
xmlHttpReq.send();


var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");

var myMsgBox=new ActiveXObject("wscript.shell");


if (objJSON.meta.code != 200)
{
   myMsgBox.Popup("request error: " + objJSON.meta.error_message);
   WScript.Quit(1);
}
/// "code":400,"error_message":"The client used for authentication is no longer active."}}
///  new cli https://api.instagram.com/v1/tags/gun/media/recent?client_id=e8d6b06f7550461e897b45b02d84c23e
///
///  {"meta":{"error_type":"OAuthParameterException","code":400,"error_message":"The client_id provided is invalid and does not match a valid application."}}


var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var fh = fso.CreateTextFile("_"+searchTag+"_tag.html", 2, true); 
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

///  getURL=objJSON.pagination.next_url
/// do you want 20 more
/// var answer = myMsgBox.Popup("do you want 20 more?",0,"title",4 )
/// if ansver !=6    WScript.Quit(2);

fh.WriteLine( "</body></html>" ); 
fh.Close(); 

