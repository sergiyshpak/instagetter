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


var searchTag = InputBox('Enter tag', 'insta');


rece1="https://api.instagram.com/v1/tags/"
rece2="/media/recent?client_id=58e5502e27644cee9bb2770ec28213c2"

var xmlHttpReq = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
xmlHttpReq.open("GET", "https://api.instagram.com/v1/tags/"+searchTag+"/media/recent?client_id=58e5502e27644cee9bb2770ec28213c2", false);
xmlHttpReq.send();


var objJSON = eval("(function(){return " + xmlHttpReq.responseText + ";})()");

var myMsgBox=new ActiveXObject("wscript.shell")
// myMsgBox.Popup (objJSON.meta.code)

var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var fh = fso.CreateTextFile(searchTag+"_tag.html", 2, true); 
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

fh.WriteLine( "</body></html>" ); 
fh.Close(); 

