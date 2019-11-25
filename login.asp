<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Session.LCID=1036
Session.CodePage=1252
Session.TimeOut=120
Server.ScriptTimeout=2500 ' max timeout 5 minuti
dim msg, formname,ThisProgram ,pw,username
ThisProgram="login.asp"
formname=left(ThisProgram, len(ThisProgram) - 4)
msg=Session("LASTMSG")
Session("LASTMSG")=""
Session("TITOLOGESTIONE")="&nbsp;IME&nbsp;ELETTROFORNITURE&nbsp;"
TitoloGestionale=Session("TITOLOGESTIONE")
%>
<!--#include file="CampoAttivo.asp" -->
<html>
<head>
<title>Login al sistema</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8" >
<!-- #include file="head.asp" -->
   
<script type="text/javascript" language="javascript" >
function Connetti() {
// alert("logon");
// Imposto un'espressione regolare per verificare che
// i caratteri inseriti nei campi UserID e Password
// siano alfanumerici, in modo da non dar fastidio all'XML
  var re = /^[a-z0-9]/;
// Verifico che i campi siano valorizzati (correttamente)
   var u_id = document.getElementById("username").value;
   var pass = document.getElementById("pw1").value;
   if (re.test(u_id) == false || re.test(pass) == false)
    {
        alert("Inserire le credenziali di accesso!");
    }
    else
   document.<%=formname%>.SW.Value="1";
   document.<%=formname%>.action="logon.asp";
   var pw1 = document.getElementById("pw1").value;
   document.<%=formname%>.PWB.value=pw1;
   document.<%=formname%>.submit();
   return true;
}

function Invio(e){
var code = e.keyCode ? e.keyCode : e.charCode ;
   if (code == 13) {
    Connetti();
 }
}
function Alfanumerici(myfield, e, dec) {
  var key;
  var keychar;
  if (window.event)
    key = window.event.keyCode;
  else if (e)
    key = e.which;
  else
    return true;
  keychar = String.fromCharCode(key);
  // control keys
  if ((key==null) || (key==0) || (key==8) || (key==9) || (key==13) || (key==27) )
    return true;
  // tabella caratteri
  else if ((("abcdefghilmnopqrstuvzxwjkyABCDEFGHILMNOPQRSTUVZXWJKY0123456789.").indexOf(keychar) > -1))
    return true;
  else 
    return false;
}
function SetCursor(campo) {
 document.getElementById(campo).focus();
}
function VisualizzaLicenza() {
   modalWin("Lic.html","Licenza programma gestionale", 900, 650);
   }
// **************************************************
// utilizzo di FANCYBOX per aprire finestre modali
// **************************************************
function ModalWin(URL, titolo, X, Y) {
    //    alert(X + " x " + Y);
    $.fancybox.open({
    href : URL , 
    type: 'iframe',
    scrolling: false,
    padding : 5,
    autoDimensions: false,
    hideOnContentClick: false,
    hideOnOverlayClick: false,
    height: Y + 50,
    width: X,
    arrows: false,
    modal: false,
    closeBtn: true
    });
}

function modalWin(URL, titolo, X, Y) {
    $.fancybox.open({
    href : URL , 
    type: 'iframe',
    arrows: false,
    width: X,
    height: Y+50,    
    closeBtn: true
    });
}

</script>
<style type="text/css">
<!--
.Stile1 {
    color: #FFFFFF;
    font-weight: bold;
}
body {
    background-color: #cccccc;
    background-image:url('images/mercatini_usato.jpg');
    background-repeat:no-repeat;
    background-position:center;
}

.P1024 {
 width:1024px;
 height:768px;
 margin:auto;
 background-repeat:no-repeat;
 background-position: center;
}
.btn {
 width:120px;
 height:26px;
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  16px;
 color: blue;
 font-weight: bold;
 cursor: pointer;
 border-radius: 6px; 
 -moz-border-radius: 6px; /* firefox */
 -webkit-border-radius: 6px; /* safari, chrome */
}
.btn:hover {
 background-color: orange;
}
.txt {
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  16px;
 width: 150px;
 color: blue;
 background-color: white;
}



-->
</style>
<body>
<script>
        $(document).ready(function() {
            $('.fancybox').fancybox();
        });
</script>

<div class="P1024" >
<script type="text/html" src="http://www.cookie-bar.eu/cookiebar-latest.js"></script>
<form  name="<%=formname%>"  id="<%=formname%>" method="post" action="" >
<input type="hidden" id="SW" name="SW" value="" />
<input type="hidden" id="PWB" name="PWB" value="" />


  <table  style="width:100%; background-color:black;">
    <TR>
      <TD width="100%"  size="100%"><table width="100%" border="0">
        <tr>
          <td width="8%" >
              <img alt="" src="images/btn/btn1_username.jpg"></td>
          <td width="14%"><input type="text" name="username" id="username" value="<%=username%>" class="txt" 
           onKeyDown="Invio(event)" maxlength="50" onKeyPress="return Alfanumerici(this, event)" <%=crsx%> /></td>
          <td width="8%">
          <img alt="" src="images/btn/btn1_password.jpg"></td>
          <td ><input type="password" name="pw1" id="pw1" value="" class="txt" onKeyDown="Invio(event)" maxlength="50" <%=crsx%>  /></td>
          <td ><input type="button" name="btfrmLogin" id="btfrmLogin" value="Login " onClick="Connetti()" class="btn" onKeyPress="return Alfanumerici(this, event)"
          title="accesso al sistema" /></td>
          <td style="color:#ffffff;"><div align="center" title="software a cura di v.manarolla &copy; 2015 - licenza abilitata: 2014">
              <a onclick="VisualizzaLicenza()" class="link"><img alt="" src="images/logoform.jpg"></a></div></td>
        </tr>
      </table></TD>
    </TR>
    <TR>
      <TD width="100%" bgcolor="#CCCCCC" size="100%"><%=msg%></TD>
    </TR>
</table>
</form>
</div>
</body>
</html>