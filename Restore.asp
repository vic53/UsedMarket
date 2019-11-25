<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="check.asp" -->
<%
public filname(10,2), k, DestPath, i, formname, ThisProgram, versione
Dim flagHeader(10,5), pgmtitle, msg, cf, admin
ThisProgram="restore.asp"
formname=left(ThisProgram, len(ThisProgram) - 4)
DestPath = Server.MapPath("database\")
'
flagHeader(1,1)="Y"
flagHeader(1,2)="icons/homepage.gif"
flagHeader(1,3)="HomePage()"
flagHeader(1,4)="Torna alla Homepage..."
flagHeader(2,1)="N"
flagHeader(2,2)="icons/backpage.gif"
flagHeader(2,3)="BackPage()"
flagHeader(2,4)="Torna al form precedente..."
flagHeader(3,2)="icons/close.jpg"
flagHeader(3,3)="ClosePage()"
flagHeader(3,4)="Chiusura immediata sessione..."
flagHeader(4,1)="N"
'flagHeader(4,2)="icons/save.gif"
'flagHeader(4,3)="SalvaScheda()"
'flagHeader(4,4)="Salva la scheda anagrafica..."
pgmtitle="Restore archivi del mercatino"
msg=Session("LASTMSG")
Session("LASTMSG")=""
%>
<!-- #include file="adovbs.inc" -->
<!-- #include file="nav_web.asp" -->
<%
admin="admin"
if admin<>Session("USERNAME") then
   Session("LASTMSG")="PERMESSI INSUFFICIENTI..."
   response.redirect ("index.asp")
end if   
%>
<!-- #include File="check.asp"-->
<!-- #include file="Connessioni.asp" -->
<%

dim mese, dtStampa, mesi
mesi="x,gennaio,febbraio,marzo,aprile,maggio,giugno,luglio,agosto,settembre,ottobre,novembre,dicembre"
mese=split(mesi,",")

dim filesys, filetxt, f, riga(10), x

SW=request.form("SW")
Set filesys = CreateObject("Scripting.FileSystemObject")
for i=1 to 3

Set f = filesys.GetFile(Server.MapPath("database") & "\BCK\BCK"+trim(i)+".accdb")
riga(i)= "BCK"+trim(i)+".accdb  data salvataggio: " & f.DateLastAccessed
next 
SW=request.form("SW")
'response.write "SW=" & SW
if SW="Y" then 
 for i=1 to Request.QueryString("versione").Count 
  x =  Request.QueryString("versione")(i) 
 next
 Const OverWrite = TRUE
 Const DeleteRdOnly = True
 SourceFile =Server.MapPath("database") & "\BCK\BCK" & trim(x) & ".accdb"
 'response.write SourceFile
 DestPath = Server.MapPath("database")&"\archivi.accdb"
 Set objFS = CreateObject("Scripting.FileSystemObject")
 ' Copy a single file to a new folder, overwrite any existing file in destination folder
 objFS.CopyFile SourceFile, DestPath, OverWrite
 Session("LASTMSG")="RESTORE: RIPRISTINO DATABASE BCK" & trim(x) & ".accdb SELEZIONATO ESEGUITO..."
end if
Set objFS = nothing
msg=Session("LASTMSG")
Session("LASTMSG")=""
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=PgmTitle%></title>
<!-- #include file="head.asp" -->
<script type="text/javascript" language="javascript">
function HomePage() {
  document.<%=formname%>.action="<% =menu%>";
  document.<%=formname%>.submit();
  return true;
    }
function ClosePage() {
var conferma = window.confirm("ESEGUIRE IL LOGOUT IMMEDIATO?..., PREMI OK PER CONFERMARE!");
if  (conferma)
  { 
  var Backlen=history.length;
  history.go(-Backlen); 
  location.href="logout.asp";
}}
function ClosePage() {
var conferma = window.confirm("ESEGUIRE IL LOGOUT IMMEDIATO?..., PREMI OK PER CONFERMARE!");
if  (conferma)
  { 
  var Backlen=history.length;
  history.go(-Backlen); 
  location.href="logout.asp";
}}


function RestoreDB (x) {
var truthBeTold = window.confirm("VUOI VERAMENTE EFFETTUARE IL RESTORE DEL DATABASE? ..., PREMI OK PER CONFERMARE!");
if  (truthBeTold) 
  { 
  document.<%=formname%>.SW.value="Y";
  document.<%=formname%>.action = "Restore.asp?versione="+x;
  document.<%=formname%>.submit();
  return true;  
  }
}

function  CallHelp(URL, frmname){
ModalWin(URL,frmname,1024,768);
}
function UserLogin(){
ModalWin("Rlogin.asp","Login page",550,300);
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

</script>
<style type="text/css">
<!--
body {
    background-color: #CCCCCC;
}
.txt {
    font-family: Verdana, Arial, Helvetica, sans-serif;
    font-size: 14px;
    color: #FFFFFF;
}
.style1
    {
        width: 6%;
    }
.Stile99 {
    font-family: Verdana, Arial, Helvetica, sans-serif;
    color: #0000CC;
}
.Stile100 {color: #000000}
.Stile101 {color: #FFFF00}

.P1024 {
 width:1024px;
 height:768px;
 margin:auto;
 background-color: #333333;
    color: #FFFF00;
}
body {
    background-color: #666666;
}
.btn {
 width:100px;
 height:25px;
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  14px;
 color: blue;
 cursor: pointer;
 border-radius: 6px; 
 -moz-border-radius: 6px; /* firefox */
 -webkit-border-radius: 6px; /* safari, chrome */
}
.btn:hover {
 background-color: orange;
}
-->
</style>
</head>
<body>
<div class="P1024">
<form name="<%=formname%>" id="<%=formname%>" action="" method="post"  >
<!-- #include file="header.asp" -->

<input type="hidden" name="SW" value="" />
 
<div align="left" class="Stile99" style="background-color:#FFFFFF;">&nbsp;<span class="Stile100"><%=msg%></span></div>

<table width="49%" border="0" align="center" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC">
<br />
<%
for i= 1 to 3
%>
 <tr>
    <td width="82%" bgcolor="#FFCC33" class="Stile100"><%=riga(i)%></td>
    <td width="18%" bgcolor="#FFCC33" ><input type="button" value=" RESTORE " class="btn" onclick="RestoreDB(<%=i%>)" /></td>
</tr>
<%
next
%>
</table>
<p class="txt"><strong>ATTENZIONE</strong>: eseguendo la funzione Restore, si perdono tutte le informazioni relative agli aggiornamenti e inserimenti effettuate dopo il salvataggio di cui si esegue il Restore. Utilizzare la funzione con cautela per evitare di perdere informazioni
    utili del sistema.</p>
</div>
</form>
</body>
</html>
<!-- #include file="Disconnessioni.asp" -->
<%
sub ScriviDataOra()
dt=date
dtStampa=day(dt)&" " & mese( month(dt)) & " " & year(dt) & " - " & Hour(time)&":"&minute(time)
response.write(dtStampa)
end sub
%>