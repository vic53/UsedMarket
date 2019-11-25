<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="check.asp" -->
<%
public filname(10,2), k, DestPath, i, formname, ThisProgram
Dim flagHeader(10,5), pgmtitle, msg, cf
ThisProgram="backup.asp" 
formname=left(ThisProgram, len(ThisProgram) - 4)

DestPath = Server.MapPath("database\")
'     sw,src,page,descr
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
pgmtitle="Backup archivi del mercatino"
msg=Session("LASTMSG")
Session("LASTMSG")=""

%>
<!-- #include file="adovbs.inc" -->
<!-- #include file="nav_web.asp" -->
<%
 filname(1,1) = DestPath+"\BCK\BCK2.accdb"
 filname(1,2) = DestPath+"\BCK\BCK3.accdb"

 filname(2,1) = DestPath+"\BCK\BCK1.accdb"
 filname(2,2) = DestPath+"\BCK\BCK2.accdb"

dim orabackup, giornobackup

cf=request.form("CONFERMA")
if cf="Y" then 
  EseguiBackup
end if
%>
<html>
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

function Conferma(){
  document.<%=formname%>.CONFERMA.value="Y";
  document.<%=formname%>.action = "Backup.asp";
  document.<%=formname%>.submit();
  return true;
}

function REORGDB(){
 var conferma = window.confirm("ESEGUIRE LA RIORGANIZZAZIONE INDICI DEL DATABASE?...\n PREMI OK PER CONFERMARE!");
if  (conferma)
  { 
    var xhttp = new XMLHttpRequest();
    var ris;
    xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
       ris=this.responseText;
       alert(ris);
     }
    };
    xhttp.open("GET","CompactDatabase.asp", true);
    xhttp.send(); 

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
.Stile98 {
    font-family: Verdana, Arial, Helvetica, sans-serif;
    font-size: 12px;
    color: #000000;
}
.Stile99 {
    font-family: Verdana, Arial, Helvetica, sans-serif;
    font-weight: bold;
    font-size: 16px;
}
    .style1
    {
        color: #0000FF;
    }
    .style2
    {
        width: 5%;
    }
.Stile100 {color: #000000}
.Stile101 {font-size: 12px; font-family: Verdana, Arial, Helvetica, sans-serif;}

.rounded {
 width:350px;
 height:50px;
 font-weight:bold;
 font-size:20px;
 text-align:center;
 background-color: #cccccc;
 border-radius: 6px;
 box-shadow: 5px 5px 5px #ffffff;
 -moz-border-radius: 6px; /* firefox */
 -webkit-border-radius: 6px; /* safari, chrome */
 }
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

.auto-style1 {
    color: #FFFF00;
}

.auto-style2 {
    color: #FFFFFF
}
.auto-style3 {
    font-family: Verdana, Arial, Helvetica, sans-serif;
    font-size: 12px;
    color: #FFFFFF;
}
.btn {
 width:200px;
 height:45px;
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  14px;
 color: blue;
 cursor: pointer;
 font-weight: bold;
 box-shadow: 5px 5px 5px #ffffff;
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
<p class="Stile100">
  <% if cf="Y" then %>
</p>
<p class="Stile100"><strong><span class="auto-style1">BACKUP eseguito alle ore 
 <%=orabackup%> del <%=giornobackup%></span></strong></p>
<span class="Stile100">
<%
end if
%>
<span class="Stile99">
<input type="hidden" name="CONFERMA" value=""  />
</span></span> 
<span class="Stile99">
<span class="auto-style2">
Conferma la richiesta di BACKUP del sistema &nbsp;
</span> 
<span class="Stile100">
<input type="button" class="btn" value="CONFERMA" onClick="Conferma()" />
</span> 
</span> 
<span class="Stile100"><br>
</span>
<p class="Stile98"><span class="auto-style2">Il sistema di backup prevede il salvataggio della copia dei database file archivi.accdb nel file BCK1.accdb; prima di copiare il database, cancella
la terza copia che rimpiazza con la seconda copia, poi rimpiazza la seconda copia con la prima copia. <br>
Questo per mantenere soltanto gli ultimi 3 database usati.</span>
</p>
<p class="auto-style3">E' consigliabile eseguire il backup ogni sera alla fine della giornata lavorativa.

<p class="auto-style3">I files salvati hanno i seguenti nomi nella cartella BCK del server:</p>
<table width="30%" border="0" align="center" cellpadding="5" cellspacing="6" >
  
  <tr>
    <td class="rounded">BCK1.accdb</td>
  </tr>
  <tr>
    <td class="rounded">BCK2.accdb</td>
  </tr>
  <tr>
    <td class="rounded">BCK3.accdb</td>
  </tr>
</table>
<p class="auto-style3">Questi files sono salvati nella cartella applicativa  &quot;database\BCK&quot; del server INTRANET. </p>
<p align="justify" class="auto-style3"> Il file BCK1.mdb &egrave; l'ultimo salvataggio eseguito. La procedura di <b>RESTORE</b> permette di scegliere quale dei file utilizzare per il ripristino del database.  Il sistema accetta il ripristino solo se si &egrave; effettuato l'accesso con l'utente amministratore del sistema.</p>
<p align="justify"><br>NB. Se si riscontrassaro anomalie o lentezza nelle ricerche dei dati, è previsto eseguire un REORG del databse tramite una funzione di compattazione DB.
E' consigliabile eseguire periodicamente la riorganizzazione del database se questo è usato massicciamente e con molte migliaia di record.
<br> PER eseguire la riorganizzazione del database cliccare sul bottone [REORG DATABASE].<br><center>
<input type="button" class="btn" value="REORG DATABASE" onclick="REORGDB()"  name="btnReorg" /></center>
</form>
</div>
</body>
</html>
<%
sub EseguiBackup()
Session.TimeOut=1200
Server.ScriptTimeout=1200   
Const OverWrite = TRUE

Const DeleteRdOnly = True

Set filesys = CreateObject("Scripting.FileSystemObject")

FromSourceFile=filname(1,1)
ToDestinationFile=filname(1,2)
if (filesys.FileExists(FromSourceFile)) then
    filesys.CopyFile FromSourceFile, ToDestinationFile, true
end if

FromSourceFile=filname(2,1)
ToDestinationFile=filname(2,2)
if (filesys.FileExists(FromSourceFile)) then
    filesys.CopyFile FromSourceFile, ToDestinationFile, true
end if

SourceFile = Server.MapPath("database")
FromSourceFile=SourceFile+"\Archivi.accdb"
ToDestinationFile=DestPath+"\BCK\BCK1.accdb"
filesys.CopyFile FromSourceFile, ToDestinationFile, true 
orabackup=formatdatetime(time,4)
giornobackup=formatdatetime(date,2)
Session.TimeOut=1200
Server.ScriptTimeout=1200
end sub
%>