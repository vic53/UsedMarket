<%
Session.LCID=1036
Session.CodePage=1252
Session.TimeOut=120
Server.ScriptTimeout=2500 ' max timeout 5 minuti
dim msg, formname,ThisProgram ,pw,username
ThisProgram="menu.asp"
Session("NAVPAGE")=ThisProgram&";"
formname=left(ThisProgram, len(ThisProgram) - 4)
msg=Session("LASTMSG")
Session("LASTMSG")=""
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="check.asp" -->
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>menu mercatino dell'usato</title>
<!-- #include file="head.asp" -->
<script type="text/javascript" language="javascript" >
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
function Sele(N) {
//alert(N);
var pgm="menu.asp";
  switch (N) {
   case 1:
      pgm="Clienti.asp";
      break;
   case 2:
      pgm="Backup.asp";
      break;
   case 3:
      pgm="OggettiDepositati.asp";
      break;
   case 4:
      pgm="Restore.asp";
      break; 
   case 5:
      pgm="OggettiVenduti.asp";
      break; 
   case 6:
      pgm="Config.asp";
      break;
   case 7:
      pgm="OggettiScaduti.asp";
      break;
   case 8:
      pgm="Statistiche.asp";
      break;
   case 9:
      pgm="Vendite.asp";
      break; 
   case 10:
      pgm="Comunicazioni.asp";
      break; 
   case 11:
      pgm="NuovaVendita.asp";
      break;
   case 12:
      pgm="BuoniMerce.asp";
      break;
   case 13:
      pgm="Fatturazione.asp";
      break;
   case 14:
      pgm="GiornaleAffari.asp";
      break; 
   case 15:
      pgm="Cassa.asp";
      break; 
   case 16:
      pgm="InserimentoOggetto.asp";
      break;
   case 17:
      pgm="Resi.asp";
      break;
   case 18:
      pgm="Mandati.asp";
      break;
   case 19:
      pgm="Stampe.asp";
      break; 
   case 20:
      pgm="Logout.asp";
      break; 
       
   default :
      pgm="menu.asp";
      break;
  }
   document.<%=formname%>.action=pgm;
   document.<%=formname%>.submit();
   return true;

}


function  CallHelp(URL, frmname){
ModalWin(URL+".html",frmname,1024,768);
}
function UserLogin(){
ModalWin("Rlogin.asp","Login page",550,300);
}
//
// sel può valere true o false
//
function SetComboList(nomecombo,sel){
 var cc=document.getElementById(nomecombo);
 for(var i=0;i<cc.options.length;i++){
 if (cc.options[i].value=='tuovalore'){
   cc.options[i].selected=sel;
   }
 }
}

</script>
<style type="text/css">
.auto-style1 {
    font-family: Arial, Helvetica, sans-serif;
    font-size: x-small;
    text-align: right;
    width:97%;
}
.round_border
 {
    border-radius:6px;
    -moz-border-radius:6px; /* Firefox 3.6 and earlier */
    -webkit-border-radius:6px; /* Safari */
 }

td {
 background-color: #006666;
 height: 50px;
 width: 160px;
 text-align:center;
 border: medium;
 font-family:Arial, Helvetica, sans-serif;
 font-weight: bolder;
 color: white ;
 cursor: pointer;
 margin:auto;
 box-shadow: 5px 5px 5px #cccccc;
}

.P1024 {
 width:1024px;
 height:768px;
 margin:auto;
 background-image: url(images/mercatini_usato.jpg) ;
 background-repeat:no-repeat;
 background-position: center;
}
td {
 border-radius: 6px; 
 -moz-border-radius: 6px; /* firefox */
 -webkit-border-radius: 6px; /* safari, chrome */
}
td:hover {
 background-color: orange;
 color: black;
}
body {
 background-color: #cccccc;
}
,fpic {
 font-size:8px;
}
.header {
  width:98%;
  height:70px;
  background-color: #8DC5E2;
  border: outset 2px;
  margin: 5px;
  padding: 5px;
}
.flex-container {
  display: flex;
}
</style>

</head>

<body>
<form  name="<%=formname%>"  id="<%=formname%>" method="post" action="" >
<div class="P1024">
<div class="header">
 <div class="flex-container">
  <div><br><br><span style="text-align:center; color:blue; font-size:14px;"> Realizzato da V.Manarolla &copy;&nbsp;2019</span></div>
  <div><img src="images/titolo3d.png" alt="" title="titolo3d" width="461" height="68" border="0" /></div>
  <div><br><span style="text-align:center; color:blue; font-size:14px;">Trasformiamo il tuo usato in denaro...&nbsp;</span><img src="icons/help.gif" height="32" width="32" alt="manuale utilizzo..." onclick="CallHelp('menu')"
      onmouseover="this.style.cursor='pointer'" title="Help contestuale..." />&nbsp;<img src="icons/userH.ico" alt="login on timeout" title="login/logout..."  onclick="UserLogin()"
      onmouseover="this.style.cursor='pointer'" height="32" width="32" /></div>
 </div>
<div >
<br>
<table style="width: 600px" align="center" cellspacing="10" cellpadding="10" >
    <tr>
        <td style="width: 300px;" onclick="Sele(1)" title="Archivio clienti registrati nel mercatino..." >&nbsp;CLIENTI</td>
        <td style="width: 300px;" onclick="Sele(2)" title="Salvataggio archivio dati...">&nbsp; BACKUP ARCHIVI</td>
    </tr>
    <tr>
        <td style="width: 300px;" onclick="Sele(3)" title="Oggetti depositati nel mercatino...">&nbsp;OGGETTI DEPOSITATI</td>
        <td style="width: 300px;" onclick="Sele(4)" title="Ripristino archivio dati dopo errore o malfunzione...">&nbsp; RESTORE ARCHIVI</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(5)" title="Oggetti venduti dal mercatino...">&nbsp; OGGETTI VENDUTI</td>
        <td style="width: 300px;" onclick="Sele(6)" title="Configurazione parametri del gestionale mercatino...">&nbsp; CONFIGURAZIONE</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(7)" title="Oggetti con data scadenza maggiore di 30/60/90 giorni...">&nbsp; OGGETTI SCADUTI</td>
        <td style="width: 300px;" onclick="Sele(8)" title="Statistiche varie nella gestione del mercatino...">&nbsp;STATISTICHE</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(9)" title="Vendite registrate nel mercatino...">&nbsp; VENDITE</td>
        <td style="width: 300px;" onclick="Sele(10)" title="Archivio messaggi e comunicazioni...">&nbsp; COMUNICAZIONI</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(11)" title="Apertura di una vendita al cliente del mercatino...">&nbsp;NUOVA VENDITA</td>
        <td style="width: 300px;" onclick="Sele(12)" title="Buoni sostitutivi merce emessi dal mercatino...">&nbsp; GESTIONE BUONI MERCE</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(13)" title="Fatturazione elettronica...">&nbsp; FATTURAZIONE</td>
        <td style="width: 300px" onclick="Sele(14)" title="Giornale degli affari (obbligatorio per legge)...">&nbsp;GIORNALE DEGLI AFFARI</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(15)" title="Prima nota contabile di cassa (vendite/acquisti)...">&nbsp; CASSA</td>
        <td style="width: 300px;" onclick="Sele(16)" title="Registrazione oggetti depositati...">&nbsp;REGISTRAZIONE OGGETTI</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(17)" title="Oggetti restituiti al mercatino...">&nbsp; RESI</td>
        <td style="width: 300px" onclick="Sele(18)" title="Archivio mandati dei clienti registrati nel mercatino...">&nbsp; MANDATI</td>
    </tr>
    <tr>
        <td style="width: 300px" onclick="Sele(19)" title="Stampe del gestionale...">&nbsp; STAMPE</td>
        <td style="width: 300px;" onclick="Sele(20)" title="Logout e ritorna al form di accesso...">&nbsp; LOGOUT</td>
    </tr>
</table>
</div>

</div>

</form>
</body>

</html>