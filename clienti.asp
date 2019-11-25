<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="check.asp" -->
<!-- #include file="Connessioni.asp" -->
<%
Dim flagHeader(10,5), PgmTitle, msg, ThisProgram, previouspage, formname,numrecpage,npage,pagina,ctrRecord
ThisProgram="Clienti.asp"
formname=left(ThisProgram, len(ThisProgram) - 4)
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
flagHeader(4,1)="Y"
flagHeader(4,2)="icons/save.gif"
flagHeader(4,3)="NuovoCliente()"
flagHeader(4,4)="Registra un nuovo cliente..."
flagHeader(4,5)=" onmouseover=""this.src='icons/save1.gif'"" onmouseout=""this.src='icons/save.gif'"" "
PgmTitle="Lista anagrafica clienti"
msg=Session("LASTMSG")
Session("LASTMSG")=""
numrecpage=LeggiConfigurazione()
pagina = 1 'default di partenza
ctrRecord=ContatoreRecord("Clienti"," ATTIVO=1 " )
' se ricicla cambiando pagina questa è stata impostata dall'utente
for i=1 to Request.form("pagina").Count
  pagina =  Request.form("pagina")(i)
next
npage = ctrRecord / numrecpage
if int(npage)<> npage then
 npage = int(npage) + 1
end if
if (npage=1) and (numrecpage > ctrRecord) then
 numrecpage=ctrRecord
end if
%>
<!--#include file="CampoAttivo.asp" -->
<!--#include file="nav_web.asp" -->
<html>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<meta name="generator" content="HAPedit 3.1">
<head>
<title>Lista clienti</title>
<!-- #include file="head.asp" -->
<script type="text/javascript" src="w2ui/w2ui-1.5.rc1.js"></script>
<link rel="stylesheet" type="text/css" href="w2ui/w2ui-1.5.rc1.css" />


<script type="text/javascript" language="javascript" >
var ric="<%=filtro%>";
//alert("<%=filtro%>");

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

function BackPage() {
  document.<%=formname%>.action="<% =previouspage%>";
  document.<%=formname%>.submit();
  return true;
    } 
function  CallHelp(URL, frmname){
ModalWin(URL,frmname,1024,768);
}
function UserLogin(){
ModalWin("Rlogin.asp","Login page",550,300);
}

// chiamata POST di una pagina asp
// con parametri passati via GET
function OpenPage(page,param) {
//alert("test");
   var p=page;
   var d=param;
   document.<%=formname%>.action=p+".asp?"+d;
   document.<%=formname%>.submit();
   return true;
}

function NuovoCliente(){
   var conferma=window.confirm("Hai verificato che non esiste il nominativo, vuoi inserire una nuovo cliente?...CONFERMI?");
   if (conferma) {
      document.<%=formname%>.action="NuovaSchedaAnagrafica.asp";
      document.<%=formname%>.submit();
      return true;
   }
}
function DeleteCliente(){

   var nome = document.getElementById("ricerca").value;
   var ric = "?nominativo=" + nome;
   var yy=grid.getSelection() - 1;
   var xx=grid.getCellValue(yy, 1);
   if (xx != "")
   {
   var conferma=window.confirm("PER CANCELLARE IL CLIENTE...,\n verificato che non abbia oggetti in archivio sia venduti che depositati...,\n \nCONFERMI IL DELETE DELLA SCHEDA n."+xx);
   if (conferma) {
   var xhttp = new XMLHttpRequest();
   var ris;
   xhttp.onreadystatechange = function() {
   if (this.readyState == 4 && this.status == 200) {
       ris=this.responseText;
       alert(ris);
       w2ui["myGrid"].load("clientiJSON.asp"+ric);
     }
    };
    xhttp.open("GET","DeleteRecordCliente.asp?codiceanagrafico="+xx, true);
    xhttp.send(); 
    } 
   }
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

function Ricerca( fld ) {
//alert(fld);
if (isNaN(fld))
   {
   var ric = "?nominativo=" + fld;
   w2ui['myGrid'].load('clientiJSON.asp'+ric);
   }
else
   {
   var ric = "?codiceanagrafico=" + fld;
   w2ui['myGrid'].load('clientiJSON.asp'+ric);
   }
}

function StampaMandato() {
 var xx = w2ui['myGrid'].getSelection();
 if (xx!="")  {
     //alert("mandato del cliente n. " + xx);
     ModalWin("MandatoPDF.asp?codiceanagrafico="+xx,"Mandato di vendita",1024,768);
    }
}

function ReportCliente() {
      ModalWin("ListaAnagrafica.asp","Lista telefonica clienti",1024,768);
}
function EnterKey(e) {
 var code = e.keyCode ? e.keyCode : e.charCode 
// alert("premuto enter");
        if (code == 13) {
           Ricerca(ricerca.value)
      }  
 }

function ReloadPage(){
 var pagina=document.getElementById("pagina").value;
 w2ui['myGrid'].load('ClientiJSON.asp?pagina='+pagina);
}

function RimuoviClienti() {
  var yy = w2ui.myGrid.getSelection();
  if (yy.length==0) {
      alert("NON HAI ESEGUITO UNA SELEZIONE...");
      return false;
    }
  var str="";
// chiedo conferma
var CF = window.confirm("VUOI CANCELLARE I CLIENTI SELEZIONATI ?... PREMI OK PER CONFERMARE!");
  if (CF) {
  //
    if (yy.length==1) {
       str = w2ui.myGrid.getCellValue(yy[0],1);
       }
    else
       {
       for(i=0; i < yy.length; i++){
        str = str + w2ui.myGrid.getCellValue(yy[i] - 1,1);
        if (i!=(yy.length - 1)) {str=str+",";}
       }
      }
 alert("stai cancellando ... " + str);

 }
}
</script>
<style type="text/css">
<!--
.P1024 {
 width:1024px;
 height:768px;
 margin:auto;
 background-color: #333333;
}
body {
    background-color: #cccccc;
    background-image:url('images/mercatini_usato.jpg');
    background-repeat:no-repeat;
    background-position:center;
}
.btn {
 width:80px;
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
 background-color: #CCeeee;
}
.text {
 width:250px;
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  14px;
 color: blue;
}
.tabe {
 width:200px;
 font-family: Verdana, Arial, Helvetica, sans-serif;
 font-size:  14px;
 color: blue;
 cursor:pointer;
}
.box1 {
 width:100%;
 height: 100%;
 border-radius: 6px; 
 -moz-border-radius: 6px; /* firefox */
 -webkit-border-radius: 6px; /* safari, chrome */
  background-color: #cccccc;
}
.box1:hover {
 background-color: #CCeeee;
 color: blue;
}
-->
</style>
</head>


</head>

<body>
<div class="P1024" >

<!-- #include file="header.asp" -->

<form name="<%=formname%>" id="<%=formname%>" action="" method="post"  >
<p>
  <table width="100%" border="0" cellspacing="3">
      <td>&nbsp;</td>
      <td style="width:160px"><img src="images/btn/btn1_ricerca.jpg" alt="" /> </td>
      <td><input type="text" id="ricerca" name="ricerca" value="" style="width:250px" class="text" <%=crsx%>  onkeydown="EnterKey(event)" /></td>
      <td><input type="button" id="btSearch" value="esegui" onclick="Ricerca(ricerca.value)" class="btn" /></td>
      <td style="color:white">pagina
       <select id="pagina" name="pagina" style="background-color:white;color:blue;" onchange="ReloadPage()">
        <%
            for p=1 to npage
             if p=pagina then
               sel=" selected "
             else
               sel = ""
             end if    
            %>
                <option value=<%=p %> <%=sel %> ><%=p %></option>
            <%
            next
            %>
       </select>
       &nbsp; n.record per pagina=<%=numrecpage%> n.pagine=<%=npage%> nrecord=<%=ctrRecord%>
       </td>
  </table>

</p>

<div id="myGrid" style="height: 540px"></div>
<p>
  <table width="100%" border="0" cellspacing="5" cellpadding="2" style="color:white;" >
  <tr>
     <td width="25%" onclick="NuovoCliente()">&nbsp;

     </td>
     <td width="25%" onclick="DeleteCliente()">
      <img src="icons/TRASH02.ICO" alt="Cancella il cliente selezionato..." width="48" height="48"
         onmouseover="this.src='icons/TRASH04.ICO'" onmouseout="this.src='icons/TRASH02.ICO'" onclick="RimuoviClienti()" />Cancella un cliente...
     </td>
     <td width="25%" onclick="StampaMandato()">
      <img src="icons/stampante.gif" alt="stampa il mandato cliente selezionato..." />stampa il mandato ...
     </td>
     <td width="25%" onclick="ReportCliente()">
      <img src="icons/notepad.gif" alt="stampa il mandato cliente selezionato..." />Report clienti...
     </td>
  </tr>
  </table>

 <div class="footer>">
  <input name="msg" type="text" id="msg" value="<%=msg%>" style="width:99.5%; " readonly="true" />
 </div>
</p>
</form>
</div>


<script type="text/javascript">
$(function () {
  $('#myGrid').w2grid({
    name   : 'myGrid',
       show: {
            selectColumn: true
        },
    method: 'GET', // need this to avoid 412 error on Safari
    url:  'clientiJSON.asp',
    columns: [
        { field: 'recid', caption: '#', size: '50px', sortable: true, attr: 'align=center' },
        { field: 'codcli', caption: '<center><b>cod.anagr.</b></center>', size: '80px',attr: 'align=center', sortable: true, resizable: true },
        { field: 'ragsoc', caption: '<center><b>Ragione sociale</b></center>', size: '150px', sortable: true, resizable: true },
        { field: 'datare', caption: '<center><b>Data iscrizione</b></center>', size: '100px', attr: 'align=center', resizable: true },
        { field: 'indiri', caption: '<center><b>Domicilio</b></center>', size: '200px', resizable: true },
        { field: 'citta', caption: '<center><b>città</b>', size: '120px', sortable: true, resizable: true },
        { field: 'telef', caption: '<center><b>Telefono</b></center>', size: '120px', resizable: true },
        { field: 'email', caption: '<center><b>Email</b></center>', size: '200px', sortable: true, resizable: true },
        { field: 'notecl', caption: '<center><b>Note cliente</b></center>', size: '200px', resizable: true },
   ],
    onDblClick: function(event) {
        var grid = this;
        event.onComplete = function () {
        var yy=grid.getSelection() - 1;
        var xx=grid.getCellValue(yy, 1);
//    var yy = w2ui.myGrid.getSelection() - 1;
//    var xx = w2ui.myGrid.getCellValue(yy, 1);
       if (xx!="")
          OpenPage('cliente','codiceanagrafico='+xx);
        }
       }
  });
  w2ui['myGrid'].load('clientiJSON.asp');
});

</script>
</body>
</html>
<%
function LeggiConfigurazione()
dim pk
pk=1
SQL="select * From configurazione where pk = " & pk
RS.Open SQL, conn1,3,3
 numrecpage = RS.Fields("numrecpage")
RS.close
LeggiConfigurazione=numrecpage
end function

function ContatoreRecord(Tabella,condizione)
SQL="select count(*) as ctr1  from " & Tabella&" where "&condizione
RS.Open SQL, conn2, 3, 3
if RS.EOF then
  ContatoreRecord = 0
else
  ContatoreRecord = RS.Fields("Ctr1")
end if
RS.Close
end function
%>
<!-- #include file="Variabili.asp" -->