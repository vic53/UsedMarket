<!-- #include file="adovbs.inc" -->
<!-- #include file="Connessioni.asp" -->
<%
Dim CtrRec,tabella,crlf,nominativo,filtro,myArray,p1,p2,limite,codiceanagrafico,npage,pagina,ctrRecord
crlf=CHR(13)&CHR(10)
' W "apice="& ASC("'") &"<br>"
filtro=""
for i=1 to Request.QueryString("nominativo").Count
 nominativo =  trim(Request.QueryString("nominativo")(i))
 if nominativo <> "" then
  MyArray = Split(nominativo, " ", -1, 1)
  if ubound(MyArray) > 0 then
     p1=MyArray(0)
     p2=MyArray(1)
     filtro=" where cognome like '%"&p1&"%' and nome like '%"&p2&"%' order by COGNOME ASC, NOME ASC"
  else
     p1=MyArray(0)
     filtro=" where cognome like '%"&p1&"%' or nome like '%"&p1&"%' order by COGNOME ASC, NOME ASC"
  end if
 end if
next
'
' N.record visualizzati x pagina
'
numrecpage=LeggiConfigurazione()
pagina = 1 'default di partenza
ctrRecord=ContatoreRecord("Clienti"," ATTIVO=1 " )
' se ricicla cambiando pagina questa è stata impostata dall'utente
pagina=1 'default pagina
'pagina di ricerca modificata?
for i=1 to Request.QueryString("pagina").Count
  pagina =  cint(Request.QueryString("pagina")(i))
next

'
' campo unico numerico identifica codiceanagrafico
'
codiceanagrafico=-1
for i=1 to Request.QueryString("codiceanagrafico").Count
 codiceanagrafico = Request.QueryString("codiceanagrafico")(i)
 filtro=" where codcli="&codiceanagrafico
next

SQL="select * from CLIENTI " & filtro

' conteggio delle pagine di record
RS.PageSize=numrecpage
RS.Open SQL, conn2, 3, 3
if not RS.EOF then
RS.AbsolutePage = pagina
'
' N.record della pagina attuale
'
 ctrRecord=RS.RecordCount
 npage=RS.PageCount
 if nominativo="" or codiceanagrafico= -1 then
  if pagina=1 then
   if ctrRecord < numrecpage  then
     ctrRec=Rs.RecordCount
   else
     ctrrec=numrecpage
   end if
'  W "<br>1:"& ctrrec &"<br>"
  else
   if pagina < npage then
      ctrRec = numrecpage
 '  W "2:"& ctrrec  &"<br>"
   else
      ctrRec=ctrRecord - (npage - 1) * numrecpage
 '  W "3:"& ctrrec &"<br>"
   end if
  end if
 else
  ctrRec=RS.RecordCount
  pagina=1
 end if

 ctr=0
 W "{"
 W """total"":" & CtrRec & ","
 W "records: ["
do while not(RS.EOF or (RS.AbsolutePage <> pagina))
   ctr=ctr+1
   W "{recid:" & ctr & ","
   W "codcli:" & CHR(39) & RS("CODCLI").value & CHR(39) & ","
   W "ragsoc:" & CHR(39) & trim(RS("COGNOME").value) & " " & trim(RS("NOME").value) & CHR(39)& ","
   W "datare:" & CHR(39) & fmtdate(RS("DATAISCRIZIONE").value) & CHR(39) & ","
   W "indiri:" & CHR(39) & trim(RS("DOMICILIO").value) & CHR(39) & ","
   W "citta:" & CHR(39) & trim(RS("CITTA").value) & CHR(39) & ","
   W "telef:" & CHR(39) & trim(RS("TELEFONO").value) & CHR(39) & ","
   W "email:" & CHR(39) & trim(RS("EMAIL").value) & CHR(39) & ","
   W "notecl:" & CHR(39) & trim(RS("NOTECLIENTE").value) & CHR(39)
   W "}"
  RS.movenext
  if (RS.EOF or (RS.AbsolutePage <> pagina)) then
   W " "
  else
   W ","
  end if
loop
 W "]"
 W "}"
end if
RS.close

sub W(msg)
 response.write(msg)
end sub
function fmtDate(dt)
if isDate(dt) then
  fmtDate=right("00"&day(dt),2)&"/"&right("00"&month(dt),2)&"/"&right("0000"&year(dt),4)
else
  fmtDate="-"
end if
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

function LeggiConfigurazione()
dim pk
pk=1
SQL="select * From configurazione where pk = " & pk
RS.Open SQL, conn1,3,3
 numrecpage = RS.Fields("numrecpage")
RS.close
LeggiConfigurazione=numrecpage
end function
%>

<!-- #include file="Disconnessioni.asp" -->