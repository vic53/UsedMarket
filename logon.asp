<!-- #include file="adovbs.inc" -->
<!-- #include file="Connessioni.asp" -->
<% 
Dim  username, password
username=request.form("username")
password=request.form("pw1")
Session("USERNAME")=username
Session("PASSWORD")=password
Session("NAVPAGE")="menu.asp;"
Session("SYSMSG")="Il sistema di messaggistica tra stazioni è attivo..."
'
SQL="select * from Operatori where username='" & trim(username) & "'"
RS.Open SQL, conn1, 3, 3 
if not RS.EOF then
 if username=RS.Fields("username") then
   if password=RS.Fields("password") then
    if RS.Fields("Attivo") <> 0 then
      Session("Authenticated")="1"
      Session("USERNAME")=username
      Session("LASTMSG")="LOGIN ESEGUITO..."
      Session("TITOLOGESTIONE")="GESTIONE MERCATINO DELL'USATO"
      Session.Timeout = 1200
      Session("LOG")="YES" ' attivato file di log per la cancellazione di records
      response.redirect "menu.asp"
     else
      Session("LASTMSG")= username & ": UTENTE IN ATTESA DI AUTORIZZAZIONE DA UN AMMINISTRATORE..."
      response.redirect "login.asp"
     end if
   else
      Session("LASTMSG")="USERNAME O PASSWORD NON CORRETTI..."
      response.redirect "login.asp"       
   end if
 else
      Session("LASTMSG")="USERNAME O PASSWORD NON CORRETTI..."
      response.redirect "login.asp"       
 end if
else
  Session("LASTMSG")="USERNAME O PASSWORD NON CORRETTI..."
  response.redirect "Login.asp"       
end if
RS.close
%>
<!-- #include file="Disconnessioni.asp" -->