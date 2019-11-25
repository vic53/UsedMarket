<%
'response.write "LOGOUT eseguito"
'Session.Abandon
 Session("Authenticated") = 0
 Session("NAVPAGE")=""
 Session("USERNAME")=""
 Session("LASTMSG")="DISCONNESSIONE ESEGUITA..."
 response.redirect("login.asp")
%>