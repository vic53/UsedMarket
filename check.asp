<% 
If Session("Authenticated") = 0 Then
 Session("LASTMSG")="CONNESSIONE IN TIME OUT O START PROGRAM SENZA LOGIN ESEGUITO..." 
 Response.Redirect ("Login.asp")
End If
%>