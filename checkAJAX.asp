<% 
'
' i miei pgm AJAX rispondono con parametri di risposta
' codice errore  (se = 0 è ok oppure 9 se in errore
' + il messaggio di risposta
' + n. campi in risposta
' + value per ognuno dei campi separati da punto e virgola
'
If (Session("Authenticated") = 0) Then
 Response.Write("9;SESSIONE SCADUTA - ESEGUIRE NUOVAMENTE IL LOGIN...")
 Response.end()
End If
%>