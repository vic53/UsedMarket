<%
Public conn1,conn2,SQL,SQL1,SQL2,RS,RS1,RS2,RS3,conn1_string,conn2_string,archivi,servizi
Session.LCID=1040 
Session.Timeout = 120
servizi="database/Servizi.accdb"
archivi="database/archivi.accdb"
Set conn1 = Server.CreateObject("ADODB.Connection")
Set conn2 = Server.CreateObject("ADODB.Connection")
conn1_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath(servizi) & ";Persist Security Info=False;"
conn1.Open conn1_string
conn2_string="Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" &Server.MapPath(archivi) & ";Persist Security Info=False;"
conn2.Open conn2_string
Set RS = Server.CreateObject("ADODB.Recordset")
Set RS1 = Server.CreateObject("ADODB.Recordset")
Set RS2 = Server.CreateObject("ADODB.Recordset")
Set RS3 = Server.CreateObject("ADODB.Recordset")
Session.CodePage="1252"
Response.Charset="windows-1252"
Response.ContentType = "text/html"
%>