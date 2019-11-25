<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
for i=1 to Request.QueryString("filename").Count 
  filename =  Request.QueryString("filename")(i) 
next
%>
<!--#include file="nav_web.asp" -->
<html>
<head>
<title>Visualizza File pdf in un div...</title>
<script language="javascript" type="text/javascript" >
function ReturnBack() {
 self.location="<%=PreviousPage%>";
}
</script>   
<meta http-equiv="Content-Type" content="application/pdf; charset=iso-8859-1" />    
</head>
<body >
<input type="button"  onclick="ReturnBack()" value=" back" />
<div id="pdf">
        <object width="1024" height="768" type="application/pdf" data="<%=filename %>?#zoom=100&scrollbar=1&toolbar=1&navpanes=0"
            id="pdf_content">
<p>Verificare se è installata la versione di Acrobat Reader richiesta per visualizzare i file PDF...</p>
        </object>
    </div>
</body>
</html>
<%
response.flush()
%>