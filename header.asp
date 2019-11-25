<style type="text/css">
<!--
.Stile1 {
color: #00FFFF;
font:Verdana, Arial, Helvetica, sans-serif;
}
.Stile2 {
    color: #66FFFF;
    font-size: 24px;
}
-->
</style>
<table style="width:100%; background-color: #000000; ">
<tr class="Stile1">
<%if flagHeader(1,1) = "Y" then%>
<td><img src="<%=flagHeader(1,2) %>" class="Stile1" title="<%=flagHeader(1,4)%>" onclick="<%=flagHeader(1,3)%>"
   style="cursor:pointer;" <%=flagHeader(1,5) %>/></td>
<%end if
if flagHeader(2,1) = "Y" then%>
<td><img src="<%=flagHeader(2,2) %>" title="<%=flagHeader(2,4)%>" onclick="<%=flagHeader(2,3) %>"
      style="cursor:pointer;" <%=flagHeader(2,5) %>/></td>
<%end if
if flagHeader(4,1) = "Y" then%>
<td><img src="<%=flagHeader(4,2) %>" title="<%=flagHeader(4,4)%>" onclick="<%=flagHeader(4,3) %>"
      style="cursor:pointer;" <%=flagHeader(4,5) %>/></td>
<%end if%>
<td style="width:100%; vertical-align: bottom; color:#66FFCC;"><span class="Stile2"><%=pgmtitle %></span></td>
<td  ><img src="images/logoform.jpg"></td>
<td style="width:100px;color:white; vertical-align: bottom;"><%=Session("USERNAME")%>&nbsp;</td>
<td style="width:24px; height:24px;" align="right">
    <img src="icons/userH.ico" alt="login on timeout" title="login/logout..."  onclick="UserLogin()"
      onmouseover="this.style.cursor='pointer'" height="32" width="32" /></td>
<td style="width:24px; height:24px;" align="right">
    <img src="icons/help.gif" alt="Help page" title="help per la pagina in corso..."  onclick="CallHelp('<%=formname%>.html')"
      onmouseover="this.style.cursor='pointer'" /></td>
<td >
    <img src="<%=flagHeader(3,2) %>" width="32" height="32" alt="<%=flagHeader(3,4) %>"  title="<%=flagHeader(3,4) %>"  onclick="<%=flagHeader(3,3)%>"
      onmouseover="this.style.cursor='pointer'" /></td>
</tr>
</table>

