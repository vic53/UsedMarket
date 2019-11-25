<%
Dim vet, NAVPAGE
menu="menu.asp"
Session("NAVPAGE") = Session("NAVPAGE") & ThisProgram & ";"
NAVPAGE = Session("NAVPAGE")
vet=split(NAVPAGE,";")
for i=0 to ubound(vet) 
  if vet(i)=ThisProgram then
     previouspage=vet(i - 1) 
     k=i
    exit for
  end if
next
Session("NAVPAGE")=""
for i=0 to k
Session("NAVPAGE") = Session("NAVPAGE") & vet(i) & ";"
next
%>