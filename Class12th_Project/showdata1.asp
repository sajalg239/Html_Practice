<HTML>
<HEAD>
<% DIM OBJCON, STRCON, OBJRS

   SET OBJCON=SERVER.CREATEOBJECT("ADODB.CONNECTION")
  STRCON= "PROVIDER=MICROSOFT.JET.OLEDB.4.0 ;DATA SOURCE="&_
                     "F:\SAJAL\SCHOOL DATA\DATA1\XII.mdb"
   OBJCON.OPEN STRCON
       SET OBJRS = SERVER.CREATEOBJECT("ADODB.RECORDSET")
       OBJRS.OPEN "XIIA", OBJCON%>
<BODY BGCOLOR=BLUE>
  <FONT SIZE=24>STUDENT TABLE</FONT>
   <TABLE BORDER=2>
   <% DO WHILE OBJRS.EOF=FALSE %>                                                           

<TR>
 <TD><%=OBJRS("Rollno")%>
 <TD><%=OBJRS("name")%>
 <TD><%=OBJRS("address")%>
 <td><%=OBJRS("fname")%>
 <TD><%=OBJRS("p_no")%>
</TR>

<% OBJRS.MOVENEXT

   LOOP
OBJRS.CLOSE
 OBJCON.CLOSE%>
</TABLE>
</BODY>
</HTML>

                                                                               