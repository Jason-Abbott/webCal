<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 8/25/98
-->

<%
Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

' query for event information and context name

query = "SELECT * FROM cal_events E INNER JOIN cal_context C" _
	& " ON (E.event_context = C.id) " _
	& " WHERE (event_id)=" & Request.QueryString("event_id")
Set rs = db.Execute(query)
%>

<!--#include virtual="/header_start.inc"-->
	<%=rs("name")%> Calendar:
	<%=MonthName(Month(rs("event_start"))) & " "%>
	<%=Day(rs("event_start")) & ", "%>
	<%=Year(rs("event_start"))%>

	</td>
	<td align=right>
	
<!-- buttons -->

<% if Session("user") <> "guest" then %>
<a href="cal_edit.asp?action=update&event_id=<%=rs("event_id")%>">
<img src="/graphics/button_edit.gif" border=0>
</a>

<a href="cal_del.asp?event_id=<%=rs("event_id")%>">
<img src="/graphics/button_del.gif" border=0>
</a>
<% end if %>

<!-- end buttons -->

<!--#include virtual="/header_end.inc"-->

<font size=5><%=rs("event_title")%></font><br>
From <%= Right(rs("event_start"),11)%> to <%=Right(rs("event_end"),11)%>
<hr size="1" color="#000000">
<%=rs("event_description")%>

<% if Session("user") <> "guest" then %>
<hr size="1" color="#000000">
<font face="arial" size=2>Updated on <b><%=rs("update_time")%></b> by <b><%=rs("update_machine")%></b></font>
<%
end if
rs.Close
db.Close
%>

<!--#include virtual="/footer.inc"-->