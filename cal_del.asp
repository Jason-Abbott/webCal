<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 6/10/98
-->

<%
' query for event information and context name

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"
query = "SELECT * FROM cal_events E INNER JOIN cal_context C" _
	& " ON (E.event_context = C.id) " _
	& " WHERE (event_id)=" & Request.QueryString("event_id")
Set rs = db.Execute(query)
%>

<!--#include virtual="/header_start.inc"-->
Delete "<%=rs("event_title")%>" from the <%=rs("name")%> calendar?
<!--#include virtual="/header_end.inc"-->


<form action="cal_deleted.asp" method="post">
<center>
<font color="#ff0000" size=6>This action is not reversable!</font>
<p>
<input type="hidden" name="event_id" value="<%=Request.QueryString("event_id")%>">
<input type="hidden" name="month" value="<%=Month(rs("event_start"))%>">
<input type="hidden" name="year" value="<%=Year(rs("event_start"))%>">
<input type="hidden" name="event_context" value="<%=rs("event_context")%>">
<input type="submit" value="delete">
</form>
<%
rs.Close
db.Close
%>
</center>

<!--#include virtual="/footer.inc"-->