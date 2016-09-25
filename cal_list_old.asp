<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 6/3/98
-->

<!--#include file="cal_nav.inc"-->
<%
call cal_nav(CDbl(Request.QueryString("month")),CDbl(Request.QueryString("year")))

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

' lookup the context name

Set rs = db.Execute("SELECT * FROM cal_context WHERE id = " & Request.QueryString("event_context"))
context_name = rs("name")
rs.Close

' lookup all events in selected month

query = "SELECT * FROM cal_events WHERE" _
	& " event_context=" & Request.QueryString("event_context") _
	& " AND event_start BETWEEN #" & m _
	& "-1-" & y & "# AND #" _
	& mnext & "-1-" _
	& ynext & "#" _
	& " ORDER BY event_start"
	
Set rs = db.Execute(query)

' uncomment this line to debug SQL:
' response.write query
%>

<!--#include virtual="/header_start.inc"-->
	<%=context_name%> Calendar:
   <%=MonthName(m) & " " & y%>

	</td>
	<td align=right>
	
<!-- ADD BUTTON -->

<% if Session("user") <> "guest" then %>
<p>
<a href="cal_edit.asp?action=add&event_context=<%=Request.QueryString("event_context")%>&context_name=<%=context_name%>&year=<%=y%>&month=<%=m%>">
<img src="/graphics/button_add.gif" border=0>
</a>
<% end if %>

<!-- END BUTTON -->

<!--#include virtual="/header_end.inc"-->

<center>
<table cellpadding=4 cellspacing=0 border=0>
<tr>
	<td align=center <%=light%>>Day</td>
	<td align=center <%=light%>>Start</td>
	<td align=center <%=light%>>Event</td>
	<td align=center <%=light%>>End</td>
	
<%
Do While Not rs.EOF
	response.write "<tr><td"
	
' highlight the event if it occurs on the day clicked

	if CDbl(Request.QueryString("day")) = Day(rs("event_start")) then
		response.write " bgcolor='#ffffbb'"
	end if
	response.write " align=right><font face='arial' size=2>" _
		& Day(rs("event_start")) & "</font></td>" _
		& "<td align=right " & light & ">" _
		& "<font face='arial' size=2>" _
		& Left(Right(rs("event_start"),11),5) _
		& " " & Right(rs("event_start"),2) _
		& "</font></td>" _
		& "<td"
	if CDbl(Request.QueryString("day")) = Day(rs("event_start")) then
		response.write " bgcolor='#ffffbb'"
	end if
	response.write "><a href='cal_detail.asp?" _
		& "event_id=" & rs("event_id") _
		& "&context_name=" & context_name & "'>" _
		& rs("event_title") & "</a></td>" _
		& "<td align=right " & light & ">" _
		& "<font face='arial' size=2>" _
		& Left(Right(rs("event_end"),11),5) _
		& " " & Right(rs("event_end"),2) _
		& "</font></td>"
	rs.MoveNext
Loop
rs.Close
db.Close
%>

</table>
<p>
<table border=0 cellpadding=0 cellspacing=0 width=350>
<tr>
	<td align=left valign=top>
	<a href="cal_list.asp?event_context=<%=Request.QueryString("event_context")%>
		&year=<%= yprev %>
		&month=<%= mprev %>
		&day=0">
	<img src="/graphics/button_left.gif" border=0></a></td>

	<form action="cal.asp" method="post">
	<td align=center valign=top>
	<input type="hidden" name="month" value="<%=m%>">
	<input type="hidden" name="year" value="<%=y%>">
	<input type="hidden" name="event_context" value="<%= Request.QueryString("event_context") %>">
	<input type=image src="/graphics/button_cal.gif" border=0></td>
	</td></form>

	<td align=right valign=top>
	<a href="cal_list.asp?event_context=<%=Request.QueryString("event_context")%>
		&year=<%= ynext %>
		&month=<%= mnext %>
		&day=0">
	<img src="/graphics/button_right.gif" border=0></a></td>
</table>
</center>

<!--#include virtual="/footer.inc"-->