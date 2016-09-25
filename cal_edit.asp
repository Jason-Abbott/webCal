<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 8/25/98

if Session("User") = "guest" then response.redirect "cal_denied.asp"

dim form, event_day, event_month, event_year, context_name
dim event_start, event_start_apm, event_end, event_end_apm, event_list

' open database

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

' set variables according to how the form was called

if Request.QueryString("action") = "add" then
	description = "Addition"
	if Request.QueryString("day") = 0 then
		event_day = 1
	else
		event_day = CDbl(Request.QueryString("day"))
	end if
	event_month = CDbl(Request.QueryString("month"))
	event_year = CDbl(Request.QueryString("year"))
	event_start = "8:00"
	event_end = "5:00"
	event_start_apm = "AM"
	event_end_apm = "PM"
	event_title = ""
	event_description = ""
	event_id = ""
   event_context = Request.QueryString("event_context")
else
	description = "Update"
	
	query = "SELECT * FROM cal_events" _
		& " WHERE (event_id)=" & Request.QueryString("event_id")
	Set rsEvents = db.Execute(query)
		
	event_day = Day(rsEvents("event_start"))
	event_month = Month(rsEvents("event_start"))
	event_year = Year(rsEvents("event_start"))
	if Hour(rsEvents("event_start")) > 12 then
		event_start = Hour(rsEvents("event_start")) - 12 _
			& ":" & Left(Right(rsEvents("event_start"),8),2)
	else
		event_start = Hour(rsEvents("event_start")) _
			& ":" & Left(Right(rsEvents("event_start"),8),2)
	end if
	event_start_apm = Right(rsEvents("event_start"),2)
	event_end_apm = Right(rsEvents("event_end"),2)
	if Hour(rsEvents("event_end")) > 12 then
		event_end = Hour(rsEvents("event_end")) - 12 _
			& ":" & Left(Right(rsEvents("event_end"),8),2)
	else
		event_end = Hour(rsEvents("event_end")) _
			& ":" & Left(Right(rsEvents("event_end"),8),2)
	end if
	event_title = rsEvents("event_title")
	event_description = rsEvents("event_description")
	event_id = rsEvents("event_id")
	event_context = rsEvents("event_context")

	rsEvents.Close
end if

' start new recordset
' create list of contexts, saving current context name

query = "SELECT * FROM cal_context ORDER BY name"
Set rsContext = db.Execute(query)
do while not rsContext.EOF
	event_list = event_list & "<option value='" & rsContext("id") & "'"
	if rsContext("id") = CDbl(event_context) then
		event_list = event_list & " selected"
		context_name = rsContext("name")
	end if
	event_list = event_list & ">" & rsContext("name") & VbCrLf
	rsContext.movenext
loop
rsContext.Close
db.Close
%>

<!--#include virtual="/header_start.inc"-->
	<%=context_name%> Calendar: Event <%=description%>
<!--#include virtual="/header_end.inc"-->

<html>
<body link="#E4C721" vlink="#E4C721" alink="#E4C721" bgcolor="#FFFFFF">
<center>
<table border=0 cellspacing=0 cellpadding=2>
<tr>
	<td align=center> </td>
<form method="post" action="cal_edited.asp">
<input type="hidden" name="action" value="<%=Request.QueryString("action")%>">
<input type="hidden" name="event_id" value="<%=event_id%>">

	<td align=center colspan=2>
	<select name="month">
<%
for m = 1 to 12
	response.write "<option value='" & m & "'"
	if m = event_month then response.write " selected"
	response.write ">" & MonthName(m)
next
%>
	</select>
	<select name="day">
<%
for d = 1 to 31
	response.write "<option"
	if d = event_day then response.write " selected"
	response.write ">" & d
next
%>
	</select>, 
	<select name="year">
<%
for y = event_year - 1 to event_year + 1
	response.write "<option"
 	if y = event_year then response.write " selected"
	response.write ">" & y
next
%>
	</select>
	</td>
<tr>
	<td align="right" <%=light%>>
	<font face="arial" size=2>Title</font></td>
	<td colspan=2>
	<input type="text" name="event_title" value="<%=event_title%>" size=45>
	</td>
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>Description</font></td>
	<td colspan=2>
	<textarea name="event_description" cols=50 rows=15 wrap="virtual"><%=event_description%></textarea>
	</td>
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>Start Time</font></td>
	<td colspan=2><font face="arial" size=2>
	<select name="event_start">
<%
for t = 1 to 12
	response.write "<option"
	if event_start = t & ":00" then response.write " selected"
	response.write ">" & t & ":00<option"
	if event_start = t & ":30" then response.write " selected"
	response.write ">" & t & ":30"
next
%>
	</select>
	<input type="radio" name="s_apm" value="am"
	<%if event_start_apm = "AM" then %> checked<%end if%>>AM
	<input type="radio" name="s_apm" value="pm"
	<%if event_start_apm = "PM" then %> checked<%end if%>>PM
	</font></td>
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>End Time</font></td>
	<td colspan=2><font face="arial" size=2>
	<select name="event_end">
<%
for t = 1 to 12
	response.write "<option"
	if event_end = t & ":00" then response.write " selected"
	response.write ">" & t & ":00<option"
	if event_end = t & ":30" then response.write " selected"
	response.write ">" & t & ":30"
next
%>
	</select>
	<input type="radio" name="e_apm" value="am"
	<%if event_end_apm = "AM" then %> checked<%end if%>>AM
	<input type="radio" name="e_apm" value="pm"
	<%if event_end_apm = "PM" then %> checked<%end if%>>PM
	</font></td>
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>Calendar</font></td>
	<td>
	<select name="event_context">
	<%=event_list%>
	</select>
	</td>
	<td align=right>
	<input type="submit" value="<%=Request.QueryString("action")%>"></td>
</table>
</form>
</center>

<!--#include virtual="/footer.inc"-->