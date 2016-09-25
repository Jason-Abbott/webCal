<!--#include file="cal_nav.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 6/3/98

if Request.Form("event_context") = "(New)" then
	response.redirect "cal_create.asp"
end if

dim event_context,FirstDay,DaysInMonth,TheDay,cal,col,row,rows,event_list

' Determine values for the month and year to display
' Form values must be converted from strings to numerics

if Request.Form("month") = "" then
	call cal_nav(Month(now),Year(now))
	event_context = 10
else
	call cal_nav(CDbl(Request.Form("month")),CDbl(Request.Form("year")))
	event_context = Request.Form("event_context")
end if

' ---------------------------------------------------------
' read in data
' ---------------------------------------------------------

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

' create recordset of events

query = "SELECT event_start FROM cal_events WHERE event_context=" & event_context _
	& " AND event_start BETWEEN #" & m & "/1/" & y & "#" _
	& " AND #" & mnext & "/1/" & ynext & "#"
Set rs = db.Execute(query)

' create a comma delimited string containing days with events

do while not rs.EOF
	events = events & "," & Day(rs("event_start")) & ","
	rs.movenext
loop

' close calendar recordset

rs.Close

' start new recordset
' create list of contexts, selecting current

query = "SELECT * FROM cal_context"
if Session("user") = "guest" then
	query = query & " WHERE private = 0" 
end if
query = query & " ORDER BY name"

Set rs = db.Execute(query)
do while not rs.EOF
	event_list = event_list & "<option value='" & rs("id") & "'"
	if rs("id") = CDbl(event_context) then
		event_list = event_list & " selected"
		context_name = rs("name")
	end if
	event_list = event_list & ">" & rs("name") & VbCrLf
	rs.movenext
loop

rs.Close
db.Close

' add option to create new calendar if not guest

if Session("user") <> "guest" then event_list = event_list & "<option>(New)"

' ---------------------------------------------------------
' begin generating calendar layout
' ---------------------------------------------------------

FirstDay = WeekDay(Dateserial(y,m,1))

' use Dateserial to find last day of this month by taking first
' day of next month and subtracting one

DaysInMonth = Day(Dateserial(y,m+1,1)-1)
TheDay = 0

' calculate total rows by finding total number of weeks after
' top row

rows = Fix((DaysInMonth - (8 - FirstDay))/7) + 2

for row = 1 to rows 
	cal = cal & "<tr>"
	for col = 1 to 7 
		active = 0
		if row = 1 AND col = FirstDay then TheDay = 1
		if TheDay > 0 AND TheDay <= DaysInMonth then

' Bracket the day if it is today

			if y & m & TheDay = Year(now) & Month(now) & Day(now) then
				today = "[<b>" & TheDay & "</b>]"
			else
				today = TheDay
			end if

			cal = cal & "<td align=center bgcolor='#"			

' Highlight the day if it has events or shade it if it's a weekend

			if InStr(1,events,"," & TheDay & ",",1) <> 0 then
				cal = cal & "ffffbb'><a href='cal_list.asp?"
			elseif col = 1 OR col = 7 then
				cal = cal & "e0e0e0'><a href='cal_edit.asp?action=add&"
			else
				cal = cal & "ffffff'><a href='cal_edit.asp?action=add&"
			end if
		
			cal = cal & "event_context=" & event_context _
				& "&context_name=" & context_name _
				& "&year=" & y _
				& "&month=" & m _
				& "&day=" & TheDay _
				& "'><font size=2 face='arial'>" & today & "</font></a>"
			TheDay = TheDay + 1
		else
			cal = cal & "<td>"
		end if
		cal = cal & "</td>"
	next
next
%>

<!--#include virtual="/header_start.inc"-->
	<%=context_name%> Calendar: <%=MonthName(m) & " " & y%>
<!--#include virtual="/header_end.inc"-->

<center>

<table border="0" cellspacing="1" cellpadding="1" width=350>
<tr>
	<form method="post" action="cal.asp">
	<td colspan=7 align=center>

<!-- Create drop-downs by looping through months and 20 years -->

	<select name = "month">
<% for mloop = 1 to 12 %>
	<option value="<%=mloop%>"<% if mloop = m then %>selected<%end if%>><%=MonthName(mloop)%>
<% next %>
	</select>

	<select name = "year">
<% for yloop = year(Now) - 10 to year(Now) + 10 %>
	<option <% if yloop = y then %>selected<%end if%>><%=yloop%>
<% next %>
	</select>

<!-- Display list of calendars with option to create new -->

	<select name="event_context">
	<%=event_list%>
	</select>
	<input type=image src="/graphics/button_go.gif" border=0>
	</td></form>
	
<!-- Create day headings by looping through day numbers -->
	
<tr>
<% for d = 1 to 7 %>
	<td width="14.285713%" align="center" color="#c0c0c0">
	<font face="Arial" size="1"><%= WeekDayName(d,1) %></font></td>
<% next %>

<!-- insert the calendar here -->

	<%=cal%>
</table>
<p>
<table border=0 cellpadding=0 cellspacing=0 width=350>
<tr>
	<form action="cal.asp" method="post">
	<td align=left valign=top>
	<input type="hidden" name="month" value="<%= mprev %>">
	<input type="hidden" name="year" value="<%= yprev %>">
	<input type="hidden" name="event_context" value="<%= event_context %>">
	<input type=image src="/graphics/button_left.gif" border=0>
	</td></form>

	<td align=center valign=top>
<%
' if there are events display a button to swith to list	

if events <> "" then
	response.write "<a href='cal_list.asp?" _
		& "event_context=" & event_context _
		& "&year=" & y _
		& "&month=" & m _
		& "&day=0'>" _
		& "<img src='/graphics/button_event.gif' border=0></a>"
end if
%>
	</td>

	<form action="cal.asp" method="post">
	<td align=right valign=top>
	<input type="hidden" name="month" value="<%= mnext %>">
	<input type="hidden" name="year" value="<%= ynext %>">
	<input type="hidden" name="event_context" value="<%= event_context %>">
	<input type=image src="/graphics/button_right.gif" border=0>
	</form>
	</td>
</table>
</center>
<p>
<font face="arial"><b>Tips:</b></font><br>
Days with events are highlighted.  Weekends are grey.  Use the arrows to move backward and forward to new months or select a month and calendar from the list at top and press "Go to."  Boise Center employees may click on a day to add new events.  The current date is bracketed.  Switch to events view to see a list of the month's events (if the month has no scheduled events this option will not be available).

<!--#include virtual="/footer.inc"-->