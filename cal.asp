<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 6/12/98

dim event_context, FirstDay, DaysInMonth, TheDay, cal, col
dim row, rows, context_list, events(31), context_name
dim m, y, mNext, mPrev, yNext, yPrev, mLoop, yLoop

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' determine how this page was called

if Request.QueryString("event_context") <> "" then
	event_context = Request.QueryString("event_context")
	m = CDbl(Request.QueryString("month"))
	y = CDbl(Request.QueryString("year"))
elseif Request.Form("event_context") = "(New)" then
	response.redirect "cal_create.asp"
elseif Request.Form("month") <> "" then
	event_context = Request.Form("event_context")
	m = CDbl(Request.Form("month"))
	y = CDbl(Request.Form("year"))
else
	event_context = 10
end if

if m = 0 OR m = "" then
	m = Month(now)
	y = Year(now)
end if
%>

<!--#include file="cal_nav.inc"-->

<%
' ---------------------------------------------------------
' read in data
' ---------------------------------------------------------

Set db = Server.CreateObject("ADODB.Connection")
DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("webCal2.mdb")
db.Open DSN

' start new recordset
' create list of contexts, saving current context name

query = "SELECT * FROM cal_context"
if Session("user") = "guest" then
	query = query & " WHERE private = 0" 
end if
query = query & " ORDER BY name"

Set rs = db.Execute(query)
do while not rs.EOF
	context_list = context_list & "<option value='" & rs("id") & "'"
	if rs("id") = CDbl(event_context) then
		context_list = context_list & " selected"
		context_name = rs("name")
	end if
	context_list = context_list & ">" & rs("name")
	if rs("private") = "True" then context_list = context_list & "*"
	context_list = context_list & VbCrLf
	rs.movenext
loop
rs.Close

' add option to create new calendar if not guest

if Session("user") <> "guest" then context_list = context_list & "<option>(New)"

' create recordset of events
' read events into day-based array

query = "SELECT * FROM cal_events WHERE event_context=" & event_context _
	& " AND event_start BETWEEN #" & m & "/1/" & y & "#" _
	& " AND #" & mnext & "/1/" & ynext & "# ORDER BY event_start"
Set rs = db.Execute(query)
do while not rs.EOF
	events(Day(rs("event_start"))) = events(Day(rs("event_start"))) _
		& "<br>" & Left(Right(rs("event_start"),11),5) _ 
		& LCase(Right(rs("event_start"),2)) & ": " _
		& "<a href='cal_detail.asp?event_id=" & rs("event_id") _
		& "'>" & rs("event_title") & "</a>"
	rs.movenext
loop
rs.Close
db.Close

' ---------------------------------------------------------
' generate calendar body
' ---------------------------------------------------------

FirstDay = WeekDay(Dateserial(y,m,1))

' use Dateserial to find last day of this month by taking first
' day of next month and

DaysInMonth = Day(Dateserial(y,m+1,1)-1)
TheDay = 0

' calculate total rows by finding total number of weeks after
' top row

rows = Fix((DaysInMonth - (8 - FirstDay))/7) + 2

for row = 1 to rows 
	cal = cal & "<tr>"
	for col = 1 to 7 
		if row = 1 AND col = FirstDay then TheDay = 1
		if TheDay > 0 AND TheDay <= DaysInMonth then
			cal = cal & "<td height='45' valign='top' bgcolor='#"			

' color weekdays white, weekends grey and today yellow

			if y & m & TheDay = Year(now) & Month(now) & Day(now) then
				cal = cal & "ffffbb'>"
			elseif col = 1 OR col = 7 then
				cal = cal & "eeeeee'>"
			else
				cal = cal & "ffffff'>"
			end if
			
			cal = cal & "<font face='arial'><font size=2>"
			if Session("user") <> "guest" then
				cal = cal & "<a href='cal_edit.asp?" _
				& "action=add" _
				& "&event_context=" & event_context _
				& "&year=" & y _
				& "&month=" & m _
				& "&day=" & TheDay _
				& "'>"
			end if
			cal = cal & TheDay
			if Session("user") <> "guest" then
				cal = cal & "</a>"
			end if
			cal = cal & "</font><font size=1>" _
				& events(TheDay) & "</font>"
			TheDay = TheDay + 1
		else
			cal = cal & "<td>"
		end if
		cal = cal & "</td>"
	next
next
%>

<!--include file="header_start.inc"-->
<%=context_name%> Calendar: <%=MonthName(m) & " " & y%>

	</td>
	<td align="right">
	<font size=1 face="arial">	
<a href="cal_print.asp?event_context=<%=event_context%>&month=<%=m%>&year=<%=y%>" target="_top">
<img src="/graphics/button_print.gif" alt="View Printable Calendar" border=0></a>
	</font>
<!--include file="header_end.inc"-->

<center>

<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="1">
<tr>

<%
' ---------------------------------------------------------
' generate calendar header
' ---------------------------------------------------------

' previous month link

response.write "<td><font face='arial' size=2><b><a href='cal.asp?" _
	& "event_context=" & event_context _
	& "&month=" & mPrev _
	& "&year=" & yPrev & "'>&lt;" _
	& MonthName(mPrev) & "</a></b></font></td>"

' navigation form
	
response.write "<form method='post' action='cal.asp'>" _
	& "<td align='center' colspan=5>" _
	& "<select name = 'month'>" & VbCrLf

' month list

for mLoop = 1 to 12
	response.write "<option value='" & mLoop & "'"
	if mLoop = m then response.write " selected"
	response.write ">" & MonthName(mLoop) & VbCrLf
next
response.write "</select><select name = 'year'>"

' year list

for yLoop = year(Now) - 10 to year(Now) + 10
	response.write "<option"
	if yLoop = y then response.write " selected"
	response.write ">" & yLoop & VbCrLf
next

' context list

response.write "</select><select name='event_context'>" _
	& context_list & "</select>" _
	& "<input type='image' src='/graphics/button_go.gif' border=0>" _
	& "</td></form>"

' next month link

response.write "<td align='right'><font face='arial' size=2><b>" _
	& "<a href='cal.asp?" _
	& "event_context=" & event_context _
	& "&month=" & mNext _
	& "&year=" & yNext & "'>" _
	& MonthName(mNext) & "&gt;</a></b></font></td>"

' day headings

response.write "<tr>" & VbCrLf
for col = 1 to 7
	response.write "<td width='14.285713%' align='center'" _
		& " color='#c0c0c0'><font face='arial' size='1'>" _
		& WeekDayName(col,0) & "</font><hr size=1 color='#000000'></td>"
next

' insert calendar body

response.write cal & "</table></center>"

if Session("user") <> "guest" then
	response.write "<font face='arial' size=2>" _
		& "*These calendars are only visible to Boise Center users"
end if
%>

<!--include virtual="/footer.inc"-->