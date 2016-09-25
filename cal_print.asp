<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 6/11/98

dim event_context, FirstDay, DaysInMonth, TheDay, cal, col
dim row, rows, event_list, events(31), context_name
dim m, y, mNext, mPrev, yNext, yPrev, mLoop, yLoop

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

m = CDbl(Request.QueryString("month"))
y = CDbl(Request.QueryString("year"))
%>

<!--#include virtual="/cal/cal_nav.inc"-->

<%
' ---------------------------------------------------------
' read in data
' ---------------------------------------------------------

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

' create recordset of events
' read events into day-based array

query = "SELECT * FROM cal_events" _
	& " WHERE event_context=" & Request.QueryString("event_context") _
	& " AND event_start BETWEEN #" & m & "/1/" & y & "#" _
	& " AND #" & mnext & "/1/" & ynext & "# ORDER BY event_start"
Set rsEvents = db.Execute(query)
do while not rsEvents.EOF
	events(Day(rsEvents("event_start"))) = events(Day(rsEvents("event_start"))) _
		& "<br>" & Left(Right(rsEvents("event_start"),11),5) _ 
		& LCase(Right(rsEvents("event_start"),2)) & ": " & rsEvents("event_title")
	rsEvents.movenext
loop
rsEvents.Close

' create separate query to find context name
' must be separate since join with above could be empty

query = "SELECT * FROM cal_context" _
	& " WHERE (id)=" & Request.QueryString("event_context")
Set rsContext = db.Execute(query)
context_name = rsContext("name")
rsContext.Close
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
			cal = cal & "<td height='45' valign='top'>" _
				& "<font face='arial'><font size=2>" _
				& TheDay & "</font><font size=1>" _
				& events(TheDay) & "</font>"
			TheDay = TheDay + 1
		else
			cal = cal & "<td>&nbsp;"
		end if
		cal = cal & "</td>"
	next
next
%>

<html>
<title><%=context_name & ": " & MonthName(m) & " " & y%></title>
<body link="#443333" vlink="#443333" bgcolor="#FFFFFF">
<center>

<table width="100%" border="0">
<tr>
	<td width="50%">
<font size=6><%=MonthName(m) & " " & y%></font><hr size=1>
<font size=2 face="arial">
<b><%=context_name%></b> Calendar as of
<%=MonthName(Month(Now)) & " " & Day(Now) & ", " & Year(Now)%>
</font>
	</td>
	<td width="50%" align="right">
	<font face="arial" size=2>
	University of Idaho Boise Center<br>
	800 Park Boulevard, Suite 200<br>
	Boise, Idaho  83712<br>
	(208) 334-2999
	</font>
	</td>
</table>
<p>
<table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">
<tr>
<%
for col = 1 to 7
	response.write "<td width='14.285713%' align='center'>" _
		& "<font face='arial' size='1'>" _
		& WeekDayName(col,0) & "</font></td>"
next

response.write cal
%>
</table>
<font face="arial" size=2>
see this calendar online at http://boise.uidaho.edu
</font>
</center>
</body>
</html>