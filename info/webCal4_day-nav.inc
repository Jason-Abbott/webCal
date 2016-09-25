<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" cellspacing=0 cellpadding=2 border=0><tr><td>
<!-- end framing table -->

<table border=0 cellspacing=1 cellpadding=1 bgcolor="#<%=g_arColor(4)%>">
<tr height=14>
<%
dim arDays(31)
dim m
dim mNext
dim strFirst
dim strLast
dim strLastMonth
dim db
dim intCol
dim strNonDay
dim d
dim strThisDay
dim y
dim rsEvents

' determine how this page was called and assign values
' for month and year accordingly

m = Month(strThisDay)
y = Year(strThisDay)

' calculate the numeric value of the next month

if m < 12 then
	mNext = m + 1
else
	mNext = 1
end if

strFirstDate	= Dateserial(Year(strSelect), Month(strSelect), 1)
strNext			= DateAdd("m", 1, strFirstDate)
strPrev			= DateAdd("m", -1, strFirstDate)
strLastDate		= DateAdd("d", -1, strNext)


' get the first column of the first day

strFirst = WeekDay(Dateserial(y, m, 1))

' get the total days of the month by subtracting one
' day from the last day of next month

strLast = Day(Dateserial(y, mNext, 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

strLastMonth = Day(Dateserial(y, m, 1) - 1)

' if the user has logged in then use their query
' otherwise use default query (show all types)

if Session(unique & "Query") <> "" then
	strQuery = Session(unique & "Query")
else
	strQuery = ""
end if

strQuery = "SELECT event_date FROM tblEventDates " _
	& "WHERE event_date BETWEEN " & strDelim _
	& m & "/1/" & y & " 12:00:00 AM" & strDelim _
	& " AND " & strDelim & m & "/" & strLast & "/" & y _
	& " 11:59:59 PM" & strDelim

	Set rsEvents = Server.CreateObject("ADODB.RecordSet")
rsEvents.Open strQuery, g_strDSN, 3, 1, &H0001

do while not rsEvents.EOF
	arDays(Day(rsEvents("event_date"))) = True
	rsEvents.MoveNext
loop
rsEvents.Close
Set rsEvents = nothing

' day of week headings
for x = 1 to 7
	response.write "<td width='14.3%' align='center'>" _
		& "<font face='" & g_arFont(0) & "' size=1 color='#ffffff'>" _
		& "<b>" & Left(WeekDayName(x),1) & "</b></font></td>" & vbCrLf
next
response.write "</tr>"

' the column variable keeps constant track of the
' current calendar column

intCol = 0

strNonDay = "<td align=""right""><font face='" & g_arFont(0) & "' size=1 color=""#" _
		& g_arColor(11) & """>"



' cycle through all the days previous to the first
' day of the active month
		
for d = 1 to strFirst - 1
	response.write strNonDay & strLastMonth - strFirst + d + 1 & "</font></td>"
	intCol = intCol + 1
next

' now cycle through all the days of the current month

for d = 1 to strLast
	intCol = intCol + 1
	response.write "<td align=""right"" bgcolor=""#"
	if y & m & d = Year(now) & Month(now) & Day(now) then
		response.write g_arColor(8)
	elseif intCol = 1 OR intCol = 7 then
		response.write g_arColor(3)
	else
		response.write g_arColor(2)
	end if
	response.write """ ><font face='" & g_arFont(0) & "' size=1>"
	
	if arDays(d) then
		response.write "<a href=""webCal4_day.asp?date=" & DateSerial(y,m,d) _
			& """ " & showStatus("View " & DateSerial(y,m,d) & " details") & "><b>" _
			& d & "</b></a>"
	else
		response.write d
	end if
	
	response.write "</font></td>"
	
	if intCol = 7 AND d < strLast then
		response.write "<tr>"
		intCol = 0
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if intCol > 0 then
	d = 1
	do while intCol < 7
		response.write strNonDay & d & "</font></td>"
		d = d + 1
		intCol = intCol + 1
	loop
end if
%>

</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->
