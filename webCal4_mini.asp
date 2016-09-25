<% Option Explicit %>
<% Response.Buffer = True %>
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 02/23/2001

' this is the mini calendar that pops up for date selection

dim m_strField		' field holding date
dim m_strForm		' form with field
dim m_strMonthList	' drop-down list
dim m_strYearList	' drop-down list
dim m_strJSDate		' format date according to LCID
dim m_strMonthGrid	' days of month in HTML
dim m_intMonth		' selected month
dim m_intYear		' selected year
dim x				' loop counter

' this opens a new browser so session state is lost
' restore location id from cookie
Session.LCID = Request.Cookies("LCID")
m_strField = Request.QueryString("field")
m_strForm = Request.QueryString("form")

' generate LCID-specific date format for JavaScript
if Session.LCID = 1033 OR Session.LCID = "" then
	m_strJSDate = "lMonth + '/1/' + lYear"
else
	m_strJSDate = "1/' + lMonth + '/' + lYear"
end if

' assign month and year
if Request.QueryString("date") <> "" then
	m_intMonth = Month(Request.QueryString("date"))
	m_intYear = Year(Request.QueryString("date"))
else
	m_intMonth = Month(now)
	m_intYear = Year(now)
end if

' build month drop-down
for x = 1 to 12
	m_strMonthList = m_strMonthList & "<option value=" & x
	if x = m_intMonth then m_strMonthList = m_strMonthList & " selected"
	m_strMonthList = m_strMonthList & ">" & MonthName(x,1) & VbCrLf
next

' build year drop-down
for x = year(Now) - 10 to year(Now) + 10
	m_strYearList = m_strYearList & "<option value=" & x
	if x = m_intYear then m_strYearList = m_strYearList & " selected"
	m_strYearList = m_strYearList & ">" & x & vbCrLf
next

' build days of month in HTML (updated 2/24/01)
' returns string ---------------------------------------------------------
Function makeMonthGrid(ByVal v_intMonth, ByVal v_intYear)
	dim intMonthNext
	dim intCol				' current column in grid
	dim intColFirst			' column of first day in month
	dim intDaysInMonth		' number of days in given month
	dim intDaysInPrevMonth	' number of days in last month
	dim strHTML
	dim intDay
	dim x

	intCol = 0
	intMonthNext = (v_intMonth + 1) Mod 12
	intColFirst = WeekDay(Dateserial(v_intYear, v_intMonth, 1))
	intDaysInMonth = Day(Dateserial(v_intYear, intMonthNext, 1) - 1)
	intDaysInPrevMonth = Day(Dateserial(v_intYear, v_intMonth, 1) - 1)

	' create the day of week headings
	for x = 1 to 7
		strHTML = strHTML & "<td width='14.3%' align='center' bgcolor='#404040'>" _
			& "<font face='Tahoma, Arial, Helvetica' size=1 color='#ffffff'>" _
			& "<b>" & Left(WeekDayName(x),1) & "</b></font></td>" & vbCrLf
	next
	strHTML = strHTML & "<tr>" & vbCrLf

	' cycle through all the days previous to the first
	' day of the active month
	for x = 1 to intColFirst - 1
		strHTML = strHTML & "<td align='right'>" _
			& "<font face='Tahoma, Arial, Helvetica' size=1 color='#777777'>" _
			& intDaysInPrevMonth - intColFirst + x + 1 & "</td>"
		intCol = intCol + 1
	next

	' cycle through all the days of the current month
	for intDay = 1 to intDaysInMonth
		intCol = intCol + 1
		strHTML = strHTML &  "<td align='right'"
		if v_intYear & v_intMonth & intDay = Year(now) & Month(now) & Day(now) then
			' highlight current day
			strHTML = strHTML &  " bgcolor='#e0e0e0'"
		elseif intCol = 1 or intCol = 7 then
			strHTML = strHTML &  " bgcolor='#999999'"
		end if
		strHTML = strHTML &  "><font face='Tahoma, Arial, Helvetica' size=1>" _
			& "<a href='#' name='day' onClick='newDay(""" _
			& DateSerial(v_intYear, v_intMonth, intDay) & """)' " _
			& "onMouseOver=""window.opener.status='Set to " & DateSerial(v_intYear, v_intMonth, intDay) _
			& "'; return true;"" onMouseOut=""window.opener.status='';"">" _
			& intDay & "</a></font></td>"
		if intCol = 7 AND intDay < intDaysInMonth then
			strHTML = strHTML &  "<tr>"
			intCol = 0
		end if
	next

	' cycle through as many days of the next month as
	' necessary to fill the calendar grid through column 7
	if intCol > 0 then
		intDay = 1
		do while intCol < 7
			strHTML = strHTML &  "<td align='right'><font face='Tahoma, Arial, Helvetica' size=1 color='#777777'>" _
				& intDay & "</font></td>"
			intDay = intDay + 1
			intCol = intCol + 1
		loop
	end if

	makeMonthGrid = strHTML
End Function
%>
<html>
<head>
<script language="javascript">
// update form field with selected date (updated 2/23/01)
// returns string and closes popup ---------------------------------------
function newDay(v_sDate) {
	var oOpenerForm = window.opener.document.<%=m_strForm%>;
	oOpenerForm.<%=m_strField%>.value = v_sDate;
	self.close();
}
// refresh page with selected month or year (updated 2/23/01)
// reposts page ----------------------------------------------------------
function newDate() {
	var oForm = document.frmCalPop;
	var lMonth = oForm.fldMonth.options[oForm.fldMonth.selectedIndex].value + '';
	var lYear = oForm.fldYear.options[oForm.fldYear.selectedIndex].value + '';
	var sDate = <%=m_strJSDate%>;
	location.href = "webCal4_mini.asp?field=<%=m_strField%>&form=<%=m_strForm%>&date=" + sDate;
}
</script>

<style type="text/css">
	A:hover { color:#000066; }
</style>

<title><%=MonthName(m_intMonth)%> &nbsp; &nbsp; &nbsp;</title>
</head>
<body bgcolor="#c0c0c0" link="#000044" vlink="#000044" alink="#000066" leftmargin=2 topmargin=2>

<center>
<table border=0 cellspacing=1 cellpadding=0>
<tr>
	<td colspan=12 align="center"><nobr>
	<form name="frmCalPop">
	<select name="fldMonth" onChange="newDate();"><%=m_strMonthList%></select>
	<select name='fldYear' onChange='newDate();'><%=m_strYearList%></select>
	</td>
	</form>
<tr height='14'>
	<%=makeMonthGrid(m_intMonth, m_intYear)%>
</table>
</center>
</body>
</html>