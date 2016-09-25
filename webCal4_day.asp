<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./data/webCal4_cache.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./include/webCal4_grid-functions.inc"-->
<!--#include file="./include/webCal4_day-functions.inc"-->
<!--#include file="./language/webCal4_language.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/3/2001

dim m_strQuery			' query to send to database
dim m_strView			' current calendar view
dim m_strDescription	' event description shown in status bar
dim m_strTitle			' title at top of page
dim m_strHTML			' page display
dim m_strLoadFrom		' text to indicate page was cached
dim m_intStartTime		' time the query
dim m_arDates			' date parameters for this day
dim m_arSegments		' array of segment properties

m_strView = "day"
m_arDates = getDayDates()
m_intStartTime = milliTime()
m_arSegments = getGridProperties(Session(g_unique & "Segments")(g_DAY))
m_strQuery = makeQuery(m_arDates(g_THIS_DATE), m_arDates(g_THIS_DATE))
m_strHTML = readCache(m_strQuery, g_DAY)

if m_strHTML = "" then
	' page not found in cache--retrieve from database
	m_strHTML = makeDayHTML(m_strQuery)
	Call saveCache(m_strQuery, m_strHTML, m_arDates(g_THIS_DATE), m_arDates(g_THIS_DATE), g_DAY)
	m_strLoadFrom = "database"
Else
	m_strLoadFrom = "cache"
End if

' now generate the title to display at the top of the calendar
m_strTitle = "<font face='" & g_arFont(1) & "' size=5 color='#" _
	& g_arColor(4) & "'><b>Day " & DatePart("y", m_arDates(g_THIS_DATE)) _
	& " in " & Year(m_arDates(g_THIS_DATE)) & "</b></font>"
%>
<html>
<head>
<style><!--#include file="./style/webCal4_common.css"--></style>
<script language="javascript" SRC="./script/webCal4_buttons.js"></script>
<script language="javascript" SRC="./script/webCal4_functions.js"></script>
</head>
<body onLoad="showMessage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">

<!-- <%=m_strQuery%> -->

<table width="100%" border=0 cellspacing=5 cellpadding=1>
<tr><td><%=m_strTitle%></td>
<tr>
	<td width="90%" bgcolor="#<%=g_arColor(6)%>"><%=m_strHTML%></td>
	<% Response.Flush %>
	<td valign="top" align="center">
		<font face="Tahoma, Arial, Helvatica" color="#<%=g_arColor(4)%>">
		<b><font size=2><%=WeekdayName(WeekDay(m_arDates(g_THIS_DATE)))%></font><br>
		<font size=7><%=Day(m_arDates(g_THIS_DATE))%></font><br>
		<font size=5><a href="webCal4_month.asp?date=<%=m_arDates(g_THIS_DATE)%>"><%=MonthName(Month(m_arDates(g_THIS_DATE)),1)%></a></font></b><br>
		<font size=4><%=Year(m_arDates(g_THIS_DATE))%></font>
		</font>
		<p>
		<%=makeNavMonthHTML(m_arDates(g_THIS_DATE))%>
	</td>
<tr>
	<td align="center" bgcolor="#<%=g_arColor(5)%>">
	<!--include file="./include/webCal4_month-options.inc"-->
	</td>
</table>

<font face="<%=g_arFont(1)%>" size=1>
<%=showLoadTime(m_strQuery, m_strLoadFrom)%><br>
<a href="http://webott.com/jason/webCal.html" target="_top">
webCal 4.0</a> 
</font>

</body>
</html>